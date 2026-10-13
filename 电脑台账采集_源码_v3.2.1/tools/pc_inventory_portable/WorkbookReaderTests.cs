using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;
using System.Text;
using PcInventory;

public static class WorkbookReaderTests
{
    private static int checks, serial;
    private static string root;
    private static string[] headers = { "序号", "设备编号", "设备类型", "品牌及型号", "整机序列号", "硬盘（型号 / 编码）", "物理MAC", "CPU", "内存", "操作系统", "所属站点", "所在位置", "使用人员", "设备状态", "取得日期", "数量", "备注", "采集时间", "待核对" };
    private const string Id = "00000000000000000001_0123456789abcdef0123456789abcdef.json";
    private const string Id2 = "00000000000000000002_0123456789abcdef0123456789abcdef.json";
    private static void Check(bool ok, string message) { if (!ok) throw new Exception("FAILED: " + message); checks++; Console.WriteLine("PASS: " + message); }
    private static string Esc(string value) { return SecurityElement.Escape(value ?? ""); }
    private static string Inline(string address, string value) { return "<c r=\"" + address + "\" t=\"inlineStr\"><is><t xml:space=\"preserve\">" + Esc(value) + "</t></is></c>"; }
    private static string Row(int number, params string[] values) { return "<row r=\"" + number + "\">" + String.Join("", values.Select((s, i) => Inline(((char)('A' + i)).ToString() + number, s))) + "</row>"; }
    private static string Sheet(string rows) { return "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" + rows + "</sheetData></worksheet>"; }
    private static string[] Data() { return new[] { "1", "000017", "台式机", "示例品牌 & 型号", "00001234", "硬盘0 — SAMPLE — 00000077\n硬盘1 — SAMPLE B — 00000088", "10:22:33:44:55:66", "CPU", "16 GB", "示例系统", "示例站", "房间", "示例人员", "在用", "2024-05", "1", "=1+1", "2026-10-12 10:15", "无" }; }
    private static string Make(string rows, string index = null, string shared = null, string styles = null, bool date1904 = false, string target = "worksheets/sheet1.xml", string targetMode = null, string extra = null)
    {
        string path = Path.Combine(root, (++serial).ToString("D3") + ".xlsx");
        using (FileStream file = File.Create(path))
        using (ZipArchive zip = new ZipArchive(file, ZipArchiveMode.Create))
        {
            Part(zip, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>");
            Part(zip, "xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr date1904=\"" + (date1904 ? "1" : "0") + "\"/><sheets><sheet name=\"电脑台账\" sheetId=\"1\" r:id=\"r1\"/>" + (index == null ? "" : "<sheet name=\"记录同步索引\" sheetId=\"2\" state=\"veryHidden\" r:id=\"r2\"/>") + "</sheets></workbook>");
            Part(zip, "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"" + Esc(target) + "\"" + (targetMode == null ? "" : " TargetMode=\"" + targetMode + "\"") + "/>" + (index == null ? "" : "<Relationship Id=\"r2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>") + "</Relationships>");
            Part(zip, "xl/worksheets/sheet1.xml", Sheet(rows));
            if (index != null) Part(zip, "xl/worksheets/sheet2.xml", Sheet(index));
            if (shared != null) Part(zip, "xl/sharedStrings.xml", shared);
            if (styles != null) Part(zip, "xl/styles.xml", styles);
            if (extra != null) Part(zip, extra, "irrelevant");
        }
        return path;
    }
    private static void Part(ZipArchive zip, string name, string text) { using (StreamWriter writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false))) writer.Write(text); }
    private static void Reject(string path, string label) { bool rejected = false; try { InventoryWorkbookReader.ReadSnapshot(path); } catch (InvalidDataException) { rejected = true; } Check(rejected, label); }
    public static int Main(string[] args)
    {
        try
        {
            root = Path.GetFullPath(args[0]); Directory.CreateDirectory(root);
            string basic = Make(Row(1, headers) + Row(2, Data()));
            byte[] original = File.ReadAllBytes(basic);
            WorkbookSnapshot read = InventoryWorkbookReader.ReadSnapshot(basic);
            Check(read.Records.Count == 1 && read.Records[0].DeviceNumber == "000017" && read.Records[0].BiosSerial == "00001234", "inline text and identifying leading zeros are preserved");
            Check(read.Records[0].DiskSerials.Contains("\n") && read.Records[0].Remarks == "=1+1", "multiline disk text and literal formula-like note preserved without execution");
            Check(read.RawRows[0].SequenceEqual(Data()) && read.Headers.SequenceEqual(headers), "raw visible row cells and headers retained for strict comparison");
            Check(read.MappedFields[1] == "DeviceNumber" && read.MappedFields[4] == "BiosSerial" && read.MappedFields[18] == "Verification", "zero-based mapped fields are available without presentation roundtrip");
            Check(read.IsLegacy && !read.HasSyncIdentity && read.StorageIds[0] == "" && read.Records[0].HardwareDataVersion == 3, "new visible schema without sync metadata remains legacy for synchronization");
            Check(read.Sha256.Length == 64 && original.SequenceEqual(File.ReadAllBytes(basic)), "reader hashes workbook without modifying any bytes");
            Check(InventoryWorkbookReader.Read(basic).Count == 1, "simple import API returns records");
            Check(read.Records[0].Site == "示例站" && read.Records[0].Location == "房间", "prior two-location headers retain both original fields");
            string[] oneLocationHeader = headers.Take(10).Concat(new[] { "位置" }).Concat(headers.Skip(12)).ToArray();
            string[] oneLocationData = Data().Take(10).Concat(new[] { "示例站 / 房间" }).Concat(Data().Skip(12)).ToArray();
            read = InventoryWorkbookReader.ReadSnapshot(Make(Row(1, oneLocationHeader) + Row(2, oneLocationData)));
            Check(read.Records[0].Site == null && read.Records[0].Location == "示例站 / 房间" && read.Records[0].AssignedTo == "示例人员", "single location header imports location without shifting remaining fields");

            string[] oldHeader = { "序号", "设备编号", "设备类型", "品牌及型号", "使用人员", "设备状态", "硬盘序列号", "网卡MAC", "取得日期", "备注" };
            string oldData = "<row r=\"4\">" + Inline("A4", "1") + "<c r=\"B4\" t=\"s\"><v>0</v></c>" + Inline("C4", "台式机") + Inline("D4", "样例型号") + Inline("E4", "示例人员") + Inline("H4", "10:22:33:44:55:66") + "</row>";
            read = InventoryWorkbookReader.ReadSnapshot(Make(Row(1, "标题") + Row(3, oldHeader) + oldData, null, "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><si><r><t>0000</t></r><r><t>17</t></r><rPh sb=\"0\" eb=\"1\"><t>phonetic ignored</t></rPh></si></sst>"));
            Check(read.Records.Count == 1 && read.Records[0].DeviceNumber == "000017" && read.OriginalRowNumbers[0] == 4, "legacy title rows and rich shared strings mapped correctly");
            Check(read.Records[0].HardwareDataVersion == 0 && read.Records[0].Site == null, "legacy import does not invent physical-MAC proof or missing site");

            string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"3\"><numFmt numFmtId=\"164\" formatCode=\"yyyy-mm\"/><numFmt numFmtId=\"165\" formatCode=\"000000\"/><numFmt numFmtId=\"166\" formatCode=\"yyyy\"/></numFmts><cellXfs count=\"5\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/><xf numFmtId=\"164\"/><xf numFmtId=\"165\"/><xf numFmtId=\"166\"/></cellXfs></styleSheet>";
            string numeric = "<row r=\"2\"><c r=\"B2\" s=\"3\"><v>17</v></c>" + Inline("C2", "台式机") + "<c r=\"O2\" s=\"1\"><v>45292</v></c></row>";
            read = InventoryWorkbookReader.ReadSnapshot(Make(Row(1, headers) + numeric, null, null, styles));
            Check(read.Records[0].DeviceNumber == "000017" && read.Records[0].AcquiredDate == "2024-01-01", "formatted numeric identifier and Excel serial date import correctly");
            read = InventoryWorkbookReader.ReadSnapshot(Make(Row(1, headers) + numeric.Replace("s=\"1\"", "s=\"2\""), null, null, styles));
            Check(read.Records[0].AcquiredDate == "2024-01", "month-only date style preserves month precision");
            read = InventoryWorkbookReader.ReadSnapshot(Make(Row(1, headers) + numeric.Replace("s=\"1\"", "s=\"4\""), null, null, styles));
            Check(read.Records[0].AcquiredDate == "2024", "year-only date style preserves year precision");
            read = InventoryWorkbookReader.ReadSnapshot(Make(Row(1, headers) + numeric.Replace("45292", "0"), null, null, styles, true));
            Check(read.Records[0].AcquiredDate == "1904-01-01", "1904 workbook date system is supported");
            read = InventoryWorkbookReader.ReadSnapshot(Make(Row(1, headers) + numeric.Replace(" s=\"1\"><v>45292", " t=\"d\"><v>2025-02-03"), null, null, styles));
            Check(read.Records[0].AcquiredDate == "2025-02-03", "ISO typed date is supported");
            read = InventoryWorkbookReader.ReadSnapshot(Make(Row(1, headers) + Row(2, Data()) + "<row r=\"3\"><c r=\"A3\"><v>2</v></c></row><row r=\"1048576\"><c r=\"A1048576\" s=\"0\"/></row>"));
            Check(read.Records.Count == 1, "sequence-only and formatted blank tail rows do not become devices");
            Check(InventoryWorkbookReader.Read(Make(Row(1, headers))).Count == 0, "recognized empty ledger is allowed");

            string[] syncHeader = headers.Concat(new[] { "记录标识（勿修改）" }).ToArray();
            string index = Row(1, "PC-INVENTORY-SYNC-1") + Row(2, Id) + Row(3, Id2);
            read = InventoryWorkbookReader.ReadSnapshot(Make(Row(1, syncHeader) + Row(2, Data().Concat(new[] { Id }).ToArray()), index));
            Check(read.HasSyncIdentity && !read.IsLegacy && read.StorageIds.SequenceEqual(new[] { Id }) && read.ExportedIds.SequenceEqual(new[] { Id, Id2 }), "sync snapshot distinguishes remaining rows from exported baseline IDs");
            Check(read.RawRows[0].Length == 20 && read.RawRows[0][19] == Id, "hidden identity column remains available in raw rows");
            Reject(Make(Row(1, syncHeader) + Row(2, Data())), "identity header without index is rejected");
            Reject(Make(Row(1, headers) + Row(2, Data()), index), "index without identity header is rejected");
            Reject(Make(Row(1, syncHeader), Row(1, "BAD-MARKER")), "damaged sync marker cannot fall back to legacy");
            Reject(Make(Row(1, syncHeader.Concat(new[] { "记录标识（勿修改）" }).ToArray()), index), "duplicate identity column cannot masquerade as valid sync metadata");
            Reject(Make(Row(1, syncHeader), Row(1, "PC-INVENTORY-SYNC-1") + "<row r=\"2\">" + Inline("A2", Id) + "<c r=\"B2\" t=\"e\"><v>#REF!</v></c></row>"), "error in sync index metadata is rejected");
            Reject(Make(Row(1, syncHeader), Row(1, "PC-INVENTORY-SYNC-1") + Row(2, Id) + Row(3, Id)), "duplicate baseline IDs are rejected");
            Reject(Make(Row(1, syncHeader), Row(1, "PC-INVENTORY-SYNC-1") + Row(2, "arbitrary-id")), "baseline IDs must be valid original record filenames");
            Reject(Make(Row(1, syncHeader) + Row(2, Data().Concat(new[] { "arbitrary-id" }).ToArray()), index), "visible row IDs must be valid original record filenames");
            Reject(Make(Row(1, syncHeader) + Row(2, Enumerable.Repeat("", 19).Concat(new[] { Id }).ToArray()), index), "identity-only empty row cannot silently become deletion");
            Reject(Make(Row(1, new[] { "姓名", "地址", "备注" }) + Row(2, "sample")), "unrelated spreadsheet is rejected");
            Reject(Make(Row(1, headers.Concat(new[] { "资产编号" }).ToArray())), "duplicate logical columns are rejected");
            Reject(Make(Row(1, headers) + "<row r=\"2\">" + Inline("Q2", "仅备注的尾注") + "</row>"), "notes-only footer is rejected rather than imported as a device");
            Reject(Make(Row(1, headers) + "<row r=\"2\"><c r=\"B2\" t=\"s\"><v>5</v></c></row>"), "invalid shared string index is rejected");
            Reject(Make(Row(1, headers) + "<row r=\"2\"><c r=\"B2\" s=\"999\"><v>1</v></c></row>"), "invalid numeric style is rejected");
            Reject(Make(Row(1, headers) + "<row r=\"2\"><c r=\"B2\"><f>1+1</f></c></row>"), "uncalculated formulas cannot silently import empty values");
            Reject(Make(Row(1, headers) + "<row r=\"2\"><c r=\"B2\" t=\"e\"><v>#VALUE!</v></c></row>"), "Excel errors in mapped fields are rejected");
            Reject(Make(Row(1, headers) + "<row r=\"2\">" + Inline("B2", "A") + Inline("B2", "B") + "</row>"), "duplicate cell coordinates are rejected");
            Reject(Make(Row(1, headers), null, null, null, false, "../../outside.xml"), "relationship traversal outside workbook is rejected");
            Reject(Make(Row(1, headers), null, null, null, false, "https://invalid.example/sheet.xml", "External"), "external worksheet targets are rejected without requests");
            Reject(Make(Row(1, headers), null, null, null, false, "worksheets/sheet1.xml", null, "../outside.xml"), "ZIP traversal names are rejected without extraction");
            string dtd = Make(Row(1, headers));
            using (FileStream file = new FileStream(dtd, FileMode.Open, FileAccess.ReadWrite))
            using (ZipArchive zip = new ZipArchive(file, ZipArchiveMode.Update))
            { zip.GetEntry("xl/worksheets/sheet1.xml").Delete(); Part(zip, "xl/worksheets/sheet1.xml", "<!DOCTYPE worksheet [<!ENTITY x SYSTEM 'file:///C:/forbidden'>]>" + Sheet(Row(1, headers))); }
            Reject(dtd, "DTD and external entities are forbidden");
            string duplicate = Make(Row(1, headers));
            using (FileStream file = new FileStream(duplicate, FileMode.Open, FileAccess.ReadWrite))
            using (ZipArchive zip = new ZipArchive(file, ZipArchiveMode.Update)) Part(zip, "xl/worksheets/sheet1.xml", Sheet(Row(1, headers)));
            Reject(duplicate, "duplicate ZIP parts are rejected");
            Reject(@"\\not-a-server\share\test.xlsx", "network path is rejected before file access");
            string manyRows = Row(1, headers) + String.Join("", Enumerable.Range(2, 20001).Select(i => "<row r=\"" + i + "\">" + Inline("B" + i, "SYNTHETIC") + "</row>"));
            Reject(Make(manyRows), "record count limit is enforced");
            Console.WriteLine("ALL " + checks + " WORKBOOK READER CHECKS PASSED");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
