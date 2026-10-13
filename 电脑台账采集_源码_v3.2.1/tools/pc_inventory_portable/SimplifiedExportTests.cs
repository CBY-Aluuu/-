using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using PcInventory;

// Synthetic fixtures only. This suite never reads the user's device records.
public static class SimplifiedExportTests
{
    private static int checks;
    private static void Assert(bool condition, string text)
    {
        if (!condition) throw new Exception("FAILED: " + text);
        checks++;
        Console.WriteLine("PASS: " + text);
    }

    private static InventoryRecord Fixture(int version)
    {
        return new InventoryRecord {
            HardwareDataVersion = version, DeviceKey = "SYNTHETIC-SAME-DEVICE", RecordId = "synthetic-record",
            DeviceNumber = "000017", DeviceType = "台式机", ManufacturerModel = "样例品牌 A&B <Model>",
            BiosSerial = "00001234", DiskSerials = "硬盘0 — SAMPLE NVME 1TB — 0000_1234_ABCD\n硬盘1 — SAMPLE SATA — 00007788\n硬盘2 — SAMPLE SSD — 00008899",
            MacAddresses = "10:22:33:44:55:66\n20:33:44:55:66:77", IPAddresses = "192.0.2.155",
            AdapterPairs = "SYNTHETIC-ADAPTER-PAIR 198.51.100.77", Cpu = "样例处理器 8核", Memory = "16 GB",
            OperatingSystem = "Windows 示例系统", Site = "示例站", Location = "示例办公室", AssignedTo = "示例人员",
            DeviceStatus = "在用", AcquiredDate = "2024-05-06", Remarks = "=HYPERLINK(\"https://invalid.example\",\"示例文本\")",
            CollectedAt = "2026-10-12T09:30:15+08:00", Verification = "无"
        };
    }

    public static int Main(string[] args)
    {
        try
        {
            if (args.Length != 1) throw new Exception("Specify a synthetic QA output directory.");
            string root = Path.GetFullPath(args[0]);
            Directory.CreateDirectory(root);
            VerifyPresentation();
            VerifyExport(root);
            Console.WriteLine("ALL " + checks + " SIMPLIFIED EXPORT CHECKS PASSED");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static void VerifyPresentation()
    {
        InventoryRecord latest = Fixture(3);
        Assert(InventoryPresentation.LocationText(latest) == "示例站 / 示例办公室", "legacy site and room combine into one location without changing stored fields");
        InventoryRecord locationOnly = new InventoryRecord { Location = "  三楼办公室  " };
        Assert(InventoryPresentation.LocationText(locationOnly) == "三楼办公室", "new record needs only location");
        locationOnly.Site = "三楼办公室";
        Assert(InventoryPresentation.LocationText(locationOnly) == "三楼办公室", "identical legacy site and room are shown once");
        locationOnly.Location = " ";
        Assert(InventoryPresentation.LocationText(locationOnly) == "三楼办公室", "legacy site-only record retains its location");
        Assert(InventoryPresentation.LocationText(new InventoryRecord()) == "", "unknown location is not invented");
        Assert(InventoryPresentation.DiskText(latest) == latest.DiskSerials, "new disk lines preserve model and identifier only");
        Assert(InventoryPresentation.MacText(latest) == latest.MacAddresses, "new physical MAC text contains plain addresses");
        Assert(InventoryPresentation.VerificationText(latest) == "无", "complete record has short verification text");
        latest.Verification = "";
        Assert(InventoryPresentation.VerificationText(latest) == "无", "empty new verification displays none");
        latest.Verification = "整机编码待核对；取得日期待核对";
        Assert(InventoryPresentation.VerificationText(latest) == latest.Verification, "specific new verification items retained");
        latest.DiskSerials = "硬盘0 — 示例盘 — 待核对";
        Assert(InventoryPresentation.DiskText(latest) == latest.DiskSerials, "missing disk serial keeps the available model");

        InventoryRecord old = Fixture(0);
        old.DiskSerials = "物理硬盘2 [SAMPLE NVME 1TB]: 0000_0000_ABCD. (1024.2 GB) | 物理硬盘0 [SAMPLE SATA]: 00001234 (512.11 GB)";
        Assert(InventoryPresentation.DiskText(old) == "硬盘0 — SAMPLE SATA — 00001234\n硬盘2 — SAMPLE NVME 1TB — 0000_0000_ABCD.", "legacy disks sort by disk number and remove duplicate capacity detail");
        Assert(!InventoryPresentation.DiskText(old).Contains("|") && !InventoryPresentation.DiskText(old).Contains("GB"), "disk summary replaces pipe blocks with short separate lines");
        old.DiskSerials = "0000001234";
        Assert(InventoryPresentation.DiskText(old).EndsWith("0000001234"), "legacy bare disk identifier retains leading zeros");
        old.DiskSerials = "物理硬盘0 [SAMPLE DISK]: 未读取到，待核实 (512 GB)";
        Assert(InventoryPresentation.DiskText(old) == "硬盘0 — SAMPLE DISK — 待核对", "unread legacy serial shortened without dropping known model");
        old.MacAddresses = "网卡0 [Realtek Ethernet]（以太网）: 10:22:33:44:55:66 | 网卡13 [VeryKuai TAP Adapter]（以太网4）: 00:FF:04:28:35:E2 | 网卡4 [Intel Wi-Fi]（无线）: 20-33-44-55-66-77 | 网卡5 [Intel Wi-Fi]（无线）: 20:33:44:55:66:77";
        Assert(InventoryPresentation.MacText(old) == "10:22:33:44:55:66\n20:33:44:55:66:77", "legacy TAP omitted and physical candidates normalized without names");
        Assert(InventoryPresentation.VerificationText(old).Contains("旧MAC请重采"), "legacy current MACs are not silently claimed permanent physical MACs");
        old.MacAddresses = "虚拟 TAP: 00:FF:04:28:35:E2";
        Assert(InventoryPresentation.MacText(old) == "待核对" && InventoryPresentation.VerificationText(old).Contains("物理MAC待核对"), "virtual-only old record visibly lacks physical MAC");
        latest.MacAddresses = "00:00:00:00:00:00\nFF:FF:FF:FF:FF:FF\n01:22:33:44:55:66\n10-22-33-44-55-66\n10:22:33:44:55:66";
        Assert(InventoryPresentation.MacText(latest) == "10:22:33:44:55:66", "MAC text excludes zero, broadcast and multicast and deduplicates");
        latest.MacAddresses = "02:22:33:44:55:66";
        Assert(InventoryPresentation.MacText(latest) == latest.MacAddresses, "locally administered unicast hardware MAC is not automatically discarded");
        Assert(InventoryPresentation.CollectedTime("2026-10-12T09:30:15+08:00") == "2026-10-12 09:30", "collection display omits seconds and timezone suffix");
        Assert(InventoryPresentation.SerialText("00001234") == "00001234" && InventoryPresentation.SerialText("未读取到，待核实") == "待核对", "serial label preserves leading zeros and shortens missing state");
    }

    private static void VerifyExport(string root)
    {
        string folder = Path.Combine(root, "synthetic-ledger");
        InventoryStore store = new InventoryStore(folder);
        InventoryRecord first = Fixture(3);
        SaveResult firstSave = store.Save(first);
        InventoryRecord second = Fixture(3);
        second.DeviceNumber = "000018";
        second.Site = "";
        second.DiskSerials = "硬盘0 — 示例长型号 NVME SOLID STATE DISK 2048GB — 0000_1111_2222_3333_4444_5555_6666_7777";
        second.MacAddresses = "30:44:55:66:77:88";
        second.Remarks = "";
        second.AcquiredDate = "";
        second.Verification = "取得日期待核对";
        store.Save(second);
        string recordFolder = Path.Combine(folder, "数据", "记录");
        string legacyPath = Path.Combine(recordFolder, "00000000000000000003_0123456789abcdef0123456789abcdef.json");
        InventoryRecord legacy = Fixture(0);
        legacy.DiskSerials = "物理硬盘2 [SAMPLE NVME 1TB]: 0000_0000_ABCD. (1024.2 GB) | 物理硬盘0 [SAMPLE SATA]: 00001234 (512.11 GB)";
        legacy.MacAddresses = "网卡0 [Realtek Ethernet]（以太网）: 10:22:33:44:55:66 | 网卡13 [VeryKuai TAP Adapter]（以太网4）: 00:FF:04:28:35:E2";
        legacy.Verification = "旧版很长的常规说明，包括IP配对规则与无关步骤。";
        using (MemoryStream stream = new MemoryStream())
        {
            new DataContractJsonSerializer(typeof(InventoryRecord)).WriteObject(stream, legacy);
            File.WriteAllText(legacyPath, "{\"FormatVersion\":1,\"SaveSequence\":3,\"Record\":" + Encoding.UTF8.GetString(stream.ToArray()) + "}", new UTF8Encoding(false));
        }
        string before = Convert.ToBase64String(File.ReadAllBytes(legacyPath));
        store.RebuildSummary();
        Assert(before == Convert.ToBase64String(File.ReadAllBytes(legacyPath)), "rebuilding new layout never rewrites old JSON");
        Assert(store.LoadRecords().Count == 3 && store.LoadRecords()[2].DiskSerials == legacy.DiskSerials, "legacy disk details remain available unchanged in raw records");
        XmlDocument sheet = ReadXml(store.SummaryPath, "xl/worksheets/sheet1.xml");
        XmlNamespaceManager ns = Namespaces(sheet);
        string[] expected = { "序号", "设备编号", "设备类型", "品牌及型号", "整机序列号", "硬盘（型号 / 编码）", "物理MAC", "CPU", "内存", "操作系统", "位置", "使用人员", "设备状态", "取得日期", "数量", "备注", "采集时间", "待核对" };
        string[] actual = sheet.SelectNodes("//s:sheetData/s:row[@r='1']/s:c/s:is/s:t", ns).Cast<XmlNode>().Select(n => n.InnerText).ToArray();
        Assert(actual.Take(18).SequenceEqual(expected) && actual.Length == 19 && actual[18] == "记录标识（勿修改）", "18 visible columns contain one location column, followed only by record identity");
        Assert(sheet.SelectSingleNode("//s:cols/s:col[@min='19' and @max='19']/@hidden", ns).Value == "1", "internal record identity is hidden from ordinary spreadsheet view");
        Assert(Cell(sheet, ns, "K2") == "示例站 / 示例办公室" && Cell(sheet, ns, "K3") == "示例办公室", "one location column preserves legacy parts and supports a new location-only record");
        WorkbookSnapshot roundTrip = InventoryWorkbookReader.ReadSnapshot(store.SummaryPath);
        Assert(roundTrip.Records[0].Site == null && roundTrip.Records[0].Location == "示例站 / 示例办公室" && roundTrip.Records[1].Location == "示例办公室", "new single-location export imports without duplicating the old site");
        Assert(store.SynchronizeSummary().SummaryReadable, "single-location workbook synchronizes against untouched legacy JSON");
        Assert(!actual.Any(h => h.Contains("IP") || h.Contains("配对")), "export has no IP address or pairing columns");
        Assert(sheet.SelectNodes("//s:sheetData/s:row", ns).Count == 4, "all three saves including same device are still separate rows");
        Assert(Cell(sheet, ns, "B2") == "000017" && Cell(sheet, ns, "E2") == "00001234", "device and BIOS serials preserve leading zeros");
        Assert(Cell(sheet, ns, "D2") == first.ManufacturerModel, "model XML special characters retain exact value");
        Assert(Cell(sheet, ns, "F2") == first.DiskSerials && Cell(sheet, ns, "G2") == first.MacAddresses, "new disk and MAC cell lines remain simple");
        Assert(Cell(sheet, ns, "F4") == InventoryPresentation.DiskText(legacy) && Cell(sheet, ns, "G4") == "10:22:33:44:55:66", "legacy export is cleaned while underlying record stays unchanged");
        Assert(!sheet.OuterXml.Contains("192.0.2.155") && !sheet.OuterXml.Contains("198.51.100.77") && !sheet.OuterXml.Contains("SYNTHETIC-ADAPTER-PAIR") && !sheet.OuterXml.Contains("00:FF:04:28:35:E2"), "IP and virtual TAP values cannot leak through exported hidden columns");
        Assert(Cell(sheet, ns, "P2") == first.Remarks && sheet.SelectNodes("//s:f", ns).Count == 0, "formula-like note remains literal text");
        Assert(sheet.SelectSingleNode("//s:c[@r='O2']/s:v", ns).InnerText == "1" && sheet.SelectSingleNode("//s:c[@r='A2']/s:v", ns).InnerText == "1", "quantity and row number remain numeric");
        Assert(Cell(sheet, ns, "N3") == "", "unknown acquired date remains blank");
        Assert(Cell(sheet, ns, "Q2") == "2026-10-12 09:30" && Cell(sheet, ns, "R2") == "无", "collection date and verification are concise");
        Assert(Cell(sheet, ns, "R4").Contains("旧MAC请重采") && !Cell(sheet, ns, "R4").Contains("IP配对"), "legacy verification becomes short actionable guidance");
        Assert(sheet.SelectSingleNode("//s:autoFilter/@ref", ns).Value == "A1:S4", "sorting includes hidden record identity with the 18 visible columns");
        Assert(sheet.SelectSingleNode("//s:pane/@state", ns).Value == "frozen", "column header stays frozen");
        XmlDocument styles = ReadXml(store.SummaryPath, "xl/styles.xml");
        XmlNamespaceManager styleNs = Namespaces(styles);
        Assert(styles.SelectNodes("//s:cellXfs/s:xf/s:alignment[@wrapText='1']", styleNs).Count > 0, "multiline disk and MAC cells have wrapping support");
        string backup = store.BackupNow(Path.Combine(root, "manual-backups"));
        string restored = store.RestoreBackupToNewFolder(backup, Path.Combine(root, "restored"));
        Assert(new InventoryStore(restored).LoadRecords().Count == 3, "backup and restore preserve every row with the new presentation");
        string restoredLegacy = Path.Combine(restored, "数据", "记录", Path.GetFileName(legacyPath));
        Assert(Convert.ToBase64String(File.ReadAllBytes(restoredLegacy)) == before, "legacy raw JSON remains byte-identical after backup and restore");
        Assert(ReadXml(new InventoryStore(restored).SummaryPath, "xl/worksheets/sheet1.xml").OuterXml == sheet.OuterXml, "restoring old and new records rebuilds the same simplified worksheet");
        File.WriteAllText(Path.Combine(root, "synthetic-workbook-path.txt"), store.SummaryPath, new UTF8Encoding(false));
    }

    private static XmlNamespaceManager Namespaces(XmlDocument xml)
    {
        XmlNamespaceManager ns = new XmlNamespaceManager(xml.NameTable);
        ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        return ns;
    }
    private static string Cell(XmlDocument xml, XmlNamespaceManager ns, string address)
    {
        XmlNode text = xml.SelectSingleNode("//s:c[@r='" + address + "']/s:is/s:t", ns);
        return text == null ? "" : text.InnerText.Replace("\r", "");
    }
    private static XmlDocument ReadXml(string path, string entry)
    {
        using (FileStream input = File.OpenRead(path))
        using (ZipArchive archive = new ZipArchive(input, ZipArchiveMode.Read))
        using (Stream stream = archive.GetEntry(entry).Open())
        { XmlDocument xml = new XmlDocument(); xml.Load(stream); return xml; }
    }
}
