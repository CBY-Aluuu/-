using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace PcInventory
{
    public sealed class WorkbookSnapshot
    {
        public List<InventoryRecord> Records { get; set; }
        public List<string> StorageIds { get; set; }
        public List<string> ExportedIds { get; set; }
        public string[] Headers { get; set; }
        public Dictionary<int, string> MappedFields { get; set; }
        public List<string[]> RawRows { get; set; }
        public List<int> OriginalRowNumbers { get; set; }
        public bool HasSyncIdentity { get; set; }
        public bool IsLegacy { get; set; }
        public string Sha256 { get; set; }
    }

    // Local OOXML reader only: never runs formulas, follows external links, or changes a workbook.
    public static class InventoryWorkbookReader
    {
        private const long MaxPackageBytes = 512L * 1024 * 1024;
        private const long MaxPartBytes = 128L * 1024 * 1024;
        private const int MaxRecords = 20000;
        private const string IdentityHeader = "记录标识（勿修改）";
        private const string IndexName = "记录同步索引";
        private const string IndexMarker = "PC-INVENTORY-SYNC-1";
        private static readonly Dictionary<string, string> FieldNames = CreateFieldNames();

        private sealed class Cell
        {
            public string Value;
            public string Problem;
        }
        private sealed class Row
        {
            public int Number;
            public SortedDictionary<int, Cell> Cells = new SortedDictionary<int, Cell>();
        }
        private sealed class Sheet
        {
            public string Name;
            public string Part;
            public bool Hidden;
            public List<Row> Rows;
            public int HeaderIndex = -1;
            public Dictionary<int, string> Fields;
        }

        public static List<InventoryRecord> Read(string path) { return ReadSnapshot(path).Records; }

        public static WorkbookSnapshot ReadSnapshot(string path)
        {
            byte[] bytes = ReadStableLocalFile(path);
            string hash;
            using (SHA256 sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
            try
            {
                using (MemoryStream input = new MemoryStream(bytes, false))
                using (ZipArchive archive = new ZipArchive(input, ZipArchiveMode.Read))
                {
                    var entries = ValidateArchive(archive);
                    XmlDocument workbook = XmlPart(entries, "xl/workbook.xml", true);
                    XmlDocument relationships = XmlPart(entries, "xl/_rels/workbook.xml.rels", true);
                    var parts = ReadRelationships(relationships);
                    var strings = ReadSharedStrings(entries);
                    var styles = ReadStyles(entries);
                    XmlNode properties = workbook.SelectSingleNode("/*[local-name()='workbook']/*[local-name()='workbookPr']");
                    bool date1904 = properties != null && (Attribute(properties, "date1904") == "1" || Attribute(properties, "date1904") == "true");
                    var sheets = new List<Sheet>();
                    foreach (XmlNode element in workbook.SelectNodes("/*[local-name()='workbook']/*[local-name()='sheets']/*[local-name()='sheet']"))
                    {
                        string id = null;
                        foreach (XmlAttribute attribute in element.Attributes)
                            if (attribute.LocalName == "id" && attribute.NamespaceURI.Length > 0) id = attribute.Value;
                        string target;
                        if (id == null || !parts.TryGetValue(id, out target)) throw Invalid("工作表链接不完整。");
                        var sheet = new Sheet { Name = Attribute(element, "name"), Part = target, Hidden = Attribute(element, "state") == "hidden" || Attribute(element, "state") == "veryHidden" };
                        if (sheets.Any(s => s.Name == sheet.Name)) throw Invalid("工作表名称重复。");
                        sheets.Add(sheet);
                    }
                    if (sheets.Count == 0 || sheets.Count > 64) throw Invalid("工作表数量不符合要求。");
                    var matches = new List<Sheet>();
                    foreach (Sheet sheet in sheets.Where(s => s.Name != IndexName && !s.Hidden))
                    {
                        sheet.Rows = ReadRows(XmlPart(entries, sheet.Part, true), strings, styles, date1904);
                        IdentifyHeader(sheet);
                        if (sheet.HeaderIndex >= 0) matches.Add(sheet);
                    }
                    Sheet selected = matches.FirstOrDefault(s => s.Name == "电脑台账");
                    if (selected == null)
                    {
                        if (matches.Count != 1) throw Invalid(matches.Count == 0 ? "没有找到电脑台账表头，请选择原台账 .xlsx 文件。" : "文件内有多个可识别台账，请每个文件只保留一份需要导入的台账。");
                        selected = matches[0];
                    }
                    var result = BuildSnapshot(selected);
                    Sheet sync = sheets.FirstOrDefault(s => s.Name == IndexName);
                    bool identity = result.Headers.Any(h => NormalizeHeader(h) == NormalizeHeader(IdentityHeader));
                    if (identity != (sync != null)) throw Invalid("记录标识或同步索引缺失，请使用完整的台账文件。");
                    if (sync != null)
                    {
                        sync.Rows = ReadRows(XmlPart(entries, sync.Part, true), strings, styles, date1904);
                        ReadIndex(sync.Rows, result.ExportedIds);
                        foreach (string id in result.StorageIds) ValidateIdentity(id);
                    }
                    result.HasSyncIdentity = identity;
                    result.IsLegacy = !identity;
                    result.Sha256 = hash;
                    return result;
                }
            }
            catch (InvalidDataException) { throw; }
            catch (XmlException) { throw Invalid("台账 XML 损坏或含有不支持的内容。"); }
            catch (ArgumentException) { throw Invalid("台账结构损坏或格式不支持。"); }
            catch (OverflowException) { throw Invalid("台账中的数值超出支持范围。"); }
        }

        private static byte[] ReadStableLocalFile(string path)
        {
            if (String.IsNullOrWhiteSpace(path)) throw Invalid("请选择本地 .xlsx 台账文件。");
            string full = Path.GetFullPath(path);
            if (full.StartsWith(@"\\", StringComparison.Ordinal) || !String.Equals(Path.GetExtension(full), ".xlsx", StringComparison.OrdinalIgnoreCase)) throw Invalid("请选择本机或 U 盘上的 .xlsx 文件。");
            if (new DriveInfo(Path.GetPathRoot(full)).DriveType == DriveType.Network) throw Invalid("不读取网络共享文件，请先复制到本机或 U 盘。");
            for (string part = full; !String.IsNullOrEmpty(part); part = Path.GetDirectoryName(part))
                if ((File.Exists(part) || Directory.Exists(part)) && (File.GetAttributes(part) & FileAttributes.ReparsePoint) != 0) throw Invalid("请使用实际文件路径，不能通过链接读取台账。");
            FileInfo before = new FileInfo(full);
            if (!before.Exists) throw new FileNotFoundException("没有找到台账文件。", full);
            long size = before.Length;
            DateTime changed = before.LastWriteTimeUtc;
            if (size <= 0 || size > MaxPackageBytes) throw Invalid("文件为空或超过 512 MB，请检查台账文件。");
            byte[] bytes = new byte[(int)size];
            using (FileStream stream = new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                int read = 0;
                while (read < bytes.Length)
                {
                    int count = stream.Read(bytes, read, Math.Min(1024 * 1024, bytes.Length - read));
                    if (count == 0) throw Invalid("读取时文件发生变化，请保存并关闭 Excel 后重试。");
                    read += count;
                }
                if (stream.Length != size) throw Invalid("读取时文件发生变化，请保存并关闭 Excel 后重试。");
            }
            FileInfo after = new FileInfo(full);
            if (after.Length != size || after.LastWriteTimeUtc != changed) throw Invalid("读取时文件发生变化，请保存并关闭 Excel 后重试。");
            return bytes;
        }

        private static Dictionary<string, ZipArchiveEntry> ValidateArchive(ZipArchive archive)
        {
            var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            long total = 0;
            if (archive.Entries.Count > 4096) throw Invalid("台账文件结构过大。");
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                string name = entry.FullName;
                if (name.StartsWith("/", StringComparison.Ordinal) || name.Contains("\\") || name.Contains(":") || name.Split('/').Any(x => x == "..")) throw Invalid("台账包含无效的内部文件路径。");
                if (name.EndsWith("/", StringComparison.Ordinal)) continue;
                if (entries.ContainsKey(name)) throw Invalid("台账包含重复的内部文件。");
                if (entry.Length > MaxPartBytes || entry.Length < 0) throw Invalid("台账单个数据部分过大。");
                total = checked(total + entry.Length);
                if (total > MaxPackageBytes) throw Invalid("台账解压后的数据超过支持范围。");
                entries.Add(name, entry);
            }
            return entries;
        }

        private static XmlDocument XmlPart(Dictionary<string, ZipArchiveEntry> entries, string name, bool required)
        {
            ZipArchiveEntry entry;
            if (!entries.TryGetValue(name, out entry))
            { if (required) throw Invalid("台账缺少必要的数据部分。"); return null; }
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaxPartBytes, MaxCharactersFromEntities = 0 };
            using (Stream stream = entry.Open())
            using (XmlReader reader = XmlReader.Create(stream, settings))
            {
                var xml = new XmlDocument { XmlResolver = null };
                xml.Load(reader);
                return xml;
            }
        }

        private static Dictionary<string, string> ReadRelationships(XmlDocument relationships)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (XmlNode element in relationships.SelectNodes("/*[local-name()='Relationships']/*[local-name()='Relationship']"))
            {
                string type = Attribute(element, "Type");
                if (!type.EndsWith("/worksheet", StringComparison.Ordinal)) continue;
                if (String.Equals(Attribute(element, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) throw Invalid("不读取外部工作表链接。");
                string target = Uri.UnescapeDataString(Attribute(element, "Target"));
                if (target.Contains(":") || target.Contains("\\") || target.Contains("?") || target.Contains("#")) throw Invalid("工作表路径无效。");
                var path = new List<string>();
                if (!target.StartsWith("/", StringComparison.Ordinal)) path.Add("xl");
                foreach (string part in target.Split('/'))
                {
                    if (part == "" || part == ".") continue;
                    if (part == "..") { if (path.Count == 0) throw Invalid("工作表路径越界。"); path.RemoveAt(path.Count - 1); }
                    else path.Add(part);
                }
                string normalized = String.Join("/", path);
                if (!normalized.StartsWith("xl/", StringComparison.Ordinal)) throw Invalid("工作表路径越界。");
                string id = Attribute(element, "Id");
                if (id.Length == 0 || result.ContainsKey(id)) throw Invalid("工作表链接标识重复。");
                result.Add(id, normalized);
            }
            return result;
        }

        private static List<string> ReadSharedStrings(Dictionary<string, ZipArchiveEntry> entries)
        {
            var result = new List<string>();
            XmlDocument xml = XmlPart(entries, "xl/sharedStrings.xml", false);
            if (xml == null) return result;
            foreach (XmlNode item in xml.SelectNodes("/*[local-name()='sst']/*[local-name()='si']"))
            {
                result.Add(RichText(item));
                if (result.Count > 500000) throw Invalid("共享文本条目过多。");
            }
            return result;
        }

        private static List<string> ReadStyles(Dictionary<string, ZipArchiveEntry> entries)
        {
            var result = new List<string>();
            XmlDocument xml = XmlPart(entries, "xl/styles.xml", false);
            if (xml == null) { result.Add(""); return result; }
            var formats = new Dictionary<int, string>();
            foreach (XmlNode item in xml.SelectNodes("/*[local-name()='styleSheet']/*[local-name()='numFmts']/*[local-name()='numFmt']"))
            {
                int id;
                if (Int32.TryParse(Attribute(item, "numFmtId"), out id)) formats[id] = Attribute(item, "formatCode");
            }
            foreach (XmlNode item in xml.SelectNodes("/*[local-name()='styleSheet']/*[local-name()='cellXfs']/*[local-name()='xf']"))
            {
                int id;
                if (!Int32.TryParse(Attribute(item, "numFmtId"), out id)) id = 0;
                string code;
                if (!formats.TryGetValue(id, out code)) code = BuiltinFormat(id);
                result.Add(code);
                if (result.Count > 65536) throw Invalid("单元格格式数量过多。");
            }
            if (result.Count == 0) result.Add("");
            return result;
        }

        private static string BuiltinFormat(int id)
        {
            if (id == 14 || (id >= 27 && id <= 36) || (id >= 50 && id <= 58)) return "yyyy-mm-dd";
            if (id == 15) return "dd-mmm-yy";
            if (id == 16) return "dd-mmm";
            if (id == 17) return "mmm-yy";
            if (id == 22) return "yyyy-mm-dd hh:mm";
            if ((id >= 18 && id <= 21) || (id >= 45 && id <= 47)) return "hh:mm:ss";
            return "";
        }

        private static List<Row> ReadRows(XmlDocument xml, List<string> strings, List<string> styles, bool date1904)
        {
            var result = new List<Row>();
            int previous = 0, cells = 0;
            foreach (XmlNode element in xml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='sheetData']/*[local-name()='row']"))
            {
                int rowNumber;
                string rowAttribute = Attribute(element, "r");
                if (rowAttribute.Length == 0) rowNumber = previous + 1;
                else if (!Int32.TryParse(rowAttribute, out rowNumber)) throw Invalid("台账行号无效。");
                if (rowNumber <= previous || rowNumber > 1048576) throw Invalid("台账行号重复或顺序无效。");
                previous = rowNumber;
                var row = new Row { Number = rowNumber };
                int lastColumn = -1;
                foreach (XmlNode cell in element.SelectNodes("*[local-name()='c']"))
                {
                    if (++cells > 1000000) throw Invalid("台账单元格数量过多。");
                    string reference = Attribute(cell, "r");
                    int column = lastColumn + 1;
                    if (reference.Length > 0)
                    {
                        Match match = Regex.Match(reference, @"^([A-Z]{1,3})([0-9]+)$", RegexOptions.IgnoreCase);
                        if (!match.Success || Int32.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) != rowNumber) throw Invalid("台账单元格位置无效。");
                        column = ColumnNumber(match.Groups[1].Value);
                    }
                    if (column <= lastColumn || column > 16383) throw Invalid("台账单元格位置重复或无效。");
                    lastColumn = column;
                    Cell value = ReadCell(cell, strings, styles, date1904);
                    if (value.Value.Length > 0 || value.Problem != null) row.Cells.Add(column, value);
                }
                // Formatting-only rows do not become records or consume the record limit.
                if (row.Cells.Count > 0) result.Add(row);
                if (result.Count > MaxRecords + 1001) throw Invalid("台账数据超过 20000 条支持上限。");
            }
            return result;
        }

        private static Cell ReadCell(XmlNode cell, List<string> strings, List<string> styles, bool date1904)
        {
            string type = Attribute(cell, "t");
            XmlNode value = cell.SelectSingleNode("*[local-name()='v']");
            string raw = value == null ? "" : value.InnerText;
            if (type == "e") return new Cell { Value = "", Problem = "含 Excel 错误值" };
            if (type == "inlineStr") return new Cell { Value = RichText(cell.SelectSingleNode("*[local-name()='is']")) };
            if (cell.SelectSingleNode("*[local-name()='f']") != null && value == null) return new Cell { Value = "", Problem = "公式没有已保存的计算结果，请在 Excel 中重新计算并保存" };
            if (type == "s")
            {
                int index;
                if (!Int32.TryParse(raw, out index) || index < 0 || index >= strings.Count) throw Invalid("共享文本索引无效。");
                return new Cell { Value = strings[index] };
            }
            if (type == "str") return new Cell { Value = DecodeText(raw) };
            if (type == "b") return new Cell { Value = raw == "1" ? "TRUE" : "FALSE" };
            if (type == "d")
            {
                DateTime date;
                if (!DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date)) throw Invalid("单元格日期无效。");
                return new Cell { Value = date.TimeOfDay == TimeSpan.Zero ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) };
            }
            int style = 0;
            string styleAttribute = Attribute(cell, "s");
            if (styleAttribute.Length > 0 && (!Int32.TryParse(styleAttribute, out style) || style < 0 || style >= styles.Count)) throw Invalid("单元格格式索引无效。");
            return new Cell { Value = FormatNumber(raw, styles[style], date1904) };
        }

        private static string FormatNumber(string raw, string format, bool date1904)
        {
            if (raw.Length == 0) return "";
            double number;
            if (!Double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out number) || Double.IsNaN(number) || Double.IsInfinity(number)) throw Invalid("数字单元格内容无效。");
            string clean = Regex.Replace(format ?? "", "\"[^\"]*\"|\\\\.|_\\S|\\*\\S|\\[[^\\]]*\\]", "").ToLowerInvariant();
            bool isDate = Regex.IsMatch(clean, "[yd]");
            bool isTime = Regex.IsMatch(clean, "[hs]");
            if (isDate || isTime)
            {
                DateTime date;
                try
                {
                    if (number < 0 || number > 2958465) throw new ArgumentOutOfRangeException();
                    double corrected = !date1904 && number < 60 ? number + 1 : number;
                    date = (date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30)).AddDays(corrected);
                }
                catch (ArgumentOutOfRangeException) { throw Invalid("Excel 日期数值超出支持范围。"); }
                if (isDate && isTime) return date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                if (isTime) return date.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                if (!Regex.IsMatch(clean, "d") && !Regex.IsMatch(clean, "m")) return date.ToString("yyyy", CultureInfo.InvariantCulture);
                if (!Regex.IsMatch(clean, "d")) return date.ToString("yyyy-MM", CultureInfo.InvariantCulture);
                if (!date1904 && Math.Floor(number) == 60) return "1900-02-29";
                return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            if (Regex.IsMatch(clean, "^0{2,30}$") && number >= 0 && Math.Floor(number) == number)
                return number.ToString(new string('0', clean.Length), CultureInfo.InvariantCulture);
            return raw;
        }

        private static void IdentifyHeader(Sheet sheet)
        {
            for (int i = 0; i < sheet.Rows.Count && i < 100; i++)
            {
                Row row = sheet.Rows[i];
                var fields = new Dictionary<int, string>();
                foreach (var pair in row.Cells)
                {
                    string field;
                    if (pair.Value.Problem == null && FieldNames.TryGetValue(NormalizeHeader(pair.Value.Value), out field)) fields.Add(pair.Key, field);
                }
                if (fields.Count < 5 || !fields.Values.Any(x => x == "DeviceNumber" || x == "BiosSerial" || x == "DiskSerials" || x == "MacAddresses") || !fields.Values.Any(x => x == "DeviceType" || x == "ManufacturerModel")) continue;
                if (fields.Values.GroupBy(x => x).Any(x => x.Count() > 1)) throw Invalid("台账表头存在重复字段，请检查列名。");
                sheet.HeaderIndex = i;
                sheet.Fields = fields;
                return;
            }
        }

        private static WorkbookSnapshot BuildSnapshot(Sheet sheet)
        {
            Row header = sheet.Rows[sheet.HeaderIndex];
            int width = header.Cells.Keys.Max() + 1;
            if (width > 256) throw Invalid("台账表头超过支持的列数。");
            var result = new WorkbookSnapshot {
                Records = new List<InventoryRecord>(), StorageIds = new List<string>(), ExportedIds = new List<string>(),
                Headers = Enumerable.Range(0, width).Select(c => Text(header, c)).ToArray(),
                MappedFields = new Dictionary<int, string>(sheet.Fields),
                RawRows = new List<string[]>(), OriginalRowNumbers = new List<int>()
            };
            int identityColumn = Array.FindIndex(result.Headers, h => NormalizeHeader(h) == NormalizeHeader(IdentityHeader));
            if (result.Headers.Count(h => NormalizeHeader(h) == NormalizeHeader(IdentityHeader)) > 1) throw Invalid("记录标识列重复，请检查台账表头。");
            bool newFormat = result.Headers.Any(h => NormalizeHeader(h) == NormalizeHeader("物理MAC"));
            foreach (Row row in sheet.Rows.Skip(sheet.HeaderIndex + 1))
            {
                bool hasData = sheet.Fields.Any(p => !String.IsNullOrWhiteSpace(Text(row, p.Key)) || (row.Cells.ContainsKey(p.Key) && row.Cells[p.Key].Problem != null));
                string identity = identityColumn < 0 ? "" : Text(row, identityColumn);
                if (!hasData && identity.Length == 0) continue;
                bool hasDevice = sheet.Fields.Any(p => IsDeviceField(p.Value) && !String.IsNullOrWhiteSpace(Text(row, p.Key)));
                foreach (int column in sheet.Fields.Keys)
                {
                    Cell cell;
                    if (row.Cells.TryGetValue(column, out cell) && cell.Problem != null) throw Invalid("第 " + row.Number + " 行" + cell.Problem + "。");
                }
                if (!hasData) throw Invalid("第 " + row.Number + " 行保留了记录标识但没有台账内容。");
                if (!hasDevice) throw Invalid("第 " + row.Number + " 行没有设备信息，请先核对是否为备注或合计行。");
                var record = new InventoryRecord { HardwareDataVersion = newFormat ? 3 : 0 };
                foreach (var pair in sheet.Fields)
                {
                    string value = Text(row, pair.Key);
                    typeof(InventoryRecord).GetProperty(pair.Value).SetValue(record, value, null);
                }
                // Import must never invent a collection time or a hardware identity.
                result.Records.Add(record);
                result.StorageIds.Add(identity);
                result.RawRows.Add(Enumerable.Range(0, width).Select(c => Text(row, c)).ToArray());
                result.OriginalRowNumbers.Add(row.Number);
                if (result.Records.Count > MaxRecords) throw Invalid("台账记录超过 20000 条支持上限。");
            }
            return result;
        }

        private static void ReadIndex(List<Row> rows, List<string> ids)
        {
            Row first = rows.FirstOrDefault(r => r.Number == 1);
            if (first == null || Text(first, 0) != IndexMarker || first.Cells.Any(c => c.Value.Problem != null || (c.Key != 0 && c.Value.Value.Length > 0))) throw Invalid("台账同步索引格式无效。");
            foreach (Row row in rows.Where(r => r.Number > 1))
            {
                if (row.Cells.Any(c => c.Value.Problem != null || (c.Key != 0 && c.Value.Value.Length > 0))) throw Invalid("台账同步索引格式无效。");
                Cell cell;
                if (row.Cells.TryGetValue(0, out cell) && cell.Problem != null) throw Invalid("台账同步索引格式无效。");
                string value = Text(row, 0);
                if (value.Length > 0) { ValidateIdentity(value); ids.Add(value); }
                if (ids.Count > MaxRecords) throw Invalid("台账同步索引过大。");
            }
            if (ids.Count != ids.Distinct(StringComparer.OrdinalIgnoreCase).Count()) throw Invalid("台账同步索引包含重复记录标识。");
        }

        private static string Text(Row row, int column)
        { Cell value; return row.Cells.TryGetValue(column, out value) ? value.Value : ""; }
        private static string Attribute(XmlNode node, string name)
        { XmlAttribute attribute = node.Attributes[name]; return attribute == null ? "" : attribute.Value; }
        private static int ColumnNumber(string letters)
        { int value = 0; foreach (char c in letters.ToUpperInvariant()) value = value * 26 + c - 'A' + 1; return value - 1; }
        private static string RichText(XmlNode parent)
        {
            if (parent == null) return "";
            var text = new StringBuilder();
            foreach (XmlNode item in parent.SelectNodes("*[local-name()='t']|*[local-name()='r']/*[local-name()='t']")) text.Append(item.InnerText);
            return DecodeText(text.ToString());
        }
        private static string DecodeText(string value)
        {
            string text = Regex.Replace(value, @"_x([0-9A-Fa-f]{4})_", m => ((char)Int32.Parse(m.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString());
            if (text.Length > 32767) throw Invalid("单元格文字超过 Excel 支持的长度。");
            try { XmlConvert.VerifyXmlChars(text); } catch (XmlException) { throw Invalid("台账文字中含有不支持的控制字符。"); }
            return text.Replace("\r\n", "\n").Replace("\r", "\n");
        }
        private static string NormalizeHeader(string value)
        { return Regex.Replace((value ?? "").Trim(), @"\s", "").Replace("（", "(").Replace("）", ")").ToUpperInvariant(); }
        private static bool IsDeviceField(string field)
        { return field == "DeviceNumber" || field == "DeviceType" || field == "ManufacturerModel" || field == "BiosSerial" || field == "DiskSerials" || field == "MacAddresses" || field == "Cpu" || field == "Memory" || field == "OperatingSystem"; }
        private static void ValidateIdentity(string id)
        {
            long sequence; Guid guid;
            if (id == null || id.Length != 58 || id[20] != '_' || !id.EndsWith(".json", StringComparison.Ordinal) ||
                !Int64.TryParse(id.Substring(0, 20), NumberStyles.None, CultureInfo.InvariantCulture, out sequence) || sequence <= 0 ||
                !Guid.TryParseExact(id.Substring(21, 32), "N", out guid)) throw Invalid("台账记录标识格式无效，请使用完整原表。");
        }
        private static Dictionary<string, string> CreateFieldNames()
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            Add(result, "DeviceNumber", "设备编号", "资产编号");
            Add(result, "DeviceType", "设备类型", "电脑类型");
            Add(result, "ManufacturerModel", "品牌及型号", "品牌型号");
            Add(result, "BiosSerial", "整机序列号", "整机序列号（BIOS）", "整机编码");
            Add(result, "DiskSerials", "硬盘（型号 / 编码）", "硬盘序列号", "硬盘（型号/序列号）");
            Add(result, "MacAddresses", "物理MAC", "物理MAC地址", "网卡MAC", "MAC地址");
            Add(result, "Cpu", "CPU", "处理器"); Add(result, "Memory", "内存");
            Add(result, "OperatingSystem", "操作系统", "系统"); Add(result, "Site", "所属站点", "站点", "站点 / 部门");
            Add(result, "Location", "位置", "所在位置", "房间 / 位置", "存放地点"); Add(result, "AssignedTo", "使用人员", "使用人", "实际使用人");
            Add(result, "DeviceStatus", "设备状态"); Add(result, "AcquiredDate", "取得日期", "购置日期");
            Add(result, "Remarks", "备注"); Add(result, "CollectedAt", "采集时间"); Add(result, "Verification", "待核对", "待核对说明");
            Add(result, "IPAddresses", "IP地址（如有）", "IP地址"); Add(result, "AdapterPairs", "网卡IP与MAC配对明细");
            return result;
        }
        private static void Add(Dictionary<string, string> map, string field, params string[] names)
        { foreach (string name in names) map[NormalizeHeader(name)] = field; }
        private static InvalidDataException Invalid(string message) { return new InvalidDataException(message); }
    }
}
