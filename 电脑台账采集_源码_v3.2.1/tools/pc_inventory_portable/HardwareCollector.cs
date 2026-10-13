using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Management;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PcInventory
{
    // Queries are local, read-only, and limited to the inventory fields below.
    // No personal files, credentials, installed-program list or remote machines are read.
    public static class HardwareCollector
    {
        private const string Pending = "待核对";
        private static readonly string[] QueryClasses = {
            "Win32_ComputerSystem", "Win32_ComputerSystemProduct", "Win32_BIOS",
            "Win32_SystemEnclosure", "Win32_Processor", "Win32_OperatingSystem",
            "Win32_DiskDrive", "Win32_NetworkAdapter"
        };
        private static readonly string[] QueryFields = {
            "Manufacturer,Model,TotalPhysicalMemory", "UUID", "SerialNumber",
            "ChassisTypes", "Name", "Caption,Version,OSArchitecture",
            "Index,Model,SerialNumber,InterfaceType,PNPDeviceID,MediaType",
            "Index,Name,Description,PhysicalAdapter,MACAddress,PNPDeviceID,GUID"
        };

        public static InventoryRecord Collect()
        {
            var data = new Dictionary<string, List<Dictionary<string, object>>>(StringComparer.OrdinalIgnoreCase);
            var issues = new List<string>();
            for (int i = 0; i < QueryClasses.Length; i++)
                data[QueryClasses[i]] = ReadLocalQuery(QueryClasses[i], QueryFields[i], issues, @"\\.\root\cimv2");
            // Optional on older Windows. This provider supplies the hardware flag and
            // factory address; no IP or network-configuration fields are requested.
            data["MSFT_NetAdapter"] = ReadLocalQuery("MSFT_NetAdapter",
                "InterfaceGuid,PnPDeviceID,HardwareInterface,Virtual,PermanentAddress,InterfaceDescription",
                null, @"\\.\root\StandardCimv2");
            return BuildRecord(data, issues);
        }

        private static List<Dictionary<string, object>> ReadLocalQuery(string className, string fields,
            List<string> issues, string localNamespace)
        {
            // WMI provider timeouts are not honoured by every driver. Bound the caller's
            // wait as well; a stalled worker cannot block the UI or collect more fields.
            Task<List<Dictionary<string, object>>> task = Task.Run(delegate
            {
                var rows = new List<Dictionary<string, object>>();
                try
                {
                    var options = new ConnectionOptions();
                    options.Timeout = TimeSpan.FromSeconds(3);
                    var scope = new ManagementScope(localNamespace, options);
                    var enumeration = new EnumerationOptions();
                    enumeration.Timeout = TimeSpan.FromSeconds(3);
                    enumeration.ReturnImmediately = true;
                    enumeration.Rewindable = false;
                    using (var searcher = new ManagementObjectSearcher(scope,
                        new ObjectQuery("SELECT " + fields + " FROM " + className), enumeration))
                    using (var results = searcher.Get())
                    {
                        foreach (ManagementObject item in results)
                        {
                            using (item)
                            {
                                var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                                foreach (string field in fields.Split(','))
                                {
                                    try { row[field] = item[field]; }
                                    catch (ManagementException) { row[field] = null; }
                                }
                                rows.Add(row);
                            }
                        }
                    }
                    return rows;
                }
                catch { return null; }
            });
            if (!task.Wait(TimeSpan.FromSeconds(5)))
            {
                if (issues != null) issues.Add(QueryIssue(className));
                return new List<Dictionary<string, object>>();
            }
            if (task.Result == null)
            {
                if (issues != null) issues.Add(QueryIssue(className));
                return new List<Dictionary<string, object>>();
            }
            return task.Result;
        }

        // Pure conversion entry point permits tests with made-up WMI data. This method
        // itself never queries a computer and never stores the raw firmware UUID.
        internal static InventoryRecord BuildRecord(
            IDictionary<string, List<Dictionary<string, object>>> data, IList<string> readIssues)
        {
            var notes = new List<string>();
            if (readIssues != null) notes.AddRange(readIssues);
            var record = new InventoryRecord();
            record.HardwareDataVersion = 3;
            record.RecordId = Guid.NewGuid().ToString("N");
            record.CollectedAt = DateTimeOffset.Now.ToString("o", CultureInfo.InvariantCulture);
            record.Site = "";
            record.Location = "";
            record.AssignedTo = "";
            record.DeviceNumber = "";
            record.DeviceStatus = "待核实";
            record.AcquiredDate = "";
            record.Remarks = "";

            var computer = First(data, "Win32_ComputerSystem");
            var bios = First(data, "Win32_BIOS");
            string manufacturer = Text(computer, "Manufacturer");
            string model = Text(computer, "Model");
            string serial = Text(bios, "SerialNumber");
            record.ManufacturerModel = Display(manufacturer) + " / " + Display(model);
            if (String.IsNullOrWhiteSpace(manufacturer) || String.IsNullOrWhiteSpace(model)) notes.Add("品牌型号待核对");
            record.BiosSerial = SerialDisplay(serial);
            if (!IsUsableSerial(serial)) notes.Add("整机序列号待核对");

            string uuid = Text(First(data, "Win32_ComputerSystemProduct"), "UUID");
            record.DeviceKey = CreateDeviceKey(uuid, serial, manufacturer, model, StablePhysicalMacs(data), notes);
            record.DeviceType = InferDeviceType(Value(First(data, "Win32_SystemEnclosure"), "ChassisTypes"));
            if (record.DeviceType == "待核实") notes.Add("设备类型待核对");

            var processors = new List<string>();
            foreach (var cpu in Rows(data, "Win32_Processor"))
            {
                string cpuName = Text(cpu, "Name");
                if (cpuName.Length > 0) processors.Add(cpuName);
            }
            record.Cpu = JoinOrPending(processors, " / ");
            if (processors.Count == 0) notes.Add("CPU待核对");
            double memory;
            record.Memory = Double.TryParse(Text(computer, "TotalPhysicalMemory"), NumberStyles.Any,
                CultureInfo.InvariantCulture, out memory) && memory > 0
                ? (memory / 1073741824d).ToString("0.##", CultureInfo.InvariantCulture) + " GiB"
                : Pending;
            if (record.Memory == Pending) notes.Add("内存待核对");
            var os = First(data, "Win32_OperatingSystem");
            var osParts = new List<string>();
            AddNonEmpty(osParts, Text(os, "Caption"));
            AddNonEmpty(osParts, Text(os, "Version"));
            AddNonEmpty(osParts, Text(os, "OSArchitecture"));
            record.OperatingSystem = JoinOrPending(osParts, " / ");
            if (osParts.Count == 0) notes.Add("操作系统待核对");

            var disks = new List<string>();
            var orderedDisks = new List<Dictionary<string, object>>(Rows(data, "Win32_DiskDrive"));
            orderedDisks.Sort(delegate(Dictionary<string, object> a, Dictionary<string, object> b)
            {
                return NumericIndex(a).CompareTo(NumericIndex(b));
            });
            foreach (var disk in orderedDisks)
            {
                string diskModel = Text(disk, "Model");
                string pnp = Text(disk, "PNPDeviceID");
                string media = Text(disk, "MediaType");
                if (String.Equals(Text(disk, "InterfaceType"), "USB", StringComparison.OrdinalIgnoreCase)
                    || Regex.IsMatch(pnp, @"^USB(STOR)?\\", RegexOptions.IgnoreCase)
                    || Regex.IsMatch(media, "Removable", RegexOptions.IgnoreCase))
                {
                    continue;
                }
                if (Regex.IsMatch(diskModel + " " + pnp, "Virtual|VMware|VBOX|QEMU|Xen|Msft Storage Space",
                    RegexOptions.IgnoreCase))
                {
                    continue;
                }
                string diskSerial = Text(disk, "SerialNumber");
                string line = "硬盘" + Display(Text(disk, "Index")) + " — " + Display(diskModel)
                    + " — " + SerialDisplay(diskSerial);
                disks.Add(line);
                if (!IsUsableSerial(diskSerial)) notes.Add("硬盘" + Text(disk, "Index") + "编码待核对");
                if (String.IsNullOrWhiteSpace(diskModel)) notes.Add("硬盘型号待核对");
            }
            record.DiskSerials = JoinOrPending(disks, "\n");
            if (disks.Count == 0) notes.Add("硬盘信息待核对");

            var macs = PhysicalMacs(data, notes);
            record.MacAddresses = JoinOrPending(macs, "\n");
            record.IPAddresses = "";
            record.AdapterPairs = "";
            if (macs.Count == 0) notes.Add("物理MAC待核对");
            var distinctNotes = new List<string>();
            foreach (string note in notes)
                if (!String.IsNullOrWhiteSpace(note) && !distinctNotes.Contains(note)) distinctNotes.Add(note);
            record.Verification = distinctNotes.Count == 0 ? "无" : String.Join("；", distinctNotes);
            return record;
        }

        internal static string CreateDeviceKey(string uuid, string biosSerial, string manufacturer, string model,
            IList<string> notes)
        {
            return CreateDeviceKey(uuid, biosSerial, manufacturer, model, "", notes);
        }

        internal static string CreateDeviceKey(string uuid, string biosSerial, string manufacturer, string model,
            string physicalMacs, IList<string> notes)
        {
            Guid id;
            string compact = (uuid ?? "").Trim();
            if (Guid.TryParse(compact, out id))
            {
                compact = id.ToString("N").ToUpperInvariant();
                if (!Regex.IsMatch(compact, @"^(0+|F+)$")
                    && compact != "03000200040005000006000700080009"
                    && compact != "00020003000400050006000700080009")
                {
                    string identity = compact + "\n" + (manufacturer ?? "").Trim().ToUpperInvariant()
                        + "\n" + (model ?? "").Trim().ToUpperInvariant();
                    if (IsUsableSerial(biosSerial))
                        return "uuid-" + Hash(identity + "\nBIOS:" + biosSerial.Trim().ToUpperInvariant());
                    if (!String.IsNullOrWhiteSpace(physicalMacs))
                    {
                        return "uuid-" + Hash(identity + "\nMAC:" + physicalMacs.Trim().ToUpperInvariant());
                    }
                    if (notes != null) notes.Add("设备识别码待核对");
                    return "unreliable-" + Guid.NewGuid().ToString("N");
                }
            }
            if (IsUsableSerial(biosSerial) && IsUsableSerial(manufacturer) && IsUsableSerial(model))
            {
                return "bios-" + Hash(biosSerial.Trim().ToUpperInvariant() + "\n"
                    + manufacturer.Trim().ToUpperInvariant() + "\n" + model.Trim().ToUpperInvariant());
            }
            if (notes != null) notes.Add("设备识别码待核对");
            return "unreliable-" + Guid.NewGuid().ToString("N");
        }

        private static string StablePhysicalMacs(IDictionary<string, List<Dictionary<string, object>>> data)
        {
            var addresses = new SortedSet<string>(StringComparer.Ordinal);
            foreach (string mac in PhysicalMacs(data, null)) addresses.Add(mac.Replace(":", ""));
            return String.Join("|", addresses);
        }

        private static List<string> PhysicalMacs(
            IDictionary<string, List<Dictionary<string, object>>> data, IList<string> notes)
        {
            var addresses = new SortedSet<string>(StringComparer.Ordinal);
            var adapters = new List<Dictionary<string, object>>(Rows(data, "Win32_NetworkAdapter"));
            adapters.Sort(delegate(Dictionary<string, object> a, Dictionary<string, object> b)
            {
                return NumericIndex(a).CompareTo(NumericIndex(b));
            });
            foreach (var adapter in adapters)
            {
                string description = Text(adapter, "Name") + " " + Text(adapter, "Description");
                string pnp = Text(adapter, "PNPDeviceID");
                if (IsExcludedAdapter(description, pnp)) continue;
                var native = MatchingNativeAdapter(adapter, Rows(data, "MSFT_NetAdapter"));
                if (native != null && (IsExcludedAdapter(Text(native, "InterfaceDescription"),
                    Text(native, "PnPDeviceID")) || Flag(native, "Virtual", true)
                    || Flag(native, "HardwareInterface", false))) continue;
                bool knownPhysical = Flag(adapter, "PhysicalAdapter", true)
                    && Regex.IsMatch(pnp, @"^(PCI|USB|PCMCIA|SD|ACPI)\\", RegexOptions.IgnoreCase);
                bool nativePhysical = native != null && Flag(native, "HardwareInterface", true)
                    && Flag(native, "Virtual", false);
                if (Flag(adapter, "PhysicalAdapter", false) || (!knownPhysical && !nativePhysical))
                {
                    // A driver-supplied name or MAC alone never proves physical hardware.
                    if (notes != null && !Flag(adapter, "PhysicalAdapter", false)) notes.Add("网卡类型待核对");
                    continue;
                }
                string mac = NormalizeMac(Text(native, "PermanentAddress"));
                if (mac.Length == 0) mac = NormalizeMac(Text(adapter, "MACAddress"));
                if (mac.Length == 0)
                {
                    if (notes != null) notes.Add("物理MAC待核对");
                    continue;
                }
                // A locally administered address may have been randomized or overridden.
                if ((Int32.Parse(mac.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) & 2) != 0
                    && notes != null) notes.Add("MAC地址待核对");
                addresses.Add(mac);
            }
            return new List<string>(addresses);
        }

        private static Dictionary<string, object> MatchingNativeAdapter(Dictionary<string, object> adapter,
            List<Dictionary<string, object>> candidates)
        {
            string pnp = Text(adapter, "PNPDeviceID");
            Guid id;
            bool hasId = Guid.TryParse(Text(adapter, "GUID"), out id) && id != Guid.Empty;
            foreach (var candidate in candidates)
            {
                string nativePnp = Text(candidate, "PnPDeviceID");
                Guid nativeId;
                bool sameId = hasId && Guid.TryParse(Text(candidate, "InterfaceGuid"), out nativeId) && id == nativeId;
                bool samePnp = pnp.Length > 0 && nativePnp.Length > 0
                    && String.Equals(pnp, nativePnp, StringComparison.OrdinalIgnoreCase);
                if (sameId || samePnp) return candidate;
            }
            return null;
        }

        private static bool IsExcludedAdapter(string description, string pnp)
        {
            return Regex.IsMatch(description ?? "",
                @"Virtual|VMware|Hyper-V|VBOX|VPN|Tunnel|Loopback|Bluetooth|WAN Miniport|Wintun|WireGuard|(?:^|[^A-Z])(?:TAP|TUN)(?:[^A-Z]|$)",
                RegexOptions.IgnoreCase)
                || Regex.IsMatch(pnp ?? "", @"^(ROOT|SWD|VMBUS)\\", RegexOptions.IgnoreCase);
        }

        internal static string NormalizeMac(string value)
        {
            string compact = Regex.Replace(value ?? "", @"[\s:\-.]", "").ToUpperInvariant();
            if (!Regex.IsMatch(compact, "^[0-9A-F]{12}$") || Regex.IsMatch(compact, "^(0+|F+)$")) return "";
            if ((Int32.Parse(compact.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) & 1) != 0) return "";
            var groups = new string[6];
            for (int i = 0; i < groups.Length; i++) groups[i] = compact.Substring(i * 2, 2);
            return String.Join(":", groups);
        }

        private static bool Flag(Dictionary<string, object> row, string field, bool expected)
        {
            bool flag;
            return Boolean.TryParse(Text(row, field), out flag) && flag == expected;
        }

        private static long NumericIndex(Dictionary<string, object> row)
        {
            long index;
            return Int64.TryParse(Text(row, "Index"), NumberStyles.None, CultureInfo.InvariantCulture, out index)
                ? index : Int64.MaxValue;
        }

        private static string QueryIssue(string className)
        {
            switch (className)
            {
                case "Win32_ComputerSystem": return "品牌型号及内存待核对";
                case "Win32_ComputerSystemProduct": return "设备识别码待核对";
                case "Win32_BIOS": return "整机序列号待核对";
                case "Win32_SystemEnclosure": return "设备类型待核对";
                case "Win32_Processor": return "CPU待核对";
                case "Win32_OperatingSystem": return "操作系统待核对";
                case "Win32_DiskDrive": return "硬盘信息待核对";
                default: return "物理MAC待核对";
            }
        }

        internal static bool IsUsableSerial(string value)
        {
            if (String.IsNullOrWhiteSpace(value)) return false;
            string text = value.Trim().Trim('\0').Trim();
            if (text.Length == 0) return false;
            if (Regex.IsMatch(text,
                @"^(To\s*Be\s*Filled\s*By\s*O[.]?E[.]?M[.]?|Default\s*string|System\s*Serial\s*Number|Serial\s*Number|System\s*Product\s*Name|System\s*manufacturer|None|Unknown|Not\s*Specified|Not\s*Available|Not\s*Applicable|N/?A|OEM|INVALID)$",
                RegexOptions.IgnoreCase)) return false;
            string compact = Regex.Replace(text, @"[\s\-:{}_.]", "");
            if (compact.Length == 0 || Regex.IsMatch(compact, @"^(0+|F+)$", RegexOptions.IgnoreCase)) return false;
            return true;
        }

        private static string SerialDisplay(string value)
        {
            if (String.IsNullOrWhiteSpace(value)) return Pending;
            return IsUsableSerial(value) ? value.Trim() : Pending;
        }

        private static string Hash(string text)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(text));
                var result = new StringBuilder(bytes.Length * 2);
                foreach (byte b in bytes) result.Append(b.ToString("x2", CultureInfo.InvariantCulture));
                return result.ToString();
            }
        }

        internal static string InferDeviceType(object chassisTypes)
        {
            var values = StringValues(chassisTypes);
            foreach (string value in values)
                if (value == "13") return "一体机";
            foreach (string value in values)
                if (value == "8" || value == "9" || value == "10" || value == "14"
                    || value == "31" || value == "32") return "笔记本";
            foreach (string value in values)
                if (value == "3" || value == "4" || value == "5" || value == "6"
                    || value == "7" || value == "15" || value == "16") return "台式机";
            return "待核实";
        }

        private static List<Dictionary<string, object>> Rows(
            IDictionary<string, List<Dictionary<string, object>>> data, string className)
        {
            List<Dictionary<string, object>> rows;
            if (data != null && data.TryGetValue(className, out rows) && rows != null) return rows;
            return new List<Dictionary<string, object>>();
        }

        private static Dictionary<string, object> First(
            IDictionary<string, List<Dictionary<string, object>>> data, string className)
        {
            var rows = Rows(data, className);
            return rows.Count == 0 ? null : rows[0];
        }

        private static object Value(Dictionary<string, object> row, string field)
        {
            object value;
            return row != null && row.TryGetValue(field, out value) ? value : null;
        }

        private static string Text(Dictionary<string, object> row, string field)
        {
            return Convert.ToString(Value(row, field), CultureInfo.InvariantCulture).Trim().Trim('\0').Trim();
        }

        private static string Display(string value)
        {
            return String.IsNullOrWhiteSpace(value) ? Pending : value.Trim();
        }

        private static void AddNonEmpty(List<string> list, string value)
        {
            if (!String.IsNullOrWhiteSpace(value)) list.Add(value.Trim());
        }

        private static string JoinOrPending(List<string> values, string separator)
        {
            return values.Count == 0 ? Pending : String.Join(separator, values);
        }

        private static List<string> StringValues(object value)
        {
            var values = new List<string>();
            if (value == null) return values;
            var array = value as IEnumerable;
            if (array != null && !(value is string))
            {
                foreach (object item in array) AddNonEmpty(values, Convert.ToString(item, CultureInfo.InvariantCulture));
            }
            else AddNonEmpty(values, Convert.ToString(value, CultureInfo.InvariantCulture));
            return values;
        }
    }
}
