using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace PcInventory
{
    // Only formats copies for the screen/export. Saved historical records stay intact.
    public static class InventoryPresentation
    {
        private const string Pending = "待核对";
        private const string MacPattern = @"(?<![0-9A-Fa-f])(?:[0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2}(?![0-9A-Fa-f])";
        private const string VirtualPattern = @"Virtual|VMware|Hyper-V|VBOX|TAP|TUN(?:NEL)?\b|VPN|Wintun|WireGuard|Loopback|Bluetooth|WAN Miniport|Wi-Fi Direct";

        public static string SerialText(string value) { return Missing(value) ? Pending : value.Trim(); }
        public static string LocationText(InventoryRecord record)
        {
            if (record == null) return "";
            // Keep both parts of legacy records without requiring two fields for new saves.
            return String.Join(" / ", new[] { record.Site, record.Location }
                .Where(value => !String.IsNullOrWhiteSpace(value)).Select(value => value.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase));
        }
        private static bool Missing(string value)
        {
            return String.IsNullOrWhiteSpace(value) || Regex.IsMatch(value, @"待核|未读|占位|未知|^unknown$|^none$|^N/?A$", RegexOptions.IgnoreCase);
        }

        public static string DiskText(InventoryRecord record)
        {
            string raw = record.DiskSerials ?? "";
            if (Missing(raw) && !Regex.IsMatch(raw, @"(?:物理)?硬盘\s*(?:\d|—)")) return Pending;
            var lines = new List<string>();
            foreach (string item in Regex.Split(raw.Replace("\r", ""), @"\s*\|\s*|\n"))
            {
                string line = item.Trim(); if (line.Length == 0) continue;
                var old = Regex.Match(line, @"^物理硬盘\s*(?<n>\d+)\s*\[(?<model>.*?)\][：:]\s*(?<serial>.*?)(?:[（(]\s*\d+(?:\.\d+)?\s*GB\s*[）)])?$", RegexOptions.IgnoreCase);
                if (old.Success)
                    line = "硬盘" + old.Groups["n"].Value + " — " + SerialText(old.Groups["model"].Value) + " — " + SerialText(old.Groups["serial"].Value);
                else if (record.HardwareDataVersion >= 3)
                    line = line.Replace("未读取到，待核实", Pending).Replace("固件占位值，待核实", Pending).Replace("待核实", Pending);
                else if (!line.Contains("硬盘")) line = "硬盘 — " + Pending + " — " + SerialText(line);
                lines.Add(line);
            }
            return lines.Count == 0 ? Pending : String.Join("\n", lines.OrderBy(DiskIndex));
        }
        private static int DiskIndex(string line)
        {
            var match = Regex.Match(line, @"^硬盘\s*(\d+)"); int index;
            return match.Success && Int32.TryParse(match.Groups[1].Value, out index) ? index : Int32.MaxValue;
        }

        public static string MacText(InventoryRecord record)
        {
            var macs = new List<string>();
            foreach (string item in Regex.Split(record.MacAddresses ?? "", @"\s*\|\s*|[\r\n]+"))
            {
                // Old records may include TAP adapters labelled physical by their driver.
                if (record.HardwareDataVersion < 3 && Regex.IsMatch(item, VirtualPattern, RegexOptions.IgnoreCase)) continue;
                foreach (Match match in Regex.Matches(item, MacPattern))
                {
                    string mac = match.Value.Replace('-', ':').ToUpperInvariant();
                    int first = Int32.Parse(mac.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    if ((first & 1) != 0 || mac == "00:00:00:00:00:00" || mac == "FF:FF:FF:FF:FF:FF") continue;
                    if (!macs.Contains(mac)) macs.Add(mac);
                }
            }
            return macs.Count == 0 ? Pending : String.Join("\n", macs);
        }

        public static string VerificationText(InventoryRecord record)
        {
            if (record.HardwareDataVersion >= 3)
                return String.IsNullOrWhiteSpace(record.Verification) ? "无" : record.Verification.Trim();
            var notes = new List<string>();
            if (Missing(record.BiosSerial)) notes.Add("整机编码待核对");
            if (Missing(DiskText(record))) notes.Add("硬盘编码待核对");
            notes.Add(MacText(record) == Pending ? "物理MAC待核对" : "旧MAC请重采");
            if (Missing(record.ManufacturerModel)) notes.Add("型号待核对");
            if (Missing(record.DeviceType)) notes.Add("类型待核对");
            if (Missing(record.Cpu)) notes.Add("CPU待核对");
            if (Missing(record.Memory)) notes.Add("内存待核对");
            if (Missing(record.OperatingSystem)) notes.Add("系统待核对");
            return String.Join("；", notes);
        }

        public static string CollectedTime(string value)
        {
            DateTimeOffset time;
            return DateTimeOffset.TryParse(value, out time) ? time.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) : value ?? "";
        }
    }
}
