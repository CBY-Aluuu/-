using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using PcInventory;

public static class HardwareSimplificationTests
{
    private static int checks;
    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new Exception("FAILED: " + message);
        checks++;
        Console.WriteLine("PASS: " + message);
    }

    private static Dictionary<string, object> Row(params object[] values)
    {
        var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < values.Length; i += 2) row.Add((string)values[i], values[i + 1]);
        return row;
    }

    private static Dictionary<string, List<Dictionary<string, object>>> Fixture()
    {
        return new Dictionary<string, List<Dictionary<string, object>>>(StringComparer.OrdinalIgnoreCase) {
            { "Win32_ComputerSystem", new List<Dictionary<string, object>> { Row("Manufacturer", "Example", "Model", "Model A", "TotalPhysicalMemory", 8589934592L) } },
            { "Win32_BIOS", new List<Dictionary<string, object>> { Row("SerialNumber", "SN-0001") } },
            { "Win32_ComputerSystemProduct", new List<Dictionary<string, object>> { Row("UUID", "11112222-3333-4444-5555-666677778888") } },
            { "Win32_SystemEnclosure", new List<Dictionary<string, object>> { Row("ChassisTypes", new int[] { 3 }) } },
            { "Win32_Processor", new List<Dictionary<string, object>> { Row("Name", "Example CPU") } },
            { "Win32_OperatingSystem", new List<Dictionary<string, object>> { Row("Caption", "Windows Example", "Version", "10.0", "OSArchitecture", "64-bit") } },
            { "Win32_DiskDrive", new List<Dictionary<string, object>> { Disk(0, "SSD-A", "0000_1111_2222_3333.") } },
            { "Win32_NetworkAdapter", new List<Dictionary<string, object>> { Adapter(0, "Example Ethernet", true, "PCI\\VEN_ABCD&DEV_0001", "00:11:22:33:44:55") } },
            { "MSFT_NetAdapter", new List<Dictionary<string, object>>() }
        };
    }

    private static Dictionary<string, object> Adapter(int index, string name, object physical, string pnp, string mac)
    {
        return Row("Index", index, "Name", name, "Description", name, "PhysicalAdapter", physical,
            "PNPDeviceID", pnp, "MACAddress", mac);
    }

    private static Dictionary<string, object> Disk(int index, string model, string serial)
    {
        return Row("Index", index, "Model", model, "SerialNumber", serial, "InterfaceType", "SCSI",
            "PNPDeviceID", "SCSI\\DISK&VEN_EXAMPLE", "MediaType", "Fixed hard disk media", "Size", 1024000000000L);
    }

    public static int Main(string[] args)
    {
        try
        {
            var data = Fixture();
            var record = HardwareCollector.BuildRecord(data, null);
            Assert(record.HardwareDataVersion == 3, "new collection has hardware version 3");
            Assert(record.DiskSerials == "硬盘0 — SSD-A — 0000_1111_2222_3333.", "disk retains exact serial with underscores and punctuation, excludes capacity");
            Assert(record.MacAddresses == "00:11:22:33:44:55", "physical MAC is a bare normalized address");
            Assert(record.IPAddresses == "" && record.AdapterPairs == "", "new collection stores no IP or pair data");
            Assert(record.Verification == "无", "healthy hardware has no routine verification prose");

            data["Win32_NetworkAdapterConfiguration"] = new List<Dictionary<string, object>> { Row("Index", 0,
                "MACAddress", "00:AA:AA:AA:AA:AA", "IPAddress", new string[] { "192.0.2.88", "2001:db8::1" }) };
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.IPAddresses.Length == 0 && record.AdapterPairs.Length == 0 && !record.MacAddresses.Contains("AA:AA"), "old configuration data is ignored");

            data = Fixture();
            data["Win32_NetworkAdapter"].Add(Adapter(13, "VeryKuai TAP Adapter", true, "PCI\\FAKE", "00:FF:04:28:35:E2"));
            data["Win32_NetworkAdapter"].Add(Adapter(14, "WireGuard Tunnel", true, "PCI\\FAKE2", "00:22:33:44:55:66"));
            data["Win32_NetworkAdapter"].Add(Adapter(15, "Virtual Ethernet", true, "PCI\\FAKE3", "00:33:44:55:66:77"));
            data["Win32_NetworkAdapter"].Add(Adapter(16, "Example Ethernet", true, "ROOT\\NET\\0000", "00:44:55:66:77:88"));
            data["Win32_NetworkAdapter"].Add(Adapter(17, "Wintun Adapter", true, "USB\\FAKE", "00:55:66:77:88:99"));
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.MacAddresses == "00:11:22:33:44:55" && record.Verification == "无", "TAP TUN VPN and software devices excluded even with physical flag true");

            data = Fixture();
            data["Win32_NetworkAdapter"].Add(Adapter(1, "Unknown Ethernet", null, "PCI\\VEN_UNKNOWN", "00:22:33:44:55:66"));
            data["Win32_NetworkAdapter"].Add(Adapter(2, "Unknown Ethernet", true, "", "00:33:44:55:66:77"));
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.MacAddresses == "00:11:22:33:44:55" && record.Verification == "网卡类型待核对", "unknown physical flag or hardware bus cannot establish physical device");

            data = Fixture();
            data["Win32_NetworkAdapter"].Add(Adapter(1, "Example USB Ethernet", true, "USB\\VID_ABCD", "00-22-33-44-55-66"));
            data["Win32_NetworkAdapter"].Add(Adapter(2, "Duplicate Ethernet", true, "PCI\\DUP", "001122334455"));
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.MacAddresses == "00:11:22:33:44:55\n00:22:33:44:55:66", "multiple physical NIC addresses use one deduplicated address per line");

            data = Fixture();
            data["Win32_NetworkAdapter"][0]["MACAddress"] = "02:11:22:33:44:55";
            data["MSFT_NetAdapter"].Add(Row("PnPDeviceID", "PCI\\VEN_ABCD&DEV_0001", "HardwareInterface", true,
                "Virtual", false, "PermanentAddress", "00-66-77-88-99-AA", "InterfaceDescription", "Example Ethernet"));
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.MacAddresses == "00:66:77:88:99:AA" && record.Verification == "无", "provider permanent address preferred over overridden current address");
            data["MSFT_NetAdapter"].Clear();
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.MacAddresses == "02:11:22:33:44:55" && record.Verification == "MAC地址待核对", "current locally administered address has one short review note");

            data = Fixture();
            data["Win32_NetworkAdapter"][0]["PhysicalAdapter"] = null;
            data["MSFT_NetAdapter"].Add(Row("PnPDeviceID", "PCI\\VEN_ABCD&DEV_0001", "HardwareInterface", true,
                "Virtual", false, "PermanentAddress", "00:66:77:88:99:AA"));
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.MacAddresses == "00:66:77:88:99:AA", "native hardware provider can confirm unknown Win32 physical flag");
            data["MSFT_NetAdapter"][0]["Virtual"] = true;
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.MacAddresses == "待核对", "native virtual flag vetoes a candidate");

            data = Fixture();
            data["MSFT_NetAdapter"].Add(Row("PnPDeviceID", "PCI\\UNRELATED", "HardwareInterface", true,
                "Virtual", false, "PermanentAddress", "00:66:77:88:99:AA", "InterfaceDescription", "Example Ethernet"));
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.MacAddresses == "00:11:22:33:44:55", "same display name never matches unrelated physical address");

            Assert(HardwareCollector.NormalizeMac("FF:FF:FF:FF:FF:FF") == "" && HardwareCollector.NormalizeMac("00:00:00:00:00:00") == ""
                && HardwareCollector.NormalizeMac("01:00:5E:00:00:01") == "" && HardwareCollector.NormalizeMac("bad-mac") == "", "invalid placeholder and multicast MAC addresses rejected");

            data = Fixture();
            var usb = Disk(3, "Collection USB", "USB-0001");
            usb["InterfaceType"] = "USB";
            data["Win32_DiskDrive"].Add(Disk(10, "SSD-C", "C-010"));
            data["Win32_DiskDrive"].Add(Disk(2, "SSD-B", "B-002"));
            data["Win32_DiskDrive"].Add(usb);
            data["Win32_DiskDrive"].Add(Disk(4, "Virtual Disk", "VIRTUAL-1"));
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.DiskSerials == "硬盘0 — SSD-A — 0000_1111_2222_3333.\n硬盘2 — SSD-B — B-002\n硬盘10 — SSD-C — C-010", "disk numbers sorted numerically; USB and virtual disks excluded");
            Assert(record.Verification == "无", "excluded removable and virtual devices add no routine technical prose");
            data["Win32_DiskDrive"][0]["SerialNumber"] = "Unknown";
            record = HardwareCollector.BuildRecord(data, null);
            Assert(record.DiskSerials.StartsWith("硬盘0 — SSD-A — 待核对") && record.Verification == "硬盘0编码待核对", "unknown disk serial never presented as a real code");
            data["Win32_DiskDrive"][0]["SerialNumber"] = "0000_0000_0000.";
            Assert(!HardwareCollector.IsUsableSerial("0000_0000_0000."), "separated all-zero disk identifier is a placeholder");

            data = Fixture();
            string originalKey = HardwareCollector.BuildRecord(data, null).DeviceKey;
            data["Win32_NetworkAdapter"][0]["MACAddress"] = "00:22:33:44:55:66";
            Assert(HardwareCollector.BuildRecord(data, null).DeviceKey == originalKey, "valid BIOS and UUID device identity remains compatible across MAC changes");

            record = HardwareCollector.BuildRecord(new Dictionary<string, List<Dictionary<string, object>>>(), null);
            Assert(record.Verification.Length < 140 && !record.Verification.Contains("IP") && !record.Verification.Contains("UUID"), "all-missing hardware still yields short plain-language notes");
            var classes = (string[])typeof(HardwareCollector).GetField("QueryClasses", BindingFlags.NonPublic | BindingFlags.Static).GetValue(null);
            var fields = (string[])typeof(HardwareCollector).GetField("QueryFields", BindingFlags.NonPublic | BindingFlags.Static).GetValue(null);
            Assert(!String.Join(",", classes).Contains("NetworkAdapterConfiguration") && !String.Join(",", fields).Contains("IPAddress"), "live query configuration requests no IP data");

            Console.WriteLine("PASS: " + checks + " hardware simplification checks");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
