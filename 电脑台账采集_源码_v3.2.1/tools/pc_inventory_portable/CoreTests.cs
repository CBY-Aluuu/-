using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using PcInventory;

public static class CoreTests
{
    private static int checks;
    private static void Assert(bool value, string text)
    {
        if (!value) throw new Exception("FAILED: " + text);
        checks++;
        Console.WriteLine("PASS: " + text);
    }

    private static InventoryRecord Fake(string key, string number, string user, string clock)
    {
        return new InventoryRecord {
            DeviceKey = key, DeviceNumber = number, AssignedTo = user, CollectedAt = clock,
            Site = "虚拟测试站", Location = "测试房间", DeviceType = "台式机", DeviceStatus = "在用",
            ManufacturerModel = "测试电脑 A&B <model>", BiosSerial = "000123", DiskSerials = "000456",
            MacAddresses = "00:00:00:00:00:01", IPAddresses = "192.0.2.1", Cpu = "虚拟CPU", Memory = "8 GB",
            OperatingSystem = "虚拟系统", Remarks = "=HYPERLINK(\"https://invalid.example\",\"not a formula\")",
            Verification = "全部为虚拟测试资料"
        };
    }

    public static int Main(string[] args)
    {
        try
        {
            if (args.Length != 1) throw new Exception("Specify a local test output directory.");
            string root = Path.Combine(Path.GetFullPath(args[0]), "test_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            InventoryStore firstComputer = new InventoryStore(root);
            SaveResult a = firstComputer.Save(Fake("PC-A", "00001", "甲", "2099-01-01"));
            Assert(a.SummaryUpdated && a.BackupUpdated && a.Count == 1, "first device saved with XLSX and complete backup");
            byte[] originalBackup = File.ReadAllBytes(a.BackupPath);
            InventoryStore secondComputer = new InventoryStore(root);
            SaveResult b = secondComputer.Save(Fake("PC-A", "00002", "乙", "2001-01-01"));
            Assert(b.Count == 2 && secondComputer.LoadLatest().Count == 2, "second store retains first computer");
            SaveResult update = secondComputer.Save(Fake("PC-A", "00001", "更新后人员 <&>", "1990-01-01"));
            Assert(update.Count == 3 && secondComputer.LoadRecords().Count == 3, "same device saved three times appends three rows");
            Assert(new InventoryStore(root).LoadRecords()[0].AssignedTo == "甲" && new InventoryStore(root).LoadRecords()[2].AssignedTo == "更新后人员 <&>", "save order preserved despite incorrect computer clock; old rows unchanged");
            Assert(Directory.GetFiles(Path.Combine(root, "数据", "记录"), "*.json").Length == 3, "all historical records retained");
            string backup = Path.Combine(root, "数据", "汇总备份", "电脑台账汇总_上一版.xlsx");
            Assert(File.Exists(backup) && ReadWorksheet(backup).OuterXml.Contains("甲"), "previous workbook retained in backup");
            Assert(Directory.GetFiles(firstComputer.BackupDirectory, "*.zip").Length == 3 && CountBackup(a.BackupPath) == 1 && CountBackup(update.BackupPath) == 3, "each save creates unique complete backup preserving history");
            Assert(Convert.ToBase64String(originalBackup) == Convert.ToBase64String(File.ReadAllBytes(a.BackupPath)), "old backup bytes are never overwritten");

            string summary = firstComputer.SummaryPath;
            using (FileStream heldByExcel = new FileStream(summary, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                SaveResult locked = secondComputer.Save(Fake("PC-C", "00003", "丙", "2026-01-01"));
                Assert(!locked.SummaryUpdated && locked.Count == 4 && locked.BackupUpdated && CountBackup(locked.BackupPath) == 4 && locked.Message.Contains("已追加"), "Excel lock still appends record and independently creates full backup");
            }
            Assert(new InventoryStore(root).LoadRecords().Count == 4, "record from locked summary persists across restart");
            firstComputer.RebuildSummary();
            XmlDocument sheet = ReadWorksheet(summary);
            XmlNamespaceManager ns = new XmlNamespaceManager(sheet.NameTable);
            ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            Assert(sheet.SelectSingleNode("//s:c[@r='B2']/s:is/s:t", ns).InnerText == "00001", "leading zeros preserved");
            Assert(sheet.SelectSingleNode("//s:c[@r='D2']/s:is/s:t", ns).InnerText == "测试电脑 A&B <model>", "XML special characters round trip");
            Assert(sheet.SelectSingleNode("//s:c[@r='P2']/@t", ns).Value == "inlineStr" && sheet.SelectSingleNode("//s:c[@r='P2']/s:is/s:t", ns).InnerText.StartsWith("=HYPERLINK") && sheet.SelectNodes("//s:f", ns).Count == 0, "formula-looking value stored as literal text");
            Assert(sheet.SelectSingleNode("//s:pane/@state", ns).Value == "frozen" && sheet.SelectSingleNode("//s:autoFilter/@ref", ns).Value == "A1:S5", "frozen header and filter cover every saved row including its hidden identity");
            Assert(sheet.SelectSingleNode("//s:c[@r='O2']/s:v", ns).InnerText == "1", "quantity always numeric one");

            InventoryRecord unknown = Fake(null, "00004", "未知甲", "2026-01-01");
            firstComputer.Save(unknown);
            InventoryRecord anotherUnknown = Fake(null, "00005", "未知乙", "2026-01-01");
            anotherUnknown.RecordId = unknown.RecordId;
            firstComputer.Save(anotherUnknown);
            Assert(firstComputer.LoadRecords().Count == 6 && anotherUnknown.RecordId != unknown.RecordId, "unknown identities append and supplied repeated RecordId is replaced");
            string beforeRepeatedSave = anotherUnknown.RecordId;
            firstComputer.Save(anotherUnknown);
            Assert(firstComputer.LoadRecords().Count == 7 && beforeRepeatedSave != anotherUnknown.RecordId, "saving identical object again appends with fresh RecordId");
            VerifyLegacy(root);
            VerifyBackupFailure(root);
            VerifyRestore(root, firstComputer);

            byte[] original = File.ReadAllBytes(summary);
            string badRecord = Path.Combine(root, "数据", "记录", "00000000000000009999_bad.json");
            File.WriteAllText(badRecord, "{broken", new UTF8Encoding(false));
            bool corruptLoad = false, corruptRebuild = false, corruptSave = false;
            try { firstComputer.LoadLatest(); } catch (InvalidDataException ex) { corruptLoad = ex.Message.Contains("bad.json"); }
            try { firstComputer.RebuildSummary(); } catch (InvalidDataException) { corruptRebuild = true; }
            try { firstComputer.Save(Fake("PC-D", "00006", "丁", "2026-01-01")); } catch (InvalidDataException) { corruptSave = true; }
            Assert(corruptLoad && corruptRebuild && corruptSave, "corrupt records block load, rebuild and save with filename");
            Assert(Convert.ToBase64String(original) == Convert.ToBase64String(File.ReadAllBytes(summary)), "corruption cannot overwrite existing summary");
            File.Move(badRecord, badRecord + ".test-quarantine");

            InventoryRecord invalid = Fake("PC-X", "00099", "坏字符\u0001", "2026-01-01");
            bool rejected = false;
            try { firstComputer.Save(invalid); } catch (InvalidDataException) { rejected = true; }
            Assert(rejected && firstComputer.LoadRecords().Count == 7, "invalid XML character rejected before persistence");
            string interrupted = Path.Combine(root, "数据", "记录", ".saving_interrupted.tmp");
            File.WriteAllText(interrupted, "{interrupted");
            bool interruptedVisible = false;
            try { firstComputer.LoadLatest(); } catch (InvalidDataException ex) { interruptedVisible = ex.Message.Contains("上次保存中断"); }
            Assert(interruptedVisible, "interrupted save is visible and never silently skipped");
            File.Move(interrupted, interrupted + ".test-quarantine");
            bool networkRejected = false;
            try { new InventoryStore(@"\\server\share\records"); } catch (IOException) { networkRejected = true; }
            Assert(networkRejected, "UNC network paths rejected without network access");
            bool networkBackup = false, networkRestore = false;
            try { firstComputer.BackupNow(@"\\server\share\backup"); } catch (IOException) { networkBackup = true; }
            try { firstComputer.RestoreBackupToNewFolder(@"\\server\share\backup.zip", root); } catch (IOException) { networkRestore = true; }
            Assert(networkBackup && networkRestore, "manual backup and restore reject UNC paths before network access");
            Console.WriteLine("ALL " + checks + " CHECKS PASSED");
            Console.WriteLine("Synthetic test files: " + root);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    private static void VerifyLegacy(string root)
    {
        string legacyRoot = Path.Combine(root, "legacy-v1");
        string records = Path.Combine(legacyRoot, "数据", "记录");
        Directory.CreateDirectory(records);
        Dictionary<string, string> originals = new Dictionary<string, string>();
        for (int i = 1; i <= 9; i++)
        {
            InventoryRecord record = Fake("SAME-LEGACY-DEVICE", i.ToString("D5"), "旧版测试" + i, (2100 - i).ToString() + "-01-01");
            record.RecordId = "same-legacy-id";
            string path = Path.Combine(records, i.ToString("D20") + "_" + Guid.NewGuid().ToString("N") + ".json");
            using (MemoryStream stream = new MemoryStream())
            {
                new DataContractJsonSerializer(typeof(InventoryRecord)).WriteObject(stream, record);
                File.WriteAllText(path, "{\"FormatVersion\":1,\"SaveSequence\":" + i + ",\"Record\":" + Encoding.UTF8.GetString(stream.ToArray()) + "}", new UTF8Encoding(false));
            }
            originals[path] = Convert.ToBase64String(File.ReadAllBytes(path));
        }
        InventoryStore legacy = new InventoryStore(legacyRoot);
        Assert(legacy.LoadRecords().Count == 9 && legacy.LoadRecords()[8].AssignedTo == "旧版测试9", "all nine v1 records reappear despite identical legacy device and record IDs");
        legacy.RebuildSummary();
        Assert(WorkbookRows(legacy.SummaryPath) == 10, "legacy summary rebuilt with all nine historical rows");
        SaveResult upgraded = legacy.Save(Fake("SAME-LEGACY-DEVICE", "00010", "升级追加", "1900-01-01"));
        Assert(upgraded.Count == 10 && upgraded.BackupUpdated && CountBackup(upgraded.BackupPath) == 10, "v2 saves append after v1 sequence and back up both formats");
        bool unchanged = true;
        foreach (KeyValuePair<string, string> item in originals) unchanged &= item.Value == Convert.ToBase64String(File.ReadAllBytes(item.Key));
        Assert(unchanged, "upgrade and rebuild never rewrite original legacy JSON");
    }

    private static void VerifyBackupFailure(string root)
    {
        string failRoot = Path.Combine(root, "backup-failure");
        Directory.CreateDirectory(failRoot);
        File.WriteAllText(Path.Combine(failRoot, "备份"), "test file deliberately blocks backup directory");
        InventoryStore store = new InventoryStore(failRoot);
        SaveResult result = store.Save(Fake("FAIL-A", "01", "甲", "2026-01-01"));
        Assert(result.Count == 1 && result.SummaryUpdated && !result.BackupUpdated && String.IsNullOrEmpty(result.BackupPath) && result.Message.Contains("自动备份失败"), "backup failure does not misreport durable record or successful summary");
        using (FileStream locked = new FileStream(store.SummaryPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            SaveResult both = store.Save(Fake("FAIL-A", "02", "乙", "2026-01-01"));
            Assert(both.Count == 2 && !both.SummaryUpdated && !both.BackupUpdated && both.Message.Contains("已追加"), "summary and backup can both fail while raw record remains saved");
        }
        Assert(new InventoryStore(failRoot).LoadRecords().Count == 2, "raw records survive failed derivative outputs across restart");
        string manual = store.BackupNow(Path.Combine(root, "other-local-disk"));
        Assert(File.Exists(manual) && CountBackup(manual) == 2, "manual backup to another local directory works after automatic destination failure");
        store.RebuildSummary();
        Assert(WorkbookRows(store.SummaryPath) == 3, "summary repairs later using all saved records");
    }

    private static void VerifyRestore(string root, InventoryStore store)
    {
        string otherDisk = Path.Combine(root, "manual-backups");
        string backup = store.BackupNow(otherDisk);
        string backup2 = store.BackupNow(otherDisk);
        Assert(backup != backup2 && File.Exists(backup) && File.Exists(backup2), "manual backups never overwrite previous ZIPs");
        string transferred = Path.Combine(root, "only-transferred-backup.zip");
        File.Copy(backup, transferred);
        string parent = Path.Combine(root, "restored-copies");
        Directory.CreateDirectory(parent);
        string marker = Path.Combine(parent, "existing-file.txt");
        File.WriteAllText(marker, "must remain");
        string originalSummary = Convert.ToBase64String(File.ReadAllBytes(store.SummaryPath));
        string restored = store.RestoreBackupToNewFolder(transferred, parent);
        InventoryStore copy = new InventoryStore(restored);
        Assert(copy.LoadRecords().Count == 7 && WorkbookRows(copy.SummaryPath) == 8, "one transferred ZIP restores all records and rebuilds workbook in new directory");
        Assert(File.ReadAllText(marker) == "must remain" && originalSummary == Convert.ToBase64String(File.ReadAllBytes(store.SummaryPath)), "restore does not touch existing files or original summary");
        string restored2 = store.RestoreBackupToNewFolder(transferred, parent);
        Assert(restored != restored2 && Directory.Exists(restored) && Directory.Exists(restored2), "repeated restore creates separate folders without overwrite");
        SaveResult saved = copy.Save(Fake("PC-A", "restored", "恢复后追加", "2026-01-01"));
        Assert(saved.Count == 8 && saved.BackupUpdated && store.LoadRecords().Count == 7, "restored folder supports independent append and backup");

        string[] attacks = { "tamper", "traversal", "absolute", "duplicate", "case-mismatch" };
        foreach (string attack in attacks)
        {
            string bad = Path.Combine(root, "bad-" + attack + ".zip");
            RewriteZip(transferred, bad, attack);
            AssertRejectedRestore(store, bad, parent, "invalid backup rejected before extraction: " + attack);
        }
        Assert(!File.Exists(Path.Combine(root, "escape.json")), "malicious ZIP cannot escape restore folder");
        string oversized = Path.Combine(root, "oversized.zip");
        using (FileStream output = File.Create(oversized))
        using (ZipArchive zip = new ZipArchive(output, ZipArchiveMode.Create))
        using (Stream entry = zip.CreateEntry("records/00000000000000000001_0123456789abcdef0123456789abcdef.json").Open())
        { byte[] zeros = new byte[8192]; for (int i = 0; i <= 1024; i++) entry.Write(zeros, 0, zeros.Length); }
        AssertRejectedRestore(store, oversized, parent, "oversized decompressed JSON is rejected");
        string truncated = Path.Combine(root, "truncated.zip");
        byte[] archive = File.ReadAllBytes(transferred);
        using (FileStream output = File.Create(truncated)) output.Write(archive, 0, archive.Length / 2);
        AssertRejectedRestore(store, truncated, parent, "truncated ZIP leaves current data unchanged");
        if (typeof(ZipArchiveEntry).GetProperty("ExternalAttributes") != null)
        {
            string symlink = Path.Combine(root, "symlink.zip");
            RewriteZip(transferred, symlink, "symlink");
            AssertRejectedRestore(store, symlink, parent, "symbolic link metadata in ZIP is rejected");
        }
    }

    private static void AssertRejectedRestore(InventoryStore store, string zip, string parent, string text)
    {
        int directories = Directory.GetDirectories(parent).Length;
        bool rejected = false;
        try { store.RestoreBackupToNewFolder(zip, parent); } catch (IOException) { rejected = true; }
        Assert(rejected && directories == Directory.GetDirectories(parent).Length && store.LoadRecords().Count == 7, text);
    }

    private static void RewriteZip(string source, string destination, string mode)
    {
        using (FileStream input = File.OpenRead(source))
        using (ZipArchive oldZip = new ZipArchive(input, ZipArchiveMode.Read))
        using (FileStream output = File.Create(destination))
        using (ZipArchive newZip = new ZipArchive(output, ZipArchiveMode.Create))
        {
            bool changed = false;
            foreach (ZipArchiveEntry oldEntry in oldZip.Entries)
            {
                string name = oldEntry.FullName;
                byte[] data;
                using (Stream stream = oldEntry.Open())
                using (MemoryStream memory = new MemoryStream()) { stream.CopyTo(memory); data = memory.ToArray(); }
                if (!changed && name.StartsWith("records/"))
                {
                    if (mode == "tamper") data[0] = (byte)'[';
                    if (mode == "traversal") name = "records/../../escape.json";
                    if (mode == "absolute") name = "C:/escape.json";
                    if (mode == "case-mismatch") name = name.ToUpperInvariant().Replace("RECORDS/", "records/").Replace(".JSON", ".json");
                    changed = true;
                }
                ZipArchiveEntry entry = newZip.CreateEntry(name);
                if (mode == "symlink" && name.StartsWith("records/")) typeof(ZipArchiveEntry).GetProperty("ExternalAttributes").SetValue(entry, unchecked((int)0xA1FF0000), null);
                using (Stream stream = entry.Open()) stream.Write(data, 0, data.Length);
                if (mode == "duplicate" && name == "manifest.json")
                    using (Stream duplicate = newZip.CreateEntry(name).Open()) duplicate.Write(data, 0, data.Length);
            }
        }
    }

    private static int CountBackup(string path)
    {
        using (FileStream input = File.OpenRead(path))
        using (ZipArchive zip = new ZipArchive(input, ZipArchiveMode.Read))
        {
            int count = 0;
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                if (entry.FullName.StartsWith("records/", StringComparison.Ordinal)) count++;
                else if (entry.FullName != "manifest.json") throw new Exception("Unexpected backup contents: " + entry.FullName);
            }
            if (zip.GetEntry("manifest.json") == null) throw new Exception("Missing backup manifest.");
            return count;
        }
    }

    private static int WorkbookRows(string path)
    {
        XmlDocument xml = ReadWorksheet(path);
        XmlNamespaceManager ns = new XmlNamespaceManager(xml.NameTable);
        ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        return xml.SelectNodes("//s:sheetData/s:row", ns).Count;
    }

    private static XmlDocument ReadWorksheet(string path)
    {
        using (FileStream input = File.OpenRead(path))
        using (ZipArchive zip = new ZipArchive(input, ZipArchiveMode.Read))
        using (Stream sheet = zip.GetEntry("xl/worksheets/sheet1.xml").Open())
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(sheet);
            return xml;
        }
    }
}
