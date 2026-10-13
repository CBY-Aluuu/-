using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using PcInventory;

public static class HistoryStoreTests
{
    private static int checks;
    private static void Check(bool value, string name) { if (!value) throw new Exception("FAILED: " + name); checks++; Console.WriteLine("PASS " + name); }
    private static InventoryRecord Fixture(string number)
    {
        return new InventoryRecord { HardwareDataVersion = 3, DeviceNumber = number, DeviceKey = "SYNTHETIC", CollectedAt = "2026-10-13T10:00:00+08:00", Site = "示例站", Location = "示例室", AssignedTo = "测试人员", DeviceType = "台式机", DeviceStatus = "在用", ManufacturerModel = "SAMPLE", BiosSerial = "00001", DiskSerials = "硬盘0 — SAMPLE — 00002", MacAddresses = "10:22:33:44:55:66", Cpu = "SAMPLE CPU", Memory = "8 GB", OperatingSystem = "SAMPLE OS", AcquiredDate = "2023-08-15", Remarks = "=文字\n第二行", Verification = "无" };
    }
    private static void EditXml(ZipArchive zip, string path, Action<XmlDocument> edit)
    {
        ZipArchiveEntry entry = zip.GetEntry(path); var xml = new XmlDocument();
        using (Stream input = entry.Open()) xml.Load(input);
        edit(xml); entry.Delete(); using (Stream output = zip.CreateEntry(path).Open()) xml.Save(output);
    }
    private static void RemoveRows(string path, params int[] rows)
    {
        using (var file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
        using (var zip = new ZipArchive(file, ZipArchiveMode.Update))
            EditXml(zip, "xl/worksheets/sheet1.xml", delegate(XmlDocument xml) {
                foreach (int row in rows) { XmlNode node = xml.SelectSingleNode("//*[local-name()='sheetData']/*[local-name()='row'][@r='" + row + "']"); if (node != null) node.ParentNode.RemoveChild(node); }
            });
    }
    private static void MakeLegacy(string path)
    {
        using (var file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
        using (var zip = new ZipArchive(file, ZipArchiveMode.Update))
        {
            EditXml(zip, "xl/worksheets/sheet1.xml", delegate(XmlDocument xml) {
                foreach (XmlNode node in xml.SelectNodes("//*[local-name()='row']/*[local-name()='c'][starts-with(@r,'S')]").Cast<XmlNode>().ToArray()) node.ParentNode.RemoveChild(node);
            });
            EditXml(zip, "xl/workbook.xml", delegate(XmlDocument xml) { XmlNode node = xml.SelectSingleNode("//*[local-name()='sheet'][@name='记录同步索引']"); node.ParentNode.RemoveChild(node); });
            EditXml(zip, "xl/_rels/workbook.xml.rels", delegate(XmlDocument xml) { XmlNode node = xml.SelectSingleNode("//*[local-name()='Relationship'][@Id='rId3']"); node.ParentNode.RemoveChild(node); });
            EditXml(zip, "[Content_Types].xml", delegate(XmlDocument xml) { XmlNode node = xml.SelectSingleNode("//*[local-name()='Override'][@PartName='/xl/worksheets/sheet2.xml']"); node.ParentNode.RemoveChild(node); });
            zip.GetEntry("xl/worksheets/sheet2.xml").Delete();
        }
    }
    private static void CopyTree(string source, string destination)
    {
        Directory.CreateDirectory(destination);
        foreach (string file in Directory.GetFiles(source)) File.Copy(file, Path.Combine(destination, Path.GetFileName(file)));
        foreach (string dir in Directory.GetDirectories(source)) CopyTree(dir, Path.Combine(destination, Path.GetFileName(dir)));
    }
    public static int Main(string[] args)
    {
        try
        {
            string root = Path.GetFullPath(args[0]); Directory.CreateDirectory(root);
            var legacy = new InventoryStore(Path.Combine(root, "legacy"));
            legacy.Save(Fixture("L1")); legacy.Save(Fixture("L2")); MakeLegacy(legacy.SummaryPath);
            using (var locked = new FileStream(legacy.SummaryPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            {
                SyncResult sync = legacy.SynchronizeSummary();
                Check(sync.SummaryReadable && sync.Records.Count == 2 && !sync.Changed, "locked legacy workbook establishes baseline without deletion");
                Check(!InventoryWorkbookReader.ReadSnapshot(legacy.SummaryPath).HasSyncIdentity, "locked legacy file is not overwritten");
                SaveResult pending = legacy.Save(Fixture("L3"));
                Check(pending.Count == 3 && !pending.SummaryUpdated && pending.BackupUpdated, "append while legacy file is locked keeps recoverable pending row");
                Check(legacy.SynchronizeSummary().SummaryReadable, "pending append does not invalidate legacy baseline");
            }
            RemoveRows(legacy.SummaryPath, 2);
            SyncResult reconciled = new InventoryStore(Path.Combine(root, "legacy")).SynchronizeSummary();
            Check(reconciled.Changed && reconciled.SummaryReadable && reconciled.Records.Count == 2, "legacy deletion after restart reconciles against exported baseline");
            Check(reconciled.Records.Any(r => r.DeviceNumber == "L3") && !reconciled.Records.Any(r => r.DeviceNumber == "L1"), "legacy unexported row survives deletion of older exported row");
            Check(InventoryWorkbookReader.ReadSnapshot(legacy.SummaryPath).HasSyncIdentity, "legacy workbook auto upgrades after lock is released");
            Check(legacy.RebuildSummary() == legacy.SummaryPath && legacy.LoadRecords().Count == 2, "rebuild cannot resurrect legacy removed row");

            var empty = new InventoryStore(Path.Combine(root, "all-deleted"));
            empty.Save(Fixture("E1")); empty.Save(Fixture("E2"));
            RemoveRows(empty.SummaryPath, 2, 3);
            SyncResult removedAll = empty.SynchronizeSummary();
            Check(removedAll.Changed && removedAll.Records.Count == 0 && removedAll.SummaryReadable, "deleting all workbook rows yields zero active records");
            Check(new InventoryStore(Path.Combine(root, "all-deleted")).LoadRecords().Count == 0, "zero-record deletion state persists across restart");
            Check(Directory.GetFiles(Path.Combine(root, "all-deleted", "数据", "记录"), "*.json").Length == 2, "delete preserves immutable original records");
            string zeroBackup = empty.BackupNow(Path.Combine(root, "zero-backup"));
            string zeroRestore = empty.RestoreBackupToNewFolder(zeroBackup, Path.Combine(root, "restore-zero"));
            Check(new InventoryStore(zeroRestore).LoadRecords().Count == 0, "backup and restore preserve all-deleted state");
            SaveResult afterZero = empty.Save(Fixture("E3"));
            Check(afterZero.SummaryUpdated && afterZero.Count == 1, "append after deleting all starts current count at one");
            Check(empty.LoadRecords()[0].StorageId.StartsWith("00000000000000000003_"), "save sequence does not reuse deleted original sequence");

            byte[] valid = File.ReadAllBytes(empty.SummaryPath); File.WriteAllText(empty.SummaryPath, "partial Excel save", Encoding.UTF8);
            SyncResult invalid = empty.SynchronizeSummary();
            Check(!invalid.SummaryReadable && invalid.Records.Count == 1 && !invalid.Changed, "incomplete Excel save cannot delete rows");
            SaveResult safeAppend = empty.Save(Fixture("E4"));
            Check(safeAppend.Count == 2 && !safeAppend.SummaryUpdated && safeAppend.BackupUpdated, "invalid summary does not lose new raw save or backup");
            Check(File.ReadAllText(empty.SummaryPath, Encoding.UTF8) == "partial Excel save", "unrecognized workbook is never silently overwritten");
            File.WriteAllBytes(empty.SummaryPath, valid);
            Check(empty.SynchronizeSummary().Records.Count == 2, "restored exported workbook does not delete pending unsummarized row");
            Check(InventoryWorkbookReader.ReadSnapshot(empty.RebuildSummary()).Records.Count == 2, "manual update exports pending rows once readable again");

            var duplicate = new InventoryStore(Path.Combine(root, "duplicate")); duplicate.Save(Fixture("D1")); duplicate.Save(Fixture("D2"));
            string[] raw = Directory.GetFiles(Path.Combine(root, "duplicate", "数据", "记录"), "*.json");
            string same = duplicate.LoadRecords()[0].RecordId;
            foreach (string path in raw) { string text = File.ReadAllText(path); string id = duplicate.LoadRecords().First(r => Path.GetFileName(path) == r.StorageId).RecordId; File.WriteAllText(path, text.Replace(id, same), new UTF8Encoding(false)); }
            IList<InventoryRecord> prior = duplicate.LoadRecords();
            Check(prior[0].RecordId == prior[1].RecordId && prior[0].StorageId != prior[1].StorageId, "legacy duplicate RecordId still has unique saved-row identity");
            SaveResult picked = duplicate.DeleteRecords(new[] { prior[1].StorageId });
            Check(picked.Count == 1 && picked.BackupUpdated && duplicate.LoadRecords()[0].DeviceNumber == "D1", "delete uses StorageId and preserves other same-RecordId row");
            string portable = Path.Combine(root, "relocated"); CopyTree(Path.Combine(root, "duplicate"), portable);
            Check(new InventoryStore(portable).LoadRecords().Count == 1, "copying whole folder retains deletion state on another path");

            var oldMismatch = new InventoryStore(Path.Combine(root, "old-mismatch")); oldMismatch.Save(Fixture("M1")); oldMismatch.Save(Fixture("M2"));
            MakeLegacy(oldMismatch.SummaryPath); RemoveRows(oldMismatch.SummaryPath, 3);
            byte[] unchanged = File.ReadAllBytes(oldMismatch.SummaryPath);
            SyncResult unknown = oldMismatch.SynchronizeSummary();
            Check(!unknown.SummaryReadable && unknown.Records.Count == 2, "pre-upgrade unmatched legacy rows require review instead of guessing exported status");
            Check(unchanged.SequenceEqual(File.ReadAllBytes(oldMismatch.SummaryPath)), "unmatched legacy workbook remains intact");
            Console.WriteLine("ALL " + checks + " HISTORY STORE CHECKS PASSED"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
