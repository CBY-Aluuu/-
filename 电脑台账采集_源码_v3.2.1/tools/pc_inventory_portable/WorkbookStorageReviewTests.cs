using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using PcInventory;

public static class WorkbookStorageReviewTests
{
    private static int checks;
    private static void Check(bool value, string message) { if (!value) throw new Exception("FAILED: " + message); checks++; Console.WriteLine("PASS: " + message); }
    private static InventoryRecord Fake(string number)
    {
        return new InventoryRecord { HardwareDataVersion = 3, DeviceKey = "SYNTHETIC", DeviceNumber = number, DeviceType = "台式机", ManufacturerModel = "SAMPLE MODEL", BiosSerial = "00001", DiskSerials = "硬盘0 — SAMPLE — 000077", MacAddresses = "10:22:33:44:55:66", Cpu = "SAMPLE CPU", Memory = "8 GB", OperatingSystem = "SAMPLE OS", Site = "示例站", Location = "示例房间", AssignedTo = "示例人员", DeviceStatus = "在用", CollectedAt = "2026-10-13 10:00", Verification = "无" };
    }
    private static void DeleteWorksheetRow(string path, int number)
    { ChangeWorksheet(path, delegate(XmlDocument xml) { XmlNode row = xml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='sheetData']/*[local-name()='row'][@r='" + number + "']"); row.ParentNode.RemoveChild(row); }); }
    private static void ChangeWorksheet(string path, Action<XmlDocument> change)
    {
        using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
        using (ZipArchive zip = new ZipArchive(file, ZipArchiveMode.Update))
        {
            XmlDocument xml = new XmlDocument();
            ZipArchiveEntry entry = zip.GetEntry("xl/worksheets/sheet1.xml");
            using (Stream input = entry.Open()) xml.Load(input);
            change(xml);
            entry.Delete();
            using (Stream output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open()) xml.Save(output);
        }
    }
    private static void MakeTwoLocationColumns(string path)
    {
        ChangeWorksheet(path, delegate(XmlDocument xml) {
            foreach (XmlElement row in xml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']"))
            {
                string number = row.GetAttribute("r");
                XmlNode location = row.SelectSingleNode("*[local-name()='c'][@r='K" + number + "']");
                XmlNode room = location.CloneNode(true);
                foreach (XmlElement cell in row.SelectNodes("*[local-name()='c']"))
                {
                    string address = cell.GetAttribute("r");
                    if (address[0] >= 'L') cell.SetAttribute("r", ((char)(address[0] + 1)).ToString() + number);
                }
                ((XmlElement)room).SetAttribute("r", "L" + number);
                location.SelectSingleNode("*[local-name()='is']/*[local-name()='t']").InnerText = number == "1" ? "所属站点" : "示例站";
                room.SelectSingleNode("*[local-name()='is']/*[local-name()='t']").InnerText = number == "1" ? "所在位置" : "示例房间";
                row.InsertAfter(room, location);
            }
        });
    }
    public static int Main(string[] args)
    {
        try
        {
            string root = Path.GetFullPath(args[0]); Directory.CreateDirectory(root);
            InventoryStore older = new InventoryStore(Path.Combine(root, "prior-two-location"));
            older.Save(Fake("OLD-1")); older.Save(Fake("OLD-2"));
            string[] olderFiles = Directory.GetFiles(Path.Combine(root, "prior-two-location", "数据", "记录"), "*.json");
            string[] olderBytes = olderFiles.Select(file => Convert.ToBase64String(File.ReadAllBytes(file))).ToArray();
            MakeTwoLocationColumns(older.SummaryPath);
            SyncResult olderSync = older.SynchronizeSummary();
            Check(olderSync.SummaryReadable && !olderSync.Changed && olderSync.Records.Count == 2, "prior separate site and room workbook still synchronizes with existing records");
            ChangeWorksheet(older.SummaryPath, delegate(XmlDocument xml) {
                xml.SelectSingleNode("//*[local-name()='c'][@r='L1']/*[local-name()='is']/*[local-name()='t']").InnerText = "位置";
            });
            olderSync = older.SynchronizeSummary();
            Check(olderSync.SummaryReadable && !olderSync.Changed && olderSync.Records.Count == 2, "legacy location alias beside separate site compares original room without combining site twice");
            DeleteWorksheetRow(older.SummaryPath, 3);
            olderSync = older.SynchronizeSummary();
            Check(olderSync.SummaryReadable && olderSync.Changed && olderSync.Records.Count == 1, "deletion from prior two-location workbook still targets the correct record");
            older.RebuildSummary();
            WorkbookSnapshot upgraded = InventoryWorkbookReader.ReadSnapshot(older.SummaryPath);
            Check(upgraded.Headers.Contains("位置") && !upgraded.Headers.Contains("所属站点") && upgraded.Records[0].Location == "示例站 / 示例房间", "rebuilding prior layout exports one location with both original parts");
            Check(olderFiles.Select(file => Convert.ToBase64String(File.ReadAllBytes(file))).SequenceEqual(olderBytes), "layout upgrade never rewrites original JSON records");
            InventoryStore first = new InventoryStore(Path.Combine(root, "ledgers"), "2026.xlsx");
            SaveResult a = first.Save(Fake("00001")), b = first.Save(Fake("00002"));
            Check(a.SummaryUpdated && b.SummaryUpdated && b.Count == 2, "first named ledger saves both rows");
            InventoryStore second = new InventoryStore(Path.Combine(root, "ledgers"), "2027.xlsx");
            second.Save(Fake("00003"));
            Check(first.LoadRecords().Count == 2 && second.LoadRecords().Count == 1 && first.BackupDirectory != second.BackupDirectory, "named year ledgers isolate records and backups");
            DeleteWorksheetRow(first.SummaryPath, 3);
            string sourceForImport = Path.Combine(root, "for-import.xlsx"); File.Copy(first.SummaryPath, sourceForImport);
            SyncResult sync = first.SynchronizeSummary();
            Check(sync.Changed && sync.SummaryReadable && sync.Records.Count == 1 && sync.Records[0].DeviceNumber == "00001", "Excel row deletion changes only the corresponding saved ID");
            Check(second.LoadRecords().Count == 1, "deleting from one ledger cannot affect the other ledger");
            SaveResult next = first.Save(Fake("00004"));
            Check(next.Count == 2 && next.SummaryUpdated, "append after external deletion does not resurrect deleted row");
            string backup = first.BackupNow(Path.Combine(root, "manual-backup"));
            string before = Convert.ToBase64String(File.ReadAllBytes(first.SummaryPath));
            string restored = first.RestoreBackupToNewFolder(backup, Path.Combine(root, "restore"));
            InventoryStore restoredStore = new InventoryStore(restored, "2026.xlsx");
            Check(restoredStore.LoadRecords().Count == 2 && File.Exists(restoredStore.SummaryPath), "backup restores active rows and selected ledger filename");
            Check(before == Convert.ToBase64String(File.ReadAllBytes(first.SummaryPath)), "restore never changes source workbook");
            string picked = first.LoadRecords()[0].StorageId;
            SaveResult deleted = first.DeleteRecords(new[] { picked });
            Check(deleted.Count == 1 && deleted.BackupUpdated && first.LoadRecords()[0].DeviceNumber == "00004", "application deletion preserves unrelated records and makes backup");

            string importRoot = Path.Combine(root, "interrupted-import"); Directory.CreateDirectory(importRoot);
            File.Copy(sourceForImport, Path.Combine(importRoot, "电脑台账汇总.xlsx"));
            string stateBlocker = Path.Combine(importRoot, "数据", "记录状态.json"); Directory.CreateDirectory(stateBlocker);
            bool interrupted = false;
            try { new InventoryStore(importRoot).LoadRecords(); } catch (IOException) { interrupted = true; }
            Check(interrupted, "state-file write failure is reported during import");
            Directory.Delete(stateBlocker, false);
            InventoryStore restarted = new InventoryStore(importRoot);
            SyncResult recovered = restarted.SynchronizeSummary();
            Check(recovered.SummaryReadable && recovered.Records.Count == 1, "restart completes interrupted import including deleted baseline IDs");
            Check(new InventoryStore(importRoot).Save(Fake("00005")).SummaryUpdated, "recovered import can append and update summary normally");

            string failedBackupRoot = Path.Combine(root, "failed-delete-backup"); Directory.CreateDirectory(failedBackupRoot);
            File.WriteAllText(Path.Combine(failedBackupRoot, "备份"), "synthetic blocker");
            InventoryStore noBackup = new InventoryStore(failedBackupRoot);
            SaveResult originalSave = noBackup.Save(Fake("00006"));
            Check(originalSave.Count == 1 && !originalSave.BackupUpdated && originalSave.SummaryUpdated, "backup failure still preserves initial appended data");
            bool deleteFailed = false;
            try { noBackup.DeleteRecords(new[] { noBackup.LoadRecords()[0].StorageId }); } catch (IOException) { deleteFailed = true; }
            Check(deleteFailed && noBackup.LoadRecords().Count == 1 && InventoryWorkbookReader.Read(noBackup.SummaryPath).Count == 1, "failed pre-deletion backup prevents any record from being deleted");

            InventoryStore pending = new InventoryStore(Path.Combine(root, "pending-summary")); pending.Save(Fake("00007"));
            using (FileStream locked = new FileStream(pending.SummaryPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                SaveResult pendingSave = pending.Save(Fake("00008"));
                Check(pendingSave.Count == 2 && !pendingSave.SummaryUpdated && pendingSave.BackupUpdated, "locked summary leaves second raw record durable with backup");
            }
            SyncResult pendingSync = new InventoryStore(Path.Combine(root, "pending-summary")).SynchronizeSummary();
            Check(pendingSync.SummaryReadable && pendingSync.Records.Count == 2 && !pendingSync.Changed, "record never exported to Excel is not mistaken for an external deletion");
            DeleteWorksheetRow(pending.SummaryPath, 2);
            SyncResult afterOldDeletion = pending.SynchronizeSummary();
            Check(afterOldDeletion.Changed && afterOldDeletion.SummaryReadable && afterOldDeletion.Records.Count == 1 && afterOldDeletion.Records[0].DeviceNumber == "00008", "deleting an old exported row cannot delete the unpublished new record");
            pending.RebuildSummary();
            Check(InventoryWorkbookReader.Read(pending.SummaryPath).Count == 1, "summary rebuild retains pending records and respects the external deletion");

            InventoryStore visibility = new InventoryStore(Path.Combine(root, "visibility")); visibility.Save(Fake("00011")); visibility.Save(Fake("00012"));
            ChangeWorksheet(visibility.SummaryPath, delegate(XmlDocument xml) { XmlElement row = (XmlElement)xml.SelectSingleNode("//*[local-name()='row'][@r='3']"); row.SetAttribute("hidden", "1"); });
            SyncResult hidden = visibility.SynchronizeSummary();
            Check(hidden.SummaryReadable && !hidden.Changed && hidden.Records.Count == 2, "filtered or hidden Excel rows remain active records");
            ChangeWorksheet(visibility.SummaryPath, delegate(XmlDocument xml) {
                XmlNode aText = xml.SelectSingleNode("//*[local-name()='c'][@r='B2']/*[local-name()='is']/*[local-name()='t']");
                XmlNode bText = xml.SelectSingleNode("//*[local-name()='c'][@r='B3']/*[local-name()='is']/*[local-name()='t']");
                string value = aText.InnerText; aText.InnerText = bText.InnerText; bText.InnerText = value;
            });
            SyncResult sorted = visibility.SynchronizeSummary();
            Check(!sorted.SummaryReadable && !sorted.Changed && visibility.LoadRecords().Count == 2, "sorting visible cells without moving hidden IDs stops synchronization safely");
            Console.WriteLine("ALL " + checks + " INDEPENDENT STORAGE REVIEW CHECKS PASSED");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
