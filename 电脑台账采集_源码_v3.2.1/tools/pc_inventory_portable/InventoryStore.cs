using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace PcInventory
{
    [DataContract]
    public class InventoryRecord
    {
        [DataMember] public string RecordId { get; set; }
        [DataMember] public string DeviceKey { get; set; }
        [DataMember] public string CollectedAt { get; set; }
        [DataMember] public string Site { get; set; }
        [DataMember] public string Location { get; set; }
        [DataMember] public string AssignedTo { get; set; }
        [DataMember] public string DeviceNumber { get; set; }
        [DataMember] public string DeviceType { get; set; }
        [DataMember] public string DeviceStatus { get; set; }
        [DataMember] public string AcquiredDate { get; set; }
        [DataMember] public string Remarks { get; set; }
        [DataMember] public string ManufacturerModel { get; set; }
        [DataMember] public string BiosSerial { get; set; }
        [DataMember] public string DiskSerials { get; set; }
        [DataMember] public string MacAddresses { get; set; }
        [DataMember] public string IPAddresses { get; set; }
        [DataMember] public string AdapterPairs { get; set; }
        [DataMember] public string Cpu { get; set; }
        [DataMember] public string Memory { get; set; }
        [DataMember] public string OperatingSystem { get; set; }
        [DataMember] public string Verification { get; set; }
        [DataMember] public int HardwareDataVersion { get; set; }
        // A raw filename identifies a saved row even when an older record reused RecordId.
        [IgnoreDataMember] public string StorageId { get; set; }
    }

    public class SyncResult
    {
        public List<InventoryRecord> Records { get; set; }
        public bool Changed { get; set; }
        public bool SummaryReadable { get; set; }
        public string Message { get; set; }
    }

    public class SaveResult
    {
        public string Message { get; set; }
        public int Count { get; set; }
        public bool SummaryUpdated { get; set; }
        public bool BackupUpdated { get; set; }
        public string BackupPath { get; set; }
    }

    [DataContract]
    internal class StoredRecord
    {
        [DataMember(IsRequired = true)] public int FormatVersion { get; set; }
        [DataMember(IsRequired = true)] public long SaveSequence { get; set; }
        [DataMember(IsRequired = true)] public InventoryRecord Record { get; set; }
    }

    [DataContract]
    internal class BackupManifest
    {
        [DataMember(IsRequired = true)] public int FormatVersion { get; set; }
        [DataMember(IsRequired = true)] public string CreatedUtc { get; set; }
        [DataMember(IsRequired = true)] public string Instructions { get; set; }
        [DataMember(IsRequired = true)] public List<BackupFile> Records { get; set; }
        [DataMember(EmitDefaultValue = false)] public List<string> DeletedStorageIds { get; set; }
        [DataMember(EmitDefaultValue = false)] public string SummaryFileName { get; set; }
    }

    [DataContract]
    internal class RecordState
    {
        [DataMember(IsRequired = true)] public int FormatVersion { get; set; }
        [DataMember(IsRequired = true)] public List<string> DeletedStorageIds { get; set; }
        [DataMember(EmitDefaultValue = false)] public List<string> ImportedBaselineIds { get; set; }
        [DataMember(EmitDefaultValue = false)] public List<string> LegacyBaselineIds { get; set; }
    }

    [DataContract]
    internal class ImportCommit
    {
        [DataMember(IsRequired = true)] public int FormatVersion { get; set; }
        [DataMember(IsRequired = true)] public string StagingDirectory { get; set; }
        [DataMember(IsRequired = true)] public List<string> RecordFiles { get; set; }
        [DataMember(IsRequired = true)] public List<string> ImportedBaselineIds { get; set; }
    }

    [DataContract]
    internal class BackupFile
    {
        [DataMember(IsRequired = true)] public string Path { get; set; }
        [DataMember(IsRequired = true)] public long Length { get; set; }
        [DataMember(IsRequired = true)] public string Sha256 { get; set; }
    }

    public class InventoryStore
    {
        private readonly string baseDirectory;
        private readonly string dataDirectory;
        private readonly string recordsDirectory;
        private readonly string summaryFileName;
        private readonly string backupDirectory;
        private string lastReadSummaryHash;
        private bool lastSummaryObservedExists;
        private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool MoveFileExW(string existingFileName, string newFileName, uint flags);
        private const uint MoveReplaceExisting = 0x1;
        private const uint MoveWriteThrough = 0x8;
        private const int MaxBackupRecords = 20000;
        private const long MaxRecordBytes = 8L * 1024 * 1024;
        private const long MaxManifestBytes = 16L * 1024 * 1024;
        private const long MaxBackupBytes = 512L * 1024 * 1024;
        public string SummaryPath { get { return Path.Combine(baseDirectory, summaryFileName); } }
        public string SummaryFileName { get { return summaryFileName; } }
        public string BackupDirectory { get { return backupDirectory; } }
        private string StatePath { get { return Path.Combine(dataDirectory, "记录状态.json"); } }
        private string ImportCommitPath { get { return Path.Combine(dataDirectory, "导入提交.json"); } }
        private const string SyncSheetName = "记录同步索引";
        private const string SyncMarker = "PC-INVENTORY-SYNC-1";
        private const string StorageHeader = "记录标识（勿修改）";
        private static readonly string[] VisibleHeaders = { "序号", "设备编号", "设备类型", "品牌及型号", "整机序列号", "硬盘（型号 / 编码）", "物理MAC", "CPU", "内存", "操作系统", "位置", "使用人员", "设备状态", "取得日期", "数量", "备注", "采集时间", "待核对" };

        public InventoryStore(string baseDirectory) : this(baseDirectory, "电脑台账汇总.xlsx") { }

        public InventoryStore(string baseDirectory, string summaryFileName)
        {
            if (String.IsNullOrWhiteSpace(baseDirectory)) throw new ArgumentException("请把整个工具文件夹放在本机或单位允许使用的U盘中。");
            ValidateSummaryFileName(summaryFileName);
            this.baseDirectory = Path.GetFullPath(baseDirectory);
            this.summaryFileName = summaryFileName;
            this.dataDirectory = Path.Combine(this.baseDirectory, "数据");
            this.backupDirectory = Path.Combine(this.baseDirectory, "备份");
            if (!String.Equals(summaryFileName, "电脑台账汇总.xlsx", StringComparison.OrdinalIgnoreCase))
            {
                string key;
                using (SHA256 hash = SHA256.Create()) key = BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(summaryFileName.ToUpperInvariant()))).Replace("-", "").ToLowerInvariant();
                this.dataDirectory = Path.Combine(this.dataDirectory, "台账", key.Substring(0, 24));
                string title = Path.GetFileNameWithoutExtension(summaryFileName);
                this.backupDirectory = Path.Combine(this.backupDirectory, (title.Length > 24 ? title.Substring(0, 24) : title) + "_" + key.Substring(0, 8));
            }
            this.recordsDirectory = Path.Combine(this.dataDirectory, "记录");
            CheckLocalPath(this.baseDirectory);
        }

        private static void ValidateSummaryFileName(string name)
        {
            if (String.IsNullOrWhiteSpace(name) || name != Path.GetFileName(name) || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                !name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || name.StartsWith("~$", StringComparison.Ordinal) || name.Length > 150)
                throw new ArgumentException("请选择程序所在文件夹内的xlsx台账文件。");
        }

        public List<InventoryRecord> LoadLatest()
        {
            // Compatibility with older callers: history is never collapsed by device identity.
            return LoadRecords();
        }

        public List<InventoryRecord> LoadRecords()
        {
            using (FileStream guard = AcquireLock())
            {
                EnsureImportedIfNeeded();
                List<StoredRecord> all = ReadAll();
                return ActiveRecords(all, ReadDeleted(all));
            }
        }

        public SyncResult SynchronizeSummary()
        {
            using (FileStream guard = AcquireLock())
            {
                EnsureImportedIfNeeded();
                return SynchronizeCore(ReadAll());
            }
        }

        public SaveResult DeleteRecords(IList<string> storageIds)
        {
            if (storageIds == null || storageIds.Count == 0) throw new ArgumentException("请先选择要删除的记录。");
            using (FileStream guard = AcquireLock())
            {
                EnsureImportedIfNeeded();
                List<StoredRecord> all = ReadAll();
                SyncResult sync = SynchronizeCore(all);
                HashSet<string> deleted = ReadDeleted(all);
                HashSet<string> known = new HashSet<string>(StringComparer.Ordinal);
                foreach (StoredRecord item in all) known.Add(item.Record.StorageId);
                List<string> additions = new List<string>();
                foreach (string id in storageIds)
                {
                    if (String.IsNullOrEmpty(id) || !known.Contains(id)) throw new InvalidDataException("选中的记录已变化，请刷新后再删除。");
                    if (!deleted.Contains(id) && !additions.Contains(id)) additions.Add(id);
                }
                if (additions.Count == 0) return new SaveResult { Count = sync.Records.Count, SummaryUpdated = true, BackupUpdated = true, Message = "记录已经删除。" };
                string backup = CreateBackup(BackupDirectory); // Failure leaves every row active.
                foreach (string id in additions) deleted.Add(id);
                WriteDeleted(deleted);
                List<InventoryRecord> records = ActiveRecords(all, deleted);
                SaveResult result = new SaveResult { Count = records.Count, BackupPath = backup, BackupUpdated = true };
                if (sync.SummaryReadable)
                {
                    try { WriteSummary(records); result.SummaryUpdated = true; }
                    catch (Exception) { }
                }
                result.Message = "已删除 " + additions.Count + " 条，当前 " + records.Count + " 条。删除前已备份。" +
                    (result.SummaryUpdated ? "" : " 汇总表暂未更新，关闭Excel/WPS后可更新。") ;
                return result;
            }
        }

        private SyncResult SynchronizeCore(List<StoredRecord> all)
        {
            HashSet<string> deleted = ReadDeleted(all);
            SyncResult result = new SyncResult { Records = ActiveRecords(all, deleted), SummaryReadable = true, Message = "已同步当前记录。" };
            lastReadSummaryHash = null; lastSummaryObservedExists = File.Exists(SummaryPath);
            if (!lastSummaryObservedExists) { result.Message = "当前 " + result.Records.Count + " 条记录。"; return result; }
            try
            {
                WorkbookSnapshot snapshot = InventoryWorkbookReader.ReadSnapshot(SummaryPath);
                lastReadSummaryHash = snapshot.Sha256;
                HashSet<string> known = new HashSet<string>(StringComparer.Ordinal);
                foreach (StoredRecord item in all) known.Add(item.Record.StorageId);
                HashSet<string> missing = new HashSet<string>(StringComparer.Ordinal);
                if (snapshot.HasSyncIdentity)
                {
                    HashSet<string> importedBaseline = ReadImportedBaseline();
                    HashSet<string> baseline = new HashSet<string>(StringComparer.Ordinal);
                    foreach (string id in snapshot.ExportedIds)
                        if ((!known.Contains(id) && !importedBaseline.Contains(id)) || !baseline.Add(id)) throw new InvalidDataException("汇总表的记录标识不匹配，已保留原始数据。");
                    HashSet<string> visible = new HashSet<string>(StringComparer.Ordinal);
                    Dictionary<string, InventoryRecord> byId = new Dictionary<string, InventoryRecord>(StringComparer.Ordinal);
                    foreach (StoredRecord item in all) byId.Add(item.Record.StorageId, item.Record);
                    for (int i = 0; i < snapshot.StorageIds.Count; i++)
                    {
                        string id = snapshot.StorageIds[i];
                        if (String.IsNullOrEmpty(id) || !known.Contains(id) || !baseline.Contains(id) || !visible.Add(id)) throw new InvalidDataException("汇总表的记录标识缺失或重复，已保留原始数据。");
                        if (!String.Equals(RowFingerprint(snapshot, i), RecordFingerprint(snapshot, byId[id]), StringComparison.Ordinal))
                            throw new InvalidDataException("表格内容或排序与记录标识不一致，已停止同步，避免删除错误记录。");
                    }
                    foreach (string id in baseline) if (known.Contains(id) && !visible.Contains(id) && !deleted.Contains(id)) missing.Add(id);
                }
                else
                {
                    // Old exports have no reliable exported-row baseline. A complete match is safe
                    // to upgrade; unmatched rows could also be saves that Excel never accepted.
                    missing = MatchLegacyRows(snapshot, all, deleted);
                }
                if (missing.Count != 0)
                {
                    string backup = CreateBackup(BackupDirectory);
                    if (!SummaryHasHash(snapshot.Sha256)) throw new IOException("表格正在保存，稍后自动同步。");
                    foreach (string id in missing) deleted.Add(id);
                    WriteDeleted(deleted);
                    result.Records = ActiveRecords(all, deleted);
                    result.Changed = true;
                    result.Message = "已同步表格删除的 " + missing.Count + " 条记录，删除前已备份。";
                }
                if (!snapshot.HasSyncIdentity)
                {
                    if (CanReplaceSummary())
                    {
                        if (!SummaryHasHash(snapshot.Sha256)) throw new IOException("表格正在保存，稍后自动同步。");
                        WriteSummary(result.Records);
                        result.Message = result.Changed ? result.Message : "台账已升级，Excel删除并保存后自动同步。";
                    }
                    else result.Message = "关闭Excel/WPS一次，即可启用自动同步；当前记录已保留。";
                }
                return result;
            }
            catch (Exception ex)
            {
                result.SummaryReadable = false;
                result.Message = "表格暂未同步，当前记录已保留。" + ex.Message;
                return result;
            }
        }

        private bool SummaryHasHash(string expected)
        {
            if (!File.Exists(SummaryPath)) return false;
            using (FileStream input = new FileStream(SummaryPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                return String.Equals(HashStream(input), expected, StringComparison.OrdinalIgnoreCase);
        }

        private static List<InventoryRecord> ActiveRecords(List<StoredRecord> all, HashSet<string> deleted)
        {
            List<InventoryRecord> records = new List<InventoryRecord>();
            foreach (StoredRecord item in all) if (!deleted.Contains(item.Record.StorageId)) records.Add(item.Record);
            return records;
        }

        private HashSet<string> ReadDeleted(List<StoredRecord> all)
        {
            HashSet<string> deleted = new HashSet<string>(StringComparer.Ordinal);
            if (!File.Exists(StatePath)) return deleted;
            CheckLocalPath(StatePath);
            RecordState state;
            using (FileStream input = new FileStream(StatePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (input.Length > MaxManifestBytes) throw new InvalidDataException("记录状态文件过大，已停止操作。");
                state = (RecordState)new DataContractJsonSerializer(typeof(RecordState)).ReadObject(input);
            }
            if (state == null || state.FormatVersion != 1 || state.DeletedStorageIds == null) throw new InvalidDataException("记录状态无法读取，已停止操作。");
            HashSet<string> known = new HashSet<string>(StringComparer.Ordinal);
            foreach (StoredRecord item in all) known.Add(item.Record.StorageId);
            foreach (string id in state.DeletedStorageIds)
                if (!known.Contains(id) || !deleted.Add(id)) throw new InvalidDataException("记录状态与原始文件不一致，已停止操作。");
            return deleted;
        }

        private HashSet<string> ReadImportedBaseline()
        {
            if (!File.Exists(StatePath)) return new HashSet<string>(StringComparer.Ordinal);
            using (FileStream input = new FileStream(StatePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                RecordState state = (RecordState)new DataContractJsonSerializer(typeof(RecordState)).ReadObject(input);
                return new HashSet<string>(state.ImportedBaselineIds ?? new List<string>(), StringComparer.Ordinal);
            }
        }

        private List<string> ReadLegacyBaseline()
        {
            if (!File.Exists(StatePath)) return null;
            using (FileStream input = new FileStream(StatePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                return ((RecordState)new DataContractJsonSerializer(typeof(RecordState)).ReadObject(input)).LegacyBaselineIds;
        }

        private void WriteDeleted(HashSet<string> deleted) { WriteDeleted(deleted, ReadImportedBaseline(), ReadLegacyBaseline()); }

        private void WriteDeleted(HashSet<string> deleted, HashSet<string> importedBaseline) { WriteDeleted(deleted, importedBaseline, ReadLegacyBaseline()); }

        private void WriteDeleted(HashSet<string> deleted, HashSet<string> importedBaseline, List<string> legacyBaseline)
        {
            CheckLocalPath(StatePath);
            List<string> ids = new List<string>(deleted); ids.Sort(StringComparer.Ordinal);
            RecordState state = new RecordState { FormatVersion = 1, DeletedStorageIds = ids, ImportedBaselineIds = new List<string>(importedBaseline), LegacyBaselineIds = legacyBaseline };
            string temporary = Path.Combine(dataDirectory, ".状态_" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (FileStream output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                { new DataContractJsonSerializer(typeof(RecordState)).WriteObject(output, state); output.Flush(true); }
                MoveFileAtomically(temporary, StatePath);
            }
            finally { DeleteTemporary(temporary); }
        }

        private void EnsureImportedIfNeeded()
        {
            RecoverImportIfNeeded();
            if (Directory.GetFiles(recordsDirectory, "*.json", SearchOption.TopDirectoryOnly).Length != 0 || File.Exists(StatePath) || !File.Exists(SummaryPath)) return;
            WorkbookSnapshot snapshot = InventoryWorkbookReader.ReadSnapshot(SummaryPath);
            string originalBackup = Path.Combine(BackupDirectory, "导入前原表_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString("N") + ".xlsx");
            CheckLocalPath(originalBackup); Directory.CreateDirectory(BackupDirectory);
            using (FileStream input = new FileStream(SummaryPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (FileStream output = new FileStream(originalBackup, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { if (!String.Equals(HashStream(input), snapshot.Sha256, StringComparison.OrdinalIgnoreCase)) throw new IOException("表格正在保存，请稍后重试。"); input.Position = 0; input.CopyTo(output); output.Flush(true); }
            using (FileStream backup = new FileStream(originalBackup, FileMode.Open, FileAccess.Read, FileShare.Read))
                if (!String.Equals(HashStream(backup), snapshot.Sha256, StringComparison.OrdinalIgnoreCase)) throw new IOException("备份时表格发生变化，请稍后重新选择台账。");
            string staging = Path.Combine(dataDirectory, ".导入_" + Guid.NewGuid().ToString("N"));
            CheckLocalPath(staging); Directory.CreateDirectory(staging);
            List<string> importFiles = new List<string>();
            HashSet<long> sequences = new HashSet<long>();
            for (int i = 0; i < snapshot.Records.Count; i++)
            {
                InventoryRecord record = snapshot.Records[i];
                long sequence = i + 1;
                string filename = sequence.ToString("D20", CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString("N") + ".json";
                if (snapshot.HasSyncIdentity)
                {
                    filename = snapshot.StorageIds[i]; sequence = ValidateRecordFileName(filename);
                    if (!sequences.Add(sequence)) throw new InvalidDataException("表格中的记录标识重复，已停止导入。");
                }
                record.RecordId = Guid.NewGuid().ToString("N"); record.StorageId = filename;
                record.DeviceKey = "import-" + record.RecordId;
                ValidateRecord(record);
                StoredRecord item = new StoredRecord { FormatVersion = 2, SaveSequence = sequence, Record = record };
                using (FileStream output = new FileStream(Path.Combine(staging, filename), FileMode.CreateNew, FileAccess.Write, FileShare.None))
                { new DataContractJsonSerializer(typeof(StoredRecord)).WriteObject(output, item); output.Flush(true); }
                importFiles.Add(filename);
            }
            if (!SummaryHasHash(snapshot.Sha256)) throw new IOException("表格正在保存，已保留原表备份，请稍后重试。");
            ImportCommit commit = new ImportCommit { FormatVersion = 1, StagingDirectory = Path.GetFileName(staging), RecordFiles = importFiles, ImportedBaselineIds = snapshot.ExportedIds ?? new List<string>() };
            string temporary = Path.Combine(dataDirectory, ".导入提交_" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (FileStream output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                { new DataContractJsonSerializer(typeof(ImportCommit)).WriteObject(output, commit); output.Flush(true); }
                MoveFileAtomically(temporary, ImportCommitPath);
            }
            finally { DeleteTemporary(temporary); }
            RecoverImportIfNeeded();
        }

        private void RecoverImportIfNeeded()
        {
            if (!File.Exists(ImportCommitPath)) return;
            CheckLocalPath(ImportCommitPath);
            ImportCommit commit;
            using (FileStream input = new FileStream(ImportCommitPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (input.Length > MaxManifestBytes) throw new InvalidDataException("导入状态过大，已停止操作。");
                commit = (ImportCommit)new DataContractJsonSerializer(typeof(ImportCommit)).ReadObject(input);
            }
            Guid id;
            if (commit == null || commit.FormatVersion != 1 || commit.StagingDirectory == null ||
                !commit.StagingDirectory.StartsWith(".导入_", StringComparison.Ordinal) ||
                !Guid.TryParseExact(commit.StagingDirectory.Substring(4), "N", out id) || commit.RecordFiles == null || commit.RecordFiles.Count > MaxBackupRecords || commit.ImportedBaselineIds == null)
                throw new InvalidDataException("导入状态无效，已保留原始文件并停止操作。");
            HashSet<string> expected = new HashSet<string>(StringComparer.Ordinal);
            foreach (string file in commit.RecordFiles) { ValidateRecordFileName(file); if (!expected.Add(file)) throw new InvalidDataException("导入状态含重复记录。"); }
            foreach (string file in commit.ImportedBaselineIds) ValidateRecordFileName(file);
            string staging = Path.Combine(dataDirectory, commit.StagingDirectory); CheckLocalPath(staging);
            if (Directory.Exists(staging))
            {
                if (Directory.GetFileSystemEntries(recordsDirectory).Length != 0) throw new InvalidDataException("导入尚未完成且目标目录已有文件，已停止操作。");
                string[] stagedFiles = Directory.GetFiles(staging, "*.json", SearchOption.TopDirectoryOnly);
                if (stagedFiles.Length != expected.Count) throw new InvalidDataException("导入记录不完整，已停止操作。");
                foreach (string file in stagedFiles) if (!expected.Contains(Path.GetFileName(file))) throw new InvalidDataException("导入记录不匹配，已停止操作。");
                // The complete batch becomes visible in one directory rename; the journal allows
                // recovery if power is lost before its deletion state has been committed.
                Directory.Delete(recordsDirectory, false);
                Directory.Move(staging, recordsDirectory);
            }
            List<StoredRecord> imported = ReadAll();
            if (imported.Count != expected.Count) throw new InvalidDataException("导入提交的记录数不一致，已停止操作。");
            foreach (StoredRecord item in imported) if (!expected.Contains(item.Record.StorageId)) throw new InvalidDataException("导入提交的记录不一致，已停止操作。");
            if (!File.Exists(StatePath)) WriteDeleted(new HashSet<string>(StringComparer.Ordinal), new HashSet<string>(commit.ImportedBaselineIds, StringComparer.Ordinal), null);
            else ReadDeleted(imported);
            File.Delete(ImportCommitPath);
        }

        private bool CanReplaceSummary()
        {
            try { using (FileStream file = new FileStream(SummaryPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { } return true; }
            catch (IOException) { return false; }
            catch (UnauthorizedAccessException) { return false; }
        }

        private HashSet<string> MatchLegacyRows(WorkbookSnapshot snapshot, List<StoredRecord> all, HashSet<string> deleted)
        {
            List<string> baseline = ReadLegacyBaseline();
            List<InventoryRecord> records = new List<InventoryRecord>();
            if (baseline == null)
            {
                records = ActiveRecords(all, deleted);
                if (snapshot.Records.Count != records.Count) throw new InvalidDataException("旧表与已保存记录数量不同，请在历史记录中核对后删除；已停止覆盖旧表。");
            }
            else
            {
                HashSet<string> ids = new HashSet<string>(baseline, StringComparer.Ordinal);
                foreach (StoredRecord item in all) if (ids.Contains(item.Record.StorageId)) records.Add(item.Record);
                if (records.Count != ids.Count) throw new InvalidDataException("旧表同步基线与原始记录不一致，已停止同步。");
            }
            List<string> expected = new List<string>();
            foreach (InventoryRecord record in records) expected.Add(RecordFingerprint(snapshot, record));
            for (int i = 0; i < snapshot.Records.Count; i++)
            {
                string key = RowFingerprint(snapshot, i); int index = expected.IndexOf(key);
                if (index < 0) throw new InvalidDataException("旧表含手工修改，已停止覆盖。可另存为新台账后选择该文件。");
                expected.RemoveAt(index);
                records.RemoveAt(index);
            }
            if (baseline == null)
            {
                baseline = new List<string>();
                foreach (InventoryRecord record in ActiveRecords(all, deleted)) baseline.Add(record.StorageId);
                WriteDeleted(deleted, ReadImportedBaseline(), baseline);
            }
            HashSet<string> missing = new HashSet<string>(StringComparer.Ordinal);
            foreach (InventoryRecord record in records) if (!deleted.Contains(record.StorageId)) missing.Add(record.StorageId);
            return missing;
        }

        private static string RowFingerprint(WorkbookSnapshot snapshot, int row)
        {
            StringBuilder key = new StringBuilder();
            foreach (int column in SortedMappedColumns(snapshot)) AppendKey(key, snapshot.RawRows[row][column]);
            return key.ToString();
        }

        private static string RecordFingerprint(WorkbookSnapshot snapshot, InventoryRecord record)
        {
            // Compare displayed cell content once; do not format a parsed display value twice.
            bool modern = Array.IndexOf(snapshot.Headers, "物理MAC") >= 0;
            bool separateSite = snapshot.MappedFields.ContainsValue("Site");
            string[] values = VisibleValues(record, 1);
            StringBuilder key = new StringBuilder();
            foreach (int column in SortedMappedColumns(snapshot))
            {
                int visibleColumn = Array.IndexOf(VisibleHeaders, snapshot.Headers[column]);
                string value;
                // "位置" was also accepted beside a separate site column in older workbooks.
                if (separateSite && snapshot.MappedFields[column] == "Location") value = record.Location;
                else if (modern && visibleColumn >= 0) value = values[visibleColumn];
                else
                {
                    string property = snapshot.MappedFields[column];
                    value = (string)typeof(InventoryRecord).GetProperty(property).GetValue(record, null);
                    if (property == "CollectedAt") value = InventoryPresentation.CollectedTime(value);
                }
                AppendKey(key, value);
            }
            return key.ToString();
        }

        private static List<int> SortedMappedColumns(WorkbookSnapshot snapshot)
        { List<int> columns = new List<int>(snapshot.MappedFields.Keys); columns.Sort(); return columns; }

        private static void AppendKey(StringBuilder key, string text)
        {
            string value = (text ?? "").Replace("\r\n", "\n").Trim();
            key.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value);
        }

        public SaveResult Save(InventoryRecord record)
        {
            if (record == null) throw new ArgumentNullException("record", "没有可保存的采集结果。");
            // A save means a new row, including repeated saves of the very same object.
            record.RecordId = Guid.NewGuid().ToString("N");
            ValidateRecord(record);
            using (FileStream guard = AcquireLock())
            {
                EnsureImportedIfNeeded();
                List<StoredRecord> all = ReadAll();
                SyncResult sync = SynchronizeCore(all);
                long sequence = all.Count == 0 ? 1 : checked(all[all.Count - 1].SaveSequence + 1);
                StoredRecord item = new StoredRecord { FormatVersion = 2, SaveSequence = sequence, Record = record };
                string filename = sequence.ToString("D20", CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString("N") + ".json";
                record.StorageId = filename;
                string destination = Path.Combine(recordsDirectory, filename);
                string temporary = Path.Combine(recordsDirectory, ".saving_" + Guid.NewGuid().ToString("N") + ".tmp");
                try
                {
                    using (FileStream stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    {
                        new DataContractJsonSerializer(typeof(StoredRecord)).WriteObject(stream, item);
                        stream.Flush(true);
                    }
                    File.Move(temporary, destination);
                }
                catch (Exception ex)
                {
                    DeleteTemporary(temporary);
                    throw new IOException("本次数据未保存。请检查U盘是否插好、是否已满、文件夹是否允许写入，然后重试。" + Environment.NewLine + ex.Message, ex);
                }
                all.Add(item);
                List<InventoryRecord> records = ActiveRecords(all, ReadDeleted(all));
                SaveResult result = new SaveResult { Count = records.Count, BackupPath = String.Empty };
                string summaryError = null;
                string backupError = null;
                try
                {
                    if (!sync.SummaryReadable) throw new IOException(sync.Message);
                    WriteSummary(records);
                    result.SummaryUpdated = true;
                }
                catch (Exception ex) { summaryError = ex.Message; }
                // Independent of Excel availability: raw records are the recoverable source.
                try
                {
                    result.BackupPath = CreateBackup(BackupDirectory);
                    result.BackupUpdated = true;
                }
                catch (Exception ex) { backupError = ex.Message; }
                StringBuilder message = new StringBuilder("已追加一条，当前 " + records.Count + " 条记录。");
                message.AppendLine();
                message.Append(result.SummaryUpdated ? "汇总表已更新。" : "汇总表暂未更新（可能在Excel或WPS中打开）。关闭表格后点击“查看 / 更新总表”；不要为更新总表而重复保存。");
                message.AppendLine();
                message.Append(result.BackupUpdated ? "已生成一份新的完整备份，旧备份保留。" : "本次自动备份失败；原始记录已保存，但本次尚无新的完整备份。请检查磁盘空间，使用“另存备份”再备份一次。");
                if (summaryError != null) message.AppendLine().Append("汇总详情：").Append(summaryError);
                if (backupError != null) message.AppendLine().Append("备份详情：").Append(backupError);
                result.Message = message.ToString();
                return result;
            }
        }

        public string RebuildSummary()
        {
            using (FileStream guard = AcquireLock())
            {
                EnsureImportedIfNeeded();
                SyncResult sync = SynchronizeCore(ReadAll());
                if (!sync.SummaryReadable) throw new IOException(sync.Message);
                List<InventoryRecord> records = sync.Records;
                try { WriteSummary(records); }
                catch (Exception ex) { throw new IOException("汇总表未更新，已保存的记录仍保留。请关闭Excel或WPS中的汇总表，检查U盘，再重试。" + Environment.NewLine + ex.Message, ex); }
                return SummaryPath;
            }
        }

        public string BackupNow(string destinationDirectory)
        {
            if (String.IsNullOrWhiteSpace(destinationDirectory)) throw new ArgumentException("请选择本机或另一只U盘中的备份文件夹。");
            CheckLocalPath(destinationDirectory);
            using (FileStream guard = AcquireLock())
            {
                EnsureImportedIfNeeded();
                SynchronizeCore(ReadAll());
                return CreateBackup(Path.GetFullPath(destinationDirectory));
            }
        }

        private string CreateBackup(string destinationDirectory)
        {
            CheckLocalPath(destinationDirectory);
            Directory.CreateDirectory(destinationDirectory);
            string[] files = Directory.GetFiles(recordsDirectory, "*.json", SearchOption.TopDirectoryOnly);
            if (files.Length > MaxBackupRecords) throw new IOException("记录超过完整备份支持的20000条，请保留原始数据并联系维护人员。");
            Array.Sort(files, StringComparer.Ordinal);
            BackupManifest manifest = new BackupManifest {
                FormatVersion = 2, CreatedUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                Instructions = "完整备份：records内是每次保存的原始记录，包含所有历史行；不按设备合并。manifest.json包含长度和SHA-256校验值。此备份不含Excel汇总表，恢复时依据原始记录重新生成总表。使用程序的恢复功能，恢复到全新文件夹，现有数据不覆盖。U盘中的备份不能防止U盘遗失或损坏，请另存一份到单位允许的其他本地磁盘或U盘。",
                Records = new List<BackupFile>(), DeletedStorageIds = new List<string>(ReadDeleted(ReadAll())), SummaryFileName = summaryFileName
            };
            string destination = Path.Combine(destinationDirectory, "台账备份_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString("N") + ".zip");
            string temporary = Path.Combine(destinationDirectory, ".backup_" + Guid.NewGuid().ToString("N") + ".tmp");
            long total = 0;
            try
            {
                using (FileStream output = new FileStream(temporary, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                {
                    using (ZipArchive zip = new ZipArchive(output, ZipArchiveMode.Create, true))
                    {
                        foreach (string file in files)
                        {
                            CheckLocalPath(file);
                            string name = Path.GetFileName(file);
                            ValidateRecordFileName(name);
                            using (FileStream input = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
                            {
                                if (input.Length > MaxRecordBytes || (total += input.Length) > MaxBackupBytes) throw new IOException("记录大小超过完整备份限制，请保留原始数据并联系维护人员。");
                                string hash = HashStream(input);
                                input.Position = 0;
                                string entryName = "records/" + name;
                                using (Stream entry = zip.CreateEntry(entryName, CompressionLevel.Optimal).Open()) input.CopyTo(entry);
                                manifest.Records.Add(new BackupFile { Path = entryName, Length = input.Length, Sha256 = hash });
                            }
                        }
                        using (Stream entry = zip.CreateEntry("manifest.json", CompressionLevel.Optimal).Open())
                            new DataContractJsonSerializer(typeof(BackupManifest)).WriteObject(entry, manifest);
                    }
                    output.Flush(true);
                }
                CheckLocalPath(destination);
                File.Move(temporary, destination); // Unique path, never replace an older backup.
                return destination;
            }
            finally { DeleteTemporary(temporary); }
        }

        public string RestoreBackupToNewFolder(string zipPath, string destinationParent)
        {
            if (String.IsNullOrWhiteSpace(zipPath) || String.IsNullOrWhiteSpace(destinationParent)) throw new ArgumentException("请选择备份文件和用于存放恢复副本的本地文件夹。");
            CheckLocalPath(zipPath);
            CheckLocalPath(destinationParent);
            string restoredDirectory = null;
            try
            {
                using (FileStream input = new FileStream(Path.GetFullPath(zipPath), FileMode.Open, FileAccess.Read, FileShare.Read))
                using (ZipArchive zip = new ZipArchive(input, ZipArchiveMode.Read))
                {
                    BackupManifest manifest = ValidateBackup(zip);
                    string parent = Path.GetFullPath(destinationParent);
                    Directory.CreateDirectory(parent);
                    restoredDirectory = Path.Combine(parent, "台账恢复_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString("N").Substring(0, 8));
                    CheckLocalPath(restoredDirectory);
                    if (Directory.Exists(restoredDirectory) || File.Exists(restoredDirectory)) throw new IOException("恢复目录意外已存在，请重试以创建新的目录。");
                    InventoryStore restored = new InventoryStore(restoredDirectory, String.IsNullOrEmpty(manifest.SummaryFileName) ? "电脑台账汇总.xlsx" : manifest.SummaryFileName);
                    string restoredRecords = restored.recordsDirectory;
                    Directory.CreateDirectory(restoredRecords);
                    foreach (BackupFile file in manifest.Records)
                    {
                        string target = Path.Combine(restoredRecords, file.Path.Substring("records/".Length));
                        CheckLocalPath(target);
                        using (Stream source = zip.GetEntry(file.Path).Open())
                        using (FileStream output = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                        {
                            CopyBounded(source, output, file.Length);
                            output.Flush(true);
                        }
                    }
                    restored.WriteDeleted(new HashSet<string>(manifest.DeletedStorageIds ?? new List<string>(), StringComparer.Ordinal));
                    if (restored.LoadRecords().Count != manifest.Records.Count - (manifest.DeletedStorageIds == null ? 0 : manifest.DeletedStorageIds.Count)) throw new InvalidDataException("恢复后的记录数不一致。");
                    restored.RebuildSummary();
                    File.WriteAllText(Path.Combine(restoredDirectory, "数据", "当前台账.txt"), restored.SummaryFileName, new UTF8Encoding(false));
                    return restoredDirectory;
                }
            }
            catch (Exception ex)
            {
                string detail = restoredDirectory == null ? "备份未通过验证，未恢复任何数据。" : "恢复未完成；已写入的恢复副本保留在：" + restoredDirectory;
                throw new IOException(detail + Environment.NewLine + "当前工具里的原有记录和汇总表没有被修改。" + Environment.NewLine + ex.Message, ex);
            }
        }

        private static BackupManifest ValidateBackup(ZipArchive zip)
        {
            if (zip.Entries.Count < 1 || zip.Entries.Count > MaxBackupRecords + 1) throw new InvalidDataException("备份条目数量不符合限制。");
            Dictionary<string, ZipArchiveEntry> entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            long total = 0;
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                string name = entry.FullName;
                if (name != "manifest.json")
                {
                    if (!name.StartsWith("records/", StringComparison.Ordinal)) throw new InvalidDataException("备份包含不允许的路径或文件。");
                    ValidateRecordFileName(name.Substring("records/".Length));
                }
                if (entries.ContainsKey(name)) throw new InvalidDataException("备份包含重复的文件路径。");
                if (entry.Length < 0 || entry.Length > (name == "manifest.json" ? MaxManifestBytes : MaxRecordBytes)) throw new InvalidDataException("备份单个文件过大。");
                if (name != "manifest.json" && (total += entry.Length) > MaxBackupBytes) throw new InvalidDataException("备份展开后的大小超过限制。");
                // No attributes are restored. Reject symbolic-link metadata where the runtime exposes it.
                System.Reflection.PropertyInfo attributes = typeof(ZipArchiveEntry).GetProperty("ExternalAttributes");
                if (attributes != null)
                {
                    int value = (int)attributes.GetValue(entry, null);
                    if ((value & (int)FileAttributes.ReparsePoint) != 0 || ((value >> 16) & 0xF000) == 0xA000)
                        throw new InvalidDataException("备份包含联接或符号链接条目。");
                }
                entries.Add(name, entry);
            }
            ZipArchiveEntry manifestEntry;
            if (!entries.TryGetValue("manifest.json", out manifestEntry)) throw new InvalidDataException("缺少备份清单。");
            BackupManifest manifest;
            using (Stream entry = manifestEntry.Open())
            using (MemoryStream bounded = ReadBounded(entry, MaxManifestBytes))
                manifest = (BackupManifest)new DataContractJsonSerializer(typeof(BackupManifest)).ReadObject(bounded);
            if (manifest == null || (manifest.FormatVersion != 1 && manifest.FormatVersion != 2) || manifest.Records == null || manifest.Records.Count != entries.Count - 1)
                throw new InvalidDataException("备份清单版本或记录数量不匹配。");
            if (!String.IsNullOrEmpty(manifest.SummaryFileName)) ValidateSummaryFileName(manifest.SummaryFileName);
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<long> sequences = new HashSet<long>();
            foreach (BackupFile file in manifest.Records)
            {
                ZipArchiveEntry entry;
                if (file == null || file.Path == null || !file.Path.StartsWith("records/", StringComparison.Ordinal) || !seen.Add(file.Path) || !entries.TryGetValue(file.Path, out entry))
                    throw new InvalidDataException("备份清单中的记录路径无效或重复。");
                long sequence = ValidateRecordFileName(file.Path.Substring("records/".Length));
                if (!String.Equals(file.Path, entry.FullName, StringComparison.Ordinal) || file.Length != entry.Length || !IsHexHash(file.Sha256)) throw new InvalidDataException("备份清单校验信息无效。");
                using (Stream source = entry.Open())
                using (MemoryStream data = ReadBounded(source, MaxRecordBytes))
                {
                    if (data.Length != file.Length || !String.Equals(HashStream(data), file.Sha256, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("备份记录校验失败：" + file.Path);
                    data.Position = 0;
                    StoredRecord item = (StoredRecord)new DataContractJsonSerializer(typeof(StoredRecord)).ReadObject(data);
                    ValidateStoredRecord(item, sequence);
                    if (!sequences.Add(item.SaveSequence)) throw new InvalidDataException("备份包含重复的保存序号。");
                }
            }
            HashSet<string> deleted = new HashSet<string>(StringComparer.Ordinal);
            foreach (string id in manifest.DeletedStorageIds ?? new List<string>())
                if (!seen.Contains("records/" + id) || !deleted.Add(id)) throw new InvalidDataException("备份中的删除状态不匹配。");
            return manifest;
        }

        private static MemoryStream ReadBounded(Stream input, long limit)
        {
            MemoryStream output = new MemoryStream();
            try { CopyBounded(input, output, limit); output.Position = 0; return output; }
            catch { output.Dispose(); throw; }
        }

        private static void CopyBounded(Stream input, Stream output, long limit)
        {
            byte[] buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += read;
                if (total > limit) throw new InvalidDataException("备份文件解压大小超过允许值。");
                output.Write(buffer, 0, read);
            }
        }

        private static string HashStream(Stream input)
        {
            using (SHA256 hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(input)).Replace("-", String.Empty).ToLowerInvariant();
        }

        private static bool IsHexHash(string value)
        {
            if (value == null || value.Length != 64) return false;
            foreach (char c in value) if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))) return false;
            return true;
        }

        private FileStream AcquireLock()
        {
            CheckLocalPath(baseDirectory);
            CheckLocalPath(dataDirectory);
            CheckLocalPath(recordsDirectory);
            try
            {
                Directory.CreateDirectory(recordsDirectory);
                string lockPath = Path.Combine(dataDirectory, "保存锁.lock");
                CheckLocalPath(lockPath);
                return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (Exception ex)
            {
                throw new IOException("无法打开数据文件夹。请确认U盘已插好、文件夹可以写入，并关闭同一文件夹中重复打开的采集程序。" + Environment.NewLine + ex.Message, ex);
            }
        }

        private List<StoredRecord> ReadAll()
        {
            CheckLocalPath(recordsDirectory);
            string[] incomplete = Directory.GetFiles(recordsDirectory, ".saving_*.tmp", SearchOption.TopDirectoryOnly);
            if (incomplete.Length != 0)
                throw new InvalidDataException("发现上次保存中断留下的文件：“" + Path.GetFileName(incomplete[0]) + "”。这次记录可能尚未保存完整，已停止操作。请保留整个工具文件夹以便检查，不要删除数据文件。");
            List<StoredRecord> all = new List<StoredRecord>();
            HashSet<long> sequences = new HashSet<long>();
            string[] files = Directory.GetFiles(recordsDirectory, "*.json", SearchOption.TopDirectoryOnly);
            foreach (string file in files)
            {
                try
                {
                    CheckLocalPath(file);
                    StoredRecord item;
                    using (FileStream stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        if (stream.Length > MaxRecordBytes) throw new InvalidDataException("单条原始记录超过允许的大小。");
                        item = (StoredRecord)new DataContractJsonSerializer(typeof(StoredRecord)).ReadObject(stream);
                    }
                    string name = Path.GetFileName(file);
                    long fileSequence = ValidateRecordFileName(name);
                    ValidateStoredRecord(item, fileSequence);
                    item.Record.StorageId = name;
                    if (!sequences.Add(item.SaveSequence)) throw new InvalidDataException("发现重复的保存序号；请保留文件并检查是否手工混合了两份工具文件夹。");
                    all.Add(item);
                }
                catch (Exception ex)
                {
                    throw new InvalidDataException("发现无法读取的记录：“" + Path.GetFileName(file) + "”。为避免漏掉电脑，已停止操作，原有记录和汇总表不会被覆盖。请保留整个工具文件夹以便检查。" + Environment.NewLine + ex.Message, ex);
                }
            }
            all.Sort(delegate(StoredRecord a, StoredRecord b) { return a.SaveSequence.CompareTo(b.SaveSequence); });
            return all;
        }

        private static long ValidateRecordFileName(string name)
        {
            long sequence;
            Guid id;
            if (name == null || name.Length != 58 || name[20] != '_' || !name.EndsWith(".json", StringComparison.Ordinal) ||
                !Int64.TryParse(name.Substring(0, 20), NumberStyles.None, CultureInfo.InvariantCulture, out sequence) || sequence <= 0 ||
                !Guid.TryParseExact(name.Substring(21, 32), "N", out id))
                throw new InvalidDataException("记录文件名不符合格式。");
            return sequence;
        }

        private static void ValidateStoredRecord(StoredRecord item, long fileSequence)
        {
            if (item == null || (item.FormatVersion != 1 && item.FormatVersion != 2) || item.SaveSequence <= 0 || item.SaveSequence != fileSequence ||
                item.Record == null || String.IsNullOrWhiteSpace(item.Record.RecordId)) throw new InvalidDataException("记录内容缺失或版本不匹配。");
            ValidateRecord(item.Record);
        }

        private static List<InventoryRecord> AllRecords(List<StoredRecord> all)
        {
            List<InventoryRecord> result = new List<InventoryRecord>();
            foreach (StoredRecord item in all) result.Add(item.Record);
            return result;
        }

        private static void ValidateRecord(InventoryRecord record)
        {
            foreach (System.Reflection.PropertyInfo property in typeof(InventoryRecord).GetProperties())
            {
                if (property.PropertyType != typeof(string)) continue;
                string value = (string)property.GetValue(record, null);
                if (value == null) continue;
                if (value.Length > 32767) throw new InvalidDataException("字段“" + property.Name + "”过长，超过Excel单元格容量，请缩短后保存。");
                try { XmlConvert.VerifyXmlChars(value); }
                catch (XmlException ex) { throw new InvalidDataException("字段“" + property.Name + "”含不能写入Excel的控制字符，请修改后保存。", ex); }
            }
        }

        private static void CheckLocalPath(string path)
        {
            string fullPath = Path.GetFullPath(path);
            if (fullPath.StartsWith(@"\\", StringComparison.Ordinal) || fullPath.StartsWith("//", StringComparison.Ordinal))
                throw new IOException("为了使资料留在本机或U盘，不能使用网络共享路径。");
            string root = Path.GetPathRoot(fullPath);
            if (String.IsNullOrEmpty(root)) throw new IOException("请选择本机或U盘中的文件夹。");
            DriveInfo drive = new DriveInfo(root);
            if (drive.DriveType == DriveType.Network) throw new IOException("不能把数据保存在映射的网络盘，请使用本机或U盘。");
            string current = fullPath;
            while (!String.IsNullOrEmpty(current))
            {
                if (Directory.Exists(current) || File.Exists(current))
                {
                    if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                        throw new IOException("不能使用联接、符号链接或重定向文件夹，请将完整工具文件夹放到普通本地目录或U盘。");
                }
                string parent = Path.GetDirectoryName(current.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                if (String.Equals(parent, current, StringComparison.OrdinalIgnoreCase)) break;
                current = parent;
            }
        }

        private void WriteSummary(List<InventoryRecord> records)
        {
            if (records.Count > 1048575) throw new InvalidDataException("记录数量超过Excel单个工作表容量。");
            CheckLocalPath(SummaryPath);
            string temporary = Path.Combine(baseDirectory, ".台账汇总_" + Guid.NewGuid().ToString("N") + ".tmp");
            string writtenHash = null;
            try
            {
                using (FileStream output = new FileStream(temporary, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                {
                    using (ZipArchive zip = new ZipArchive(output, ZipArchiveMode.Create, true))
                    {
                        WritePart(zip, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/></Types>");
                        WritePart(zip, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                        WritePart(zip, "xl/workbook.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"" + SpreadsheetNs + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"电脑台账\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"" + SyncSheetName + "\" sheetId=\"2\" state=\"veryHidden\" r:id=\"rId3\"/></sheets></workbook>");
                        WritePart(zip, "xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/></Relationships>");
                        WritePart(zip, "xl/styles.xml", Styles());
                        ZipArchiveEntry sheet = zip.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Optimal);
                        using (Stream stream = sheet.Open())
                        using (XmlWriter xml = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false }))
                            WriteWorksheet(xml, records);
                        using (Stream stream = zip.CreateEntry("xl/worksheets/sheet2.xml", CompressionLevel.Optimal).Open())
                        using (XmlWriter xml = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false }))
                            WriteSyncIndex(xml, records);
                    }
                    output.Flush(true);
                    output.Position = 0; writtenHash = HashStream(output);
                }
                EnsureSummaryUnchanged();
                if (File.Exists(SummaryPath))
                {
                    // MoveFileEx performs same-volume replace without DeleteFile, including USB
                    // filesystems where ReplaceFile's metadata merging is not available.
                    BackupSummary();
                    EnsureSummaryUnchanged();
                    MoveFileAtomically(temporary, SummaryPath);
                }
                else MoveFileAtomically(temporary, SummaryPath);
                lastReadSummaryHash = writtenHash; lastSummaryObservedExists = true;
            }
            finally { DeleteTemporary(temporary); }
        }

        private void EnsureSummaryUnchanged()
        {
            if (lastSummaryObservedExists ? lastReadSummaryHash == null || !SummaryHasHash(lastReadSummaryHash) : File.Exists(SummaryPath))
                throw new IOException("表格刚刚发生变化，已保留现有表格，请稍后再更新。");
        }

        private void BackupSummary()
        {
            string backupDirectory = Path.Combine(dataDirectory, "汇总备份");
            CheckLocalPath(backupDirectory);
            Directory.CreateDirectory(backupDirectory);
            string backupPath = Path.Combine(backupDirectory, "电脑台账汇总_上一版.xlsx");
            CheckLocalPath(backupPath);
            string temporary = Path.Combine(backupDirectory, ".backup_" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                // An exclusive Excel lock fails here before either workbook is replaced.
                File.Copy(SummaryPath, temporary, false);
                using (FileStream output = new FileStream(temporary, FileMode.Open, FileAccess.Write, FileShare.None)) output.Flush(true);
                MoveFileAtomically(temporary, backupPath);
            }
            finally { DeleteTemporary(temporary); }
        }

        private static void MoveFileAtomically(string source, string destination)
        {
            if (!String.Equals(Path.GetDirectoryName(source), Path.GetDirectoryName(destination), StringComparison.OrdinalIgnoreCase))
                throw new IOException("临时文件必须与目标表格保存在同一文件夹。");
            if (!MoveFileExW(source, destination, MoveReplaceExisting | MoveWriteThrough))
            {
                int error = Marshal.GetLastWin32Error();
                throw new IOException("无法更新表格，原表格保持不变。" + new Win32Exception(error).Message, new Win32Exception(error));
            }
        }

        private static void DeleteTemporary(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private static void WritePart(ZipArchive zip, string name, string value)
        {
            using (Stream stream = zip.CreateEntry(name, CompressionLevel.Optimal).Open())
            using (StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(false))) writer.Write(value);
        }

        private static string Styles()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"" + SpreadsheetNs + "\"><fonts count=\"2\"><font><sz val=\"11\"/><name val=\"等线\"/></font><font><b/><sz val=\"11\"/><color rgb=\"FFFFFFFF\"/><name val=\"等线\"/></font></fonts><fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF24577A\"/><bgColor indexed=\"64\"/></patternFill></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"3\"><xf numFmtId=\"49\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\" wrapText=\"1\"/></xf><xf numFmtId=\"49\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"center\" wrapText=\"1\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"top\" horizontal=\"center\"/></xf></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";
        }

        private static void WriteWorksheet(XmlWriter xml, List<InventoryRecord> records)
        {
            int visibleCount = VisibleHeaders.Length;
            string[] headers = new string[visibleCount + 1]; Array.Copy(VisibleHeaders, headers, visibleCount); headers[visibleCount] = StorageHeader;
            double[] widths = { 8, 18, 14, 28, 24, 62, 24, 36, 14, 30, 32, 16, 14, 16, 8, 30, 22, 30, 80 };
            string extent = "A1:" + ((char)('A' + visibleCount)).ToString() + (records.Count + 1).ToString(CultureInfo.InvariantCulture);
            xml.WriteStartDocument(true);
            xml.WriteStartElement("worksheet", SpreadsheetNs);
            xml.WriteStartElement("dimension"); xml.WriteAttributeString("ref", extent); xml.WriteEndElement();
            xml.WriteStartElement("sheetViews"); xml.WriteStartElement("sheetView"); xml.WriteAttributeString("workbookViewId", "0");
            xml.WriteStartElement("pane"); xml.WriteAttributeString("xSplit", "2"); xml.WriteAttributeString("ySplit", "1"); xml.WriteAttributeString("topLeftCell", "C2"); xml.WriteAttributeString("activePane", "bottomRight"); xml.WriteAttributeString("state", "frozen"); xml.WriteEndElement();
            xml.WriteEndElement(); xml.WriteEndElement();
            xml.WriteStartElement("sheetFormatPr"); xml.WriteAttributeString("defaultRowHeight", "30"); xml.WriteEndElement();
            xml.WriteStartElement("cols");
            for (int i = 0; i < widths.Length; i++)
            {
                xml.WriteStartElement("col"); xml.WriteAttributeString("min", (i + 1).ToString(CultureInfo.InvariantCulture)); xml.WriteAttributeString("max", (i + 1).ToString(CultureInfo.InvariantCulture)); xml.WriteAttributeString("width", widths[i].ToString(CultureInfo.InvariantCulture)); xml.WriteAttributeString("customWidth", "1"); if (i == visibleCount) xml.WriteAttributeString("hidden", "1"); xml.WriteEndElement();
            }
            xml.WriteEndElement();
            xml.WriteStartElement("sheetData");
            WriteRow(xml, 1, headers, true, widths);
            for (int i = 0; i < records.Count; i++)
            {
                InventoryRecord r = records[i];
                string[] values = new string[visibleCount + 1]; Array.Copy(VisibleValues(r, i + 1), values, visibleCount); values[visibleCount] = r.StorageId;
                WriteRow(xml, i + 2, values, false, widths);
            }
            xml.WriteEndElement();
            xml.WriteStartElement("autoFilter"); xml.WriteAttributeString("ref", extent); xml.WriteEndElement();
            xml.WriteStartElement("pageMargins"); xml.WriteAttributeString("left", "0.25"); xml.WriteAttributeString("right", "0.25"); xml.WriteAttributeString("top", "0.4"); xml.WriteAttributeString("bottom", "0.4"); xml.WriteAttributeString("header", "0.2"); xml.WriteAttributeString("footer", "0.2"); xml.WriteEndElement();
            xml.WriteStartElement("pageSetup"); xml.WriteAttributeString("orientation", "landscape"); xml.WriteAttributeString("paperSize", "9"); xml.WriteEndElement();
            xml.WriteEndElement(); xml.WriteEndDocument();
        }

        private static string[] VisibleValues(InventoryRecord r, int sequence)
        {
            return new string[] { sequence.ToString(CultureInfo.InvariantCulture), r.DeviceNumber, r.DeviceType, r.ManufacturerModel, InventoryPresentation.SerialText(r.BiosSerial), InventoryPresentation.DiskText(r), InventoryPresentation.MacText(r), r.Cpu, r.Memory, r.OperatingSystem, InventoryPresentation.LocationText(r), r.AssignedTo, r.DeviceStatus, r.AcquiredDate, "1", r.Remarks, InventoryPresentation.CollectedTime(r.CollectedAt), InventoryPresentation.VerificationText(r) };
        }

        private static void WriteSyncIndex(XmlWriter xml, List<InventoryRecord> records)
        {
            xml.WriteStartDocument(true); xml.WriteStartElement("worksheet", SpreadsheetNs); xml.WriteStartElement("sheetData");
            for (int i = 0; i <= records.Count; i++)
            {
                xml.WriteStartElement("row"); xml.WriteAttributeString("r", (i + 1).ToString(CultureInfo.InvariantCulture));
                xml.WriteStartElement("c"); xml.WriteAttributeString("r", "A" + (i + 1).ToString(CultureInfo.InvariantCulture)); xml.WriteAttributeString("t", "inlineStr");
                xml.WriteStartElement("is"); xml.WriteElementString("t", i == 0 ? SyncMarker : records[i - 1].StorageId); xml.WriteEndElement(); xml.WriteEndElement(); xml.WriteEndElement();
            }
            xml.WriteEndElement(); xml.WriteEndElement(); xml.WriteEndDocument();
        }

        private static void WriteRow(XmlWriter xml, int rowNumber, string[] values, bool header, double[] widths)
        {
            xml.WriteStartElement("row"); xml.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
            if (header) { xml.WriteAttributeString("ht", "36"); xml.WriteAttributeString("customHeight", "1"); }
            else
            {
                int lines = 1;
                for (int column = 0; column < values.Length; column++)
                {
                    int cellLines = 0;
                    foreach (string part in (values[column] ?? "").Replace("\r", "").Split('\n'))
                    {
                        double units = 0;
                        foreach (char ch in part) units += ch >= 0x2E80 ? 2 : 1;
                        cellLines += Math.Max(1, (int)Math.Ceiling(units / Math.Max(8, widths[column] - 3)));
                    }
                    lines = Math.Max(lines, cellLines);
                }
                xml.WriteAttributeString("ht", Math.Min(409, Math.Max(32, 10 + lines * 16)).ToString(CultureInfo.InvariantCulture));
                xml.WriteAttributeString("customHeight", "1");
            }
            for (int i = 0; i < values.Length; i++)
            {
                xml.WriteStartElement("c"); xml.WriteAttributeString("r", ((char)('A' + i)).ToString() + rowNumber.ToString(CultureInfo.InvariantCulture));
                bool numeric = !header && i < VisibleHeaders.Length && (VisibleHeaders[i] == "序号" || VisibleHeaders[i] == "数量");
                xml.WriteAttributeString("s", header ? "1" : (numeric ? "2" : "0"));
                if (numeric) xml.WriteElementString("v", values[i]);
                else
                {
                    // inlineStr is always text, even when it begins with =, +, - or @.
                    xml.WriteAttributeString("t", "inlineStr"); xml.WriteStartElement("is"); xml.WriteStartElement("t"); xml.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve"); xml.WriteString(values[i] ?? String.Empty); xml.WriteEndElement(); xml.WriteEndElement();
                }
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }
    }
}
