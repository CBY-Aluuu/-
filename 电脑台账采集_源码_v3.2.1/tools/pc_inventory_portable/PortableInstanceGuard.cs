using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace PcInventory.Ios
{
    // One interactive app per portable folder. The kernel releases this lock if the
    // process crashes; no lock file, device information or network connection is used.
    internal sealed class PortableInstanceGuard : IDisposable
    {
        private Mutex mutex;
        private const string WindowMarker = "CBY.PcInventory.ActiveInstance.v1";
        public bool IsOwner { get; private set; }
        internal bool RecoveredAbandonedLock { get; private set; }

        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ShowWindowAsync(IntPtr window, int command);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr window);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool SetProp(IntPtr window, string name, IntPtr data);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetProp(IntPtr window, string name);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr RemoveProp(IntPtr window, string name);

        private PortableInstanceGuard() { }

        public static PortableInstanceGuard Acquire(string directory)
        {
            string key;
            using (SHA256 sha = SHA256.Create())
                key = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(NormalizeDirectory(directory)))).Replace("-", "");
            var guard = new PortableInstanceGuard();
            try
            {
                guard.mutex = new Mutex(false, "Local\\CBY.PcInventory.Directory." + key);
                try { guard.IsOwner = guard.mutex.WaitOne(0, false); }
                catch (AbandonedMutexException) { guard.IsOwner = true; guard.RecoveredAbandonedLock = true; }
                return guard;
            }
            catch { guard.Dispose(); throw; }
        }

        internal static string NormalizeDirectory(string directory)
        {
            if (String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("程序所在文件夹无效。");
            return Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).ToUpperInvariant();
        }

        public static void MarkWindow(IntPtr window)
        {
            if (window == IntPtr.Zero || !SetProp(window, WindowMarker, new IntPtr(1)))
                throw new InvalidOperationException("程序窗口初始化失败，请重新打开。");
        }

        public static void UnmarkWindow(IntPtr window)
        { if (window != IntPtr.Zero) RemoveProp(window, WindowMarker); }

        // Only the owner checks for old releases. New releases participate in the mutex,
        // so simultaneously launched secondary processes cannot make the owner exit.
        public static bool ActivateLegacyInstance(string directory)
        {
            List<int> candidates = FindCandidates(directory, true);
            if (candidates.Count == 0) return false;
            ActivateCandidates(candidates, 1200);
            return true;
        }

        internal static bool HasLegacyInstance(string directory)
        { return FindCandidates(directory, true).Count != 0; }

        public static bool ActivateExistingInstance(string directory)
        {
            DateTime deadline = DateTime.UtcNow.AddMilliseconds(2500);
            do
            {
                if (ActivateCandidates(FindCandidates(directory, false), 0)) return true;
                Thread.Sleep(100);
            }
            while (DateTime.UtcNow < deadline);
            return false;
        }

        private static List<int> FindCandidates(string directory, bool legacyOnly)
        {
            string wanted = NormalizeDirectory(directory);
            int current, session;
            using (Process own = Process.GetCurrentProcess()) { current = own.Id; session = own.SessionId; }
            var found = new List<int>();
            foreach (Process process in Process.GetProcesses())
            {
                using (process)
                {
                    try
                    {
                        if (process.Id == current || process.SessionId != session) continue;
                        ProcessModule module = process.MainModule;
                        if (module == null || !String.Equals(NormalizeDirectory(Path.GetDirectoryName(module.FileName)), wanted, StringComparison.Ordinal)) continue;
                        string filename = Path.GetFileNameWithoutExtension(module.FileName);
                        bool collector = filename.StartsWith("电脑台账采集", StringComparison.OrdinalIgnoreCase);
                        if (!collector)
                        {
                            FileVersionInfo info = module.FileVersionInfo;
                            collector = (info.ProductName ?? "").StartsWith("电脑台账采集", StringComparison.Ordinal) ||
                                String.Equals(info.FileDescription, "电脑台账采集", StringComparison.Ordinal);
                        }
                        if (!collector) continue;
                        if (legacyOnly)
                        {
                            // The on-disk executable may already have been upgraded while an
                            // older image is still running. Only a live window identifies the
                            // version participating in this mutex; disk version metadata cannot.
                            IntPtr window = process.MainWindowHandle;
                            if (window == IntPtr.Zero || GetProp(window, WindowMarker) == new IntPtr(1)) continue;
                        }
                        found.Add(process.Id);
                    }
                    catch (Exception ex)
                    {
                        // Other users' or elevated processes may be unreadable. They are not
                        // guessed to be this app, and their details are never shown or logged.
                        if (!(ex is System.ComponentModel.Win32Exception) && !(ex is InvalidOperationException) &&
                            !(ex is NotSupportedException) && !(ex is ArgumentException)) throw;
                    }
                }
            }
            return found;
        }

        private static bool ActivateCandidates(List<int> candidates, int waitMilliseconds)
        {
            DateTime deadline = DateTime.UtcNow.AddMilliseconds(waitMilliseconds);
            do
            {
                foreach (int id in candidates)
                {
                    try
                    {
                        using (Process process = Process.GetProcessById(id))
                        {
                            process.Refresh();
                            IntPtr window = process.MainWindowHandle;
                            if (window == IntPtr.Zero) continue;
                            ShowWindowAsync(window, IsIconic(window) ? 9 : 5);
                            BringWindowToTop(window);
                            SetForegroundWindow(window);
                            return true;
                        }
                    }
                    catch (ArgumentException) { }
                    catch (InvalidOperationException) { }
                    catch (System.ComponentModel.Win32Exception) { }
                }
                if (waitMilliseconds == 0) break;
                Thread.Sleep(100);
            }
            while (DateTime.UtcNow < deadline);
            return false;
        }

        public void Dispose()
        {
            if (mutex == null) return;
            try { if (IsOwner) mutex.ReleaseMutex(); }
            finally { IsOwner = false; mutex.Dispose(); mutex = null; }
        }
    }
}
