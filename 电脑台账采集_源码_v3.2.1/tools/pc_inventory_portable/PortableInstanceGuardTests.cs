using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using PcInventory.Ios;

public static class PortableInstanceGuardTests
{
    private static int checks;
    private static readonly List<Process> Children = new List<Process>();
    private static void Check(bool value, string name) { if (!value) throw new Exception("FAILED: " + name); checks++; Console.WriteLine("PASS " + name); }
    private static string Quote(string value)
    {
        var argument = new StringBuilder("\""); int slashes = 0;
        foreach (char ch in value)
        {
            if (ch == '\\') { slashes++; continue; }
            argument.Append('\\', ch == '"' ? slashes * 2 + 1 : slashes); argument.Append(ch); slashes = 0;
        }
        return argument.Append('\\', slashes * 2).Append('"').ToString();
    }
    private static Process Start(string mode, string directory, string result, string release)
    {
        string executable = System.Reflection.Assembly.GetExecutingAssembly().Location;
        return StartAt(executable, mode, directory, result, release);
    }
    private static Process StartAt(string executable, string mode, string directory, string result, string release)
    {
        var process = Process.Start(new ProcessStartInfo(executable, mode + " " + Quote(directory) + " " + Quote(result) + " " + Quote(release)) { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden });
        Children.Add(process); return process;
    }
    private static string ReadResult(string path)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(10);
        do { try { if (File.Exists(path)) { string text = File.ReadAllText(path); if (text.Length != 0) return text; } } catch (IOException) { } Thread.Sleep(25); } while (DateTime.UtcNow < deadline);
        throw new Exception("Child did not report in time.");
    }
    private static void Stop(Process child, string release)
    { File.WriteAllText(release, "release"); if (!child.WaitForExit(5000)) throw new Exception("Child failed to release lock."); }
    private static string Probe(string directory, string root, string name)
    {
        string result = Path.Combine(root, name + ".result"), release = Path.Combine(root, name + ".release");
        Process child = Start("probe", directory, result, release);
        string text = ReadResult(result); if (!child.WaitForExit(5000)) throw new Exception("Probe did not exit."); return text;
    }
    public static int Main(string[] args)
    {
        if (args.Length == 4 && (args[0] == "hold" || args[0] == "probe" || args[0] == "legacywindow" || args[0] == "markedwindow"))
        {
            using (var instance = args[0] == "legacywindow" ? null : PortableInstanceGuard.Acquire(args[1]))
            {
                bool owner = instance == null || instance.IsOwner;
                System.Windows.Forms.Form window = null;
                if (owner && (args[0] == "legacywindow" || args[0] == "markedwindow"))
                {
                    window = new System.Windows.Forms.Form { Text = "Synthetic single-instance test", ShowInTaskbar = true, StartPosition = System.Windows.Forms.FormStartPosition.Manual, Location = new System.Drawing.Point(-20000, -20000), Size = new System.Drawing.Size(220, 90) };
                    if (args[0] == "markedwindow") PortableInstanceGuard.MarkWindow(window.Handle);
                    window.Show(); System.Windows.Forms.Application.DoEvents();
                }
                File.WriteAllText(args[2], owner ? "owner" : "duplicate", new UTF8Encoding(false));
                if (owner && args[0] != "probe")
                { DateTime deadline = DateTime.UtcNow.AddSeconds(30); while (!File.Exists(args[3]) && DateTime.UtcNow < deadline) { if (window != null) System.Windows.Forms.Application.DoEvents(); Thread.Sleep(30); } }
                if (window != null) { PortableInstanceGuard.UnmarkWindow(window.Handle); window.Close(); window.Dispose(); }
                return 0;
            }
        }
        try
        {
            string root = Path.GetFullPath(args[0]); Directory.CreateDirectory(root);
            string a = Path.Combine(root, "folder A"), b = Path.Combine(root, "folder B"); Directory.CreateDirectory(a); Directory.CreateDirectory(b);
            string readyA = Path.Combine(root, "a.ready"), releaseA = Path.Combine(root, "a.release");
            Process first = Start("hold", a, readyA, releaseA);
            Check(ReadResult(readyA) == "owner", "first app owns its folder lock");
            Check(Probe(a, root, "same") == "duplicate", "a second process in the same folder cannot acquire lock");
            Check(Probe(a.ToUpperInvariant() + Path.DirectorySeparatorChar, root, "alias") == "duplicate", "case and trailing slash aliases share the lock");
            Check(Probe(b, root, "different") == "owner", "another portable folder can run independently");
            Stop(first, releaseA);
            Check(Probe(a, root, "after-release") == "owner", "normal shutdown allows opening the folder again");

            string crashReady = Path.Combine(root, "crash.ready"), crashRelease = Path.Combine(root, "crash.release");
            Process crashed = Start("hold", a, crashReady, crashRelease);
            Check(ReadResult(crashReady) == "owner", "synthetic crash process owns lock before termination");
            using (var observer = PortableInstanceGuard.Acquire(a))
            {
                Check(!observer.IsOwner, "observer keeps existing mutex handle without owning it");
                // Terminate only this test's own synthetic child, never a user application.
                crashed.Kill(); crashed.WaitForExit();
                using (var recovered = PortableInstanceGuard.Acquire(a))
                    Check(recovered.IsOwner && recovered.RecoveredAbandonedLock, "unexpected process exit is recovered without stale lock");
            }
            Check(Probe(a, root, "after-crash") == "owner", "recovery releases cleanly for next normal start");

            string ready1 = Path.Combine(root, "race1.ready"), ready2 = Path.Combine(root, "race2.ready");
            string release1 = Path.Combine(root, "race1.release"), release2 = Path.Combine(root, "race2.release");
            Process race1 = Start("hold", a, ready1, release1), race2 = Start("hold", a, ready2, release2);
            string outcome1 = ReadResult(ready1), outcome2 = ReadResult(ready2);
            Check((outcome1 == "owner") != (outcome2 == "owner"), "simultaneous double launch has exactly one owner");
            File.WriteAllText(release1, "release"); File.WriteAllText(release2, "release");
            Check(race1.WaitForExit(5000) && race2.WaitForExit(5000), "both synthetic launch processes exit cleanly");
            Check(Directory.GetFiles(a).Length == 0 && Directory.GetFiles(b).Length == 0, "single-instance mechanism writes no data or lock files");
            string legacyExecutable = Path.Combine(a, "电脑台账采集_合成旧版.exe");
            File.Copy(System.Reflection.Assembly.GetExecutingAssembly().Location, legacyExecutable);
            string legacyReady = Path.Combine(root, "legacy.ready"), legacyRelease = Path.Combine(root, "legacy.release");
            Process old = StartAt(legacyExecutable, "legacywindow", a, legacyReady, legacyRelease);
            Check(ReadResult(legacyReady) == "owner", "synthetic older-named collector process starts in its own folder");
            Check(PortableInstanceGuard.HasLegacyInstance(a), "legacy collector is identified by exact folder and live unmarked window");
            Check(!PortableInstanceGuard.HasLegacyInstance(b), "legacy collector in another folder is never selected");
            if (args.Length > 1)
            {
                File.Move(legacyExecutable, Path.Combine(a, "电脑台账采集_合成旧进程.exe"));
                File.Copy(Path.GetFullPath(args[1]), legacyExecutable);
                Check(FileVersionInfo.GetVersionInfo(legacyExecutable).FileMajorPart == 3 && FileVersionInfo.GetVersionInfo(legacyExecutable).FileMinorPart >= 1, "old executable path now contains new-release metadata");
                Check(PortableInstanceGuard.HasLegacyInstance(a), "upgrading disk executable cannot disguise a live old window");
            }
            Stop(old, legacyRelease);
            Check(!PortableInstanceGuard.HasLegacyInstance(a), "closed legacy process does not block a later start");
            string markedExecutable = Path.Combine(b, "电脑台账采集_活体标记测试.exe");
            File.Copy(System.Reflection.Assembly.GetExecutingAssembly().Location, markedExecutable);
            string markedReady = Path.Combine(root, "marked.ready"), markedRelease = Path.Combine(root, "marked.release");
            Process marked = StartAt(markedExecutable, "markedwindow", b, markedReady, markedRelease);
            Check(ReadResult(markedReady) == "owner", "synthetic mutex-aware window starts and is marked");
            Check(!PortableInstanceGuard.HasLegacyInstance(b), "live window marker distinguishes current instance from legacy disk metadata");
            Stop(marked, markedRelease);
            Console.WriteLine("ALL " + checks + " SINGLE INSTANCE CHECKS PASSED"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            foreach (Process child in Children)
            {
                try { if (!child.HasExited) { child.Kill(); child.WaitForExit(); } } catch (InvalidOperationException) { }
                child.Dispose();
            }
        }
    }
}
