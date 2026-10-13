using System;
using System.IO;
using System.Collections.Generic;
using System.Security.Cryptography;
using PcInventory;

internal static class UpgradeCheck
{
    static string Hash(string path) { using(var s=File.OpenRead(path)) using(var h=SHA256.Create()) return Convert.ToBase64String(h.ComputeHash(s)); }
    static int Main(string[] args)
    {
        try
        {
            string path=Path.GetFullPath(args[0]); var store=new InventoryStore(path);
            var hashes=new Dictionary<string,string>();
            foreach(string file in Directory.GetFiles(Path.Combine(path,"数据","记录"),"*.json")) hashes[file]=Hash(file);
            int before=store.LoadRecords().Count;
            string backup=store.BackupNow(store.BackupDirectory);
            store.RebuildSummary();
            foreach(var pair in hashes) if(Hash(pair.Key)!=pair.Value) throw new Exception("Original JSON changed.");
            Console.WriteLine("Historical records: "+before);
            Console.WriteLine("Original JSON unchanged: "+hashes.Count);
            Console.WriteLine("Complete backup created: "+File.Exists(backup));
            Console.WriteLine("Summary regenerated from every record: "+File.Exists(store.SummaryPath));
            return 0;
        }
        catch(Exception e) { Console.WriteLine(e.Message); return 1; }
    }
}
