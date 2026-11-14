using System;
using System.IO;
using FreeCutInstaller;

class InstallPackageTests
{
    static void Assert(bool result, string message)
    {
        if (!result) throw new Exception(message);
    }

    static void Rejected(string directory)
    {
        try { InstallPackage.Validate(directory); }
        catch (InvalidDataException) { return; }
        throw new Exception("Incomplete package was accepted");
    }

    static int Main()
    {
        string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(root);
        string cwd = Environment.CurrentDirectory;
        try
        {
            Assert(InstallPackage.FindSourceDirectory(root) == null, "Empty directory found");
            foreach (string name in InstallPackage.RequiredFiles)
                File.WriteAllText(Path.Combine(root, name), "test");
            InstallPackage.Validate(root);
            foreach (string name in InstallPackage.RequiredFiles)
            {
                string path = Path.Combine(root, name);
                File.Delete(path);
                Rejected(root);
                File.WriteAllText(path, "");
                Rejected(root);
                File.WriteAllText(path, "test");
            }
            string nested = Path.Combine(root, @"FreeCutPPT\PowerPointAddIn1\bin\Release");
            Directory.CreateDirectory(nested);
            foreach (string name in InstallPackage.RequiredFiles)
                File.Move(Path.Combine(root, name), Path.Combine(nested, name));
            Environment.CurrentDirectory = Path.GetTempPath();
            Assert(InstallPackage.FindSourceDirectory(root) == nested, "Search depends on working directory");
            InstallPackage.Validate(nested);
            string dll = Path.Combine(nested, InstallPackage.RequiredFiles[0]);
            InstallPackage.CopyFile(dll, dll);
            string destination = Path.Combine(root, "copy.dll");
            InstallPackage.CopyFile(dll, destination);
            Assert(File.ReadAllText(destination) == "test", "Copy failed");
            File.WriteAllText(destination, "old");
            InstallPackage.CopyFile(dll, destination);
            Assert(File.ReadAllText(destination) == "test", "Overwrite failed");
            Console.WriteLine("All installer package regression tests passed.");
            return 0;
        }
        finally
        {
            Environment.CurrentDirectory = cwd;
            Directory.Delete(root, true);
        }
    }
}
