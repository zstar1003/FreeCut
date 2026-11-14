using System;
using System.IO;

namespace FreeCutInstaller
{
    // Files must come from one build; never fill a missing manifest from an old install.
    internal static class InstallPackage
    {
        internal static readonly string[] RequiredFiles = {
            "PowerPointAddIn1.dll",
            "PowerPointAddIn1.dll.manifest",
            "PowerPointAddIn1.vsto"
        };

        internal static string FindSourceDirectory(string installerDirectory)
        {
            string[] relativePaths = {
                ".",
                @"FreeCutPPT\PowerPointAddIn1\bin\Release",
                @"FreeCutPPT\PowerPointAddIn1\bin\Debug",
                @"..\FreeCutPPT\PowerPointAddIn1\bin\Release",
                @"..\FreeCutPPT\PowerPointAddIn1\bin\Debug",
                @"bin\Release",
                @"bin\Debug"
            };
            foreach (string relativePath in relativePaths)
            {
                string directory = Path.GetFullPath(Path.Combine(installerDirectory, relativePath));
                if (File.Exists(Path.Combine(directory, RequiredFiles[0])))
                    return directory;
            }
            return null;
        }

        internal static void Validate(string directory)
        {
            foreach (string name in RequiredFiles)
            {
                string path = Path.Combine(directory, name);
                if (!File.Exists(path) || new FileInfo(path).Length == 0)
                    throw new InvalidDataException("安装包缺少文件或文件为空: " + path +
                        "。请使用同一次构建生成的完整安装包，勿单独复制 DLL。");
            }
        }

        internal static void CopyFile(string source, string destination)
        {
            // Allow re-running an installer located inside C:\FreeCut.
            if (!string.Equals(Path.GetFullPath(source), Path.GetFullPath(destination),
                StringComparison.OrdinalIgnoreCase))
                File.Copy(source, destination, true);
        }
    }
}
