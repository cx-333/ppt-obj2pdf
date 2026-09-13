using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

internal static class Program
{
    private const string GuidText = "{E6FA6B59-37ED-4E13-94C5-7F088736308D}";
    private const string ProgId = "Obj2Pdf.Connect";
    private const string Version = "0.1.1";
    private static readonly string Root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Obj2Pdf");
    [STAThread]
    private static int Main(string[] args)
    {
        bool silent = Array.IndexOf(args, "/quiet") >= 0;
        bool uninstall = Array.IndexOf(args, "/uninstall") >= 0;
        Application.EnableVisualStyles();
        try
        {
            if (Process.GetProcessesByName("POWERPNT").Length != 0) throw new InvalidOperationException("请保存文件并关闭所有 PowerPoint 窗口，然后重新运行安装程序。");
            if (uninstall)
            {
                if (!silent && MessageBox.Show("卸载 obj2pdf 插件？", "obj2pdf", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return 0;
                Uninstall();
            }
            else
            {
                if (!silent && MessageBox.Show("安装 obj2pdf " + Version + "？\n\n将为当前用户安装 PowerPoint 插件，无需管理员权限。\n安装后打开 PowerPoint，在 obj2pdf 选项卡中导出选中对象。", "obj2pdf 安装", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return 0;
                Install();
            }
            if (!silent) MessageBox.Show(uninstall ? "已卸载插件。" : "安装完成。\n打开 PowerPoint → obj2pdf → 选中对象转 PDF。", "obj2pdf", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 0;
        }
        catch (Exception ex)
        {
            if (!silent) MessageBox.Show(ex.Message, "obj2pdf 安装未完成", MessageBoxButtons.OK, MessageBoxIcon.Error);
            try { File.WriteAllText(Path.Combine(Path.GetTempPath(), "obj2pdf-setup-error.log"), ex.ToString()); } catch { }
            return 1;
        }
    }
    private static RegistryView[] Views { get { return Environment.Is64BitOperatingSystem ? new[] { RegistryView.Registry32, RegistryView.Registry64 } : new[] { RegistryView.Registry32 }; } }
    private static void Install()
    {
        using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
            if (key == null || Convert.ToInt32(key.GetValue("Release", 0)) < 528040) throw new InvalidOperationException("需要 .NET Framework 4.8。请通过 Windows 更新安装后重试。");
        bool hasPowerPoint = false;
        foreach (var view in Views)
        using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
        using (var key = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE"))
            if (key != null) hasPowerPoint = true;
        if (!hasPowerPoint) throw new InvalidOperationException("未检测到 Windows 桌面版 Microsoft PowerPoint。网页版和 WPS 不支持此插件。");
        Directory.CreateDirectory(Root);
        // New isolated directory makes upgrades independent of older dependency files.
        string target = Path.Combine(Root, "app-" + Version + "-" + DateTime.UtcNow.Ticks);
        Directory.CreateDirectory(target);
        using (var payload = Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.zip"))
        using (var zip = new ZipArchive(payload, ZipArchiveMode.Read))
        {
            foreach (var entry in zip.Entries)
            {
                if (entry.Name != entry.FullName) throw new InvalidDataException("安装包目录结构无效。");
                using (var input = entry.Open())
                using (var output = File.Create(Path.Combine(target, entry.Name))) input.CopyTo(output);
            }
        }
        string dll = Path.Combine(target, "Obj2Pdf.dll");
        string assembly = AssemblyName.GetAssemblyName(dll).FullName;
        foreach (var view in Views)
        using (var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
        {
            using (var clsid = hkcu.CreateSubKey(@"Software\Classes\CLSID\" + GuidText))
            {
                clsid.SetValue("", "obj2pdf PowerPoint Add-in");
                using (var prog = clsid.CreateSubKey("ProgId")) prog.SetValue("", ProgId);
                using (var server = clsid.CreateSubKey("InprocServer32"))
                {
                    server.SetValue("", "mscoree.dll"); server.SetValue("ThreadingModel", "Both");
                    server.SetValue("Class", "Obj2Pdf.AddIn"); server.SetValue("Assembly", assembly);
                    server.SetValue("RuntimeVersion", "v4.0.30319"); server.SetValue("CodeBase", new Uri(dll).AbsoluteUri);
                }
            }
            using (var prog = hkcu.CreateSubKey(@"Software\Classes\" + ProgId + @"\CLSID")) prog.SetValue("", GuidText);
            using (var addin = hkcu.CreateSubKey(@"Software\Microsoft\Office\PowerPoint\Addins\" + ProgId))
            {
                addin.SetValue("FriendlyName", "obj2pdf"); addin.SetValue("Description", "将选中对象导出为独立 PDF");
                addin.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            }
        }
        string setup = Path.Combine(Root, "Uninstall.exe");
        if (!string.Equals(Assembly.GetExecutingAssembly().Location, setup, StringComparison.OrdinalIgnoreCase)) File.Copy(Assembly.GetExecutingAssembly().Location, setup, true);
        using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\Obj2Pdf"))
        {
            key.SetValue("DisplayName", "obj2pdf - PowerPoint 对象转 PDF"); key.SetValue("DisplayVersion", Version);
            key.SetValue("Publisher", "obj2pdf"); key.SetValue("InstallLocation", Root);
            key.SetValue("UninstallString", "\"" + setup + "\" /uninstall");
            key.SetValue("QuietUninstallString", "\"" + setup + "\" /uninstall /quiet");
            key.SetValue("NoModify", 1); key.SetValue("NoRepair", 1);
        }
        foreach (string directory in Directory.GetDirectories(Root, "app-*"))
            if (!string.Equals(directory, target, StringComparison.OrdinalIgnoreCase)) SafeRemove(directory);
    }
    private static void Uninstall()
    {
        foreach (var view in Views)
        using (var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
        {
            hkcu.DeleteSubKeyTree(@"Software\Microsoft\Office\PowerPoint\Addins\" + ProgId, false);
            hkcu.DeleteSubKeyTree(@"Software\Classes\" + ProgId, false);
            hkcu.DeleteSubKeyTree(@"Software\Classes\CLSID\" + GuidText, false);
        }
        Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\Obj2Pdf", false);
        Registry.CurrentUser.DeleteSubKeyTree(@"Software\Obj2Pdf", false);
        if (Directory.Exists(Root)) foreach (string directory in Directory.GetDirectories(Root, "app-*")) SafeRemove(directory);
        // The executing uninstaller is intentionally retained; no delayed shell deletion.
        string log = Path.Combine(Root, "last-error.log");
        if (File.Exists(log)) File.Delete(log);
    }
    private static void SafeRemove(string directory)
    {
        string full = Path.GetFullPath(directory);
        if (!full.StartsWith(Path.GetFullPath(Root) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("安装目录校验失败。");
        if ((File.GetAttributes(full) & FileAttributes.ReparsePoint) != 0) throw new InvalidOperationException("安装目录不能是符号链接。");
        Directory.Delete(full, true);
    }
}
