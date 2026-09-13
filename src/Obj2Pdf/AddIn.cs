using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: ComVisible(false)]
namespace Obj2Pdf
{
    [ComVisible(true), Guid("E6FA6B59-37ED-4E13-94C5-7F088736308D"), ProgId("Obj2Pdf.Connect"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private object application;
        private bool busy;
        public AddIn() { Trace("Constructed"); }
        private static void Trace(string message)
        {
            try { File.AppendAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Obj2Pdf", "startup.log"), DateTime.Now.ToString("s") + " " + message + "\n"); } catch { }
        }
        public void OnConnection(object app, int mode, object addIn, ref Array custom) { Trace("Connected"); application = app; }
        public void OnDisconnection(int mode, ref Array custom) { application = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
        public string GetCustomUI(string ribbonId)
        {
            Trace("Ribbon requested");
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs>
              <tab id='Obj2PdfTab' label='obj2pdf'><group id='Obj2PdfGroup' label='对象导出'>
                <button id='ExportObjects' label='选中对象转 PDF' size='large' imageMso='FileSaveAsPdfOrXps' onAction='ExportSelected' screentip='将选中对象保存为 PDF' supertip='选择一个或多个完整对象。多选时合并为一页，自动调整 PDF 页面边界。'/>
                <button id='ExportSettings' label='导出设置' imageMso='PropertySheet' onAction='ShowSettings'/>
              </group></tab></tabs></ribbon></customUI>";
        }
        public void ExportSelected(object control)
        {
            if (busy || application == null) return;
            try
            {
                dynamic app = application;
                if ((int)app.Presentations.Count == 0) throw new InvalidOperationException("请先打开一个演示文稿。");
                dynamic window = app.ActiveWindow;
                dynamic selection = window.Selection;
                if ((int)selection.Type != 2 && (int)selection.Type != 3) throw new InvalidOperationException("请先选择形状、图片、文本框或组合。按住 Ctrl 可以多选。");
                if ((bool)selection.HasChildShapeRange) throw new InvalidOperationException("请退出组合内部编辑并选中整个组合后导出。");
                dynamic shapes = selection.ShapeRange;
                int[] ids = new int[(int)shapes.Count];
                for (int i = 0; i < ids.Length; i++) ids[i] = (int)shapes[i + 1].Id;
                int slideIndex = window.View.Slide.SlideIndex;
                object source = app.ActivePresentation;
                var settings = Settings.Load();
                using (var dialog = new SaveFileDialog { Filter = "PDF 文件 (*.pdf)|*.pdf", DefaultExt = "pdf", AddExtension = true, OverwritePrompt = true,
                    FileName = Path.GetFileNameWithoutExtension((string)((dynamic)source).Name) + "-对象.pdf", Title = "将选中对象保存为 PDF" })
                {
                    if (Directory.Exists(settings.Folder)) dialog.InitialDirectory = settings.Folder;
                    if (dialog.ShowDialog(new OfficeWindow(app)) != DialogResult.OK) return;
                    busy = true;
                    Cursor.Current = Cursors.WaitCursor;
                    ExportEngine.Export(application, source, slideIndex, ids, dialog.FileName, settings.MarginMm * 72 / 25.4, settings.KeepBackground);
                    settings.Folder = Path.GetDirectoryName(dialog.FileName); settings.Save();
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show(new OfficeWindow(app), "已导出：\n" + dialog.FileName, "obj2pdf", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Obj2Pdf");
                    Directory.CreateDirectory(dir);
                    File.WriteAllText(Path.Combine(dir, "last-error.log"), DateTime.Now + "\n" + ex);
                }
                catch { }
                MessageBox.Show("导出未完成：" + ex.Message, "obj2pdf", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { busy = false; Cursor.Current = Cursors.Default; }
        }
        public void ShowSettings(object control)
        {
            try { using (var form = new SettingsForm()) form.ShowDialog(new OfficeWindow(application)); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "obj2pdf"); }
        }
    }
    internal sealed class OfficeWindow : IWin32Window
    {
        public IntPtr Handle { get; private set; }
        public OfficeWindow(dynamic app) { try { Handle = new IntPtr((int)app.HWND); } catch { Handle = IntPtr.Zero; } }
    }
    internal sealed class Settings
    {
        public double MarginMm = 0.7;
        public bool KeepBackground;
        public string Folder = "";
        public static Settings Load()
        {
            var s = new Settings();
            using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Obj2Pdf\Settings"))
            {
                if (key == null) return s;
                s.MarginMm = Math.Min(50.8, Math.Max(0, Convert.ToInt32(key.GetValue("MarginMicrons", 700)) / 1000.0));
                s.KeepBackground = Convert.ToInt32(key.GetValue("KeepBackground", 0)) != 0;
                s.Folder = Convert.ToString(key.GetValue("Folder", ""));
            }
            return s;
        }
        public void Save()
        {
            using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Obj2Pdf\Settings"))
            {
                key.SetValue("MarginMicrons", (int)Math.Round(MarginMm * 1000)); key.SetValue("KeepBackground", KeepBackground ? 1 : 0); key.SetValue("Folder", Folder);
            }
        }
    }
    internal sealed class SettingsForm : Form
    {
        public SettingsForm()
        {
            var settings = Settings.Load();
            Text = "obj2pdf · 导出设置"; ClientSize = new Size(440, 230); Font = new Font("Microsoft YaHei UI", 10);
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false; StartPosition = FormStartPosition.CenterParent;
            Controls.Add(new Label { Text = "四周留白（毫米）", Location = new Point(20, 24), AutoSize = true });
            var margin = new NumericUpDown { Location = new Point(200, 20), Width = 110, DecimalPlaces = 1, Increment = 0.1m, Maximum = 50.8m, Value = (decimal)settings.MarginMm };
            Controls.Add(margin);
            var background = new CheckBox { Text = "保留幻灯片背景（默认白底）", Location = new Point(20, 64), AutoSize = true, Checked = settings.KeepBackground };
            Controls.Add(background);
            Controls.Add(new Label { Text = "按对象几何边界裁切；阴影、发光、文字溢出时请增加留白。\n多选对象合并为一页 PDF，保持相对位置。", Location = new Point(20, 105), Size = new Size(400, 60) });
            var save = new Button { Text = "保存", Location = new Point(230, 180), Size = new Size(85, 32) };
            save.Click += (s, e) => { settings.MarginMm = (double)margin.Value; settings.KeepBackground = background.Checked; settings.Save(); DialogResult = DialogResult.OK; };
            var cancel = new Button { Text = "取消", Location = new Point(330, 180), Size = new Size(85, 32), DialogResult = DialogResult.Cancel };
            Controls.Add(save); Controls.Add(cancel); AcceptButton = save; CancelButton = cancel;
        }
    }
}
