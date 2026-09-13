using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.IO.Compression;
using System.Xml.Linq;

namespace Obj2Pdf
{
    public struct Bounds
    {
        public double Left, Top, Right, Bottom;
        public double Width { get { return Right - Left; } }
        public double Height { get { return Bottom - Top; } }
        public static Bounds Rotated(double x, double y, double width, double height, double angle, double stroke)
        {
            double r = angle * Math.PI / 180;
            double w = Math.Abs(width * Math.Cos(r)) + Math.Abs(height * Math.Sin(r));
            double h = Math.Abs(width * Math.Sin(r)) + Math.Abs(height * Math.Cos(r));
            return new Bounds { Left = x + width / 2 - w / 2 - stroke / 2, Top = y + height / 2 - h / 2 - stroke / 2,
                Right = x + width / 2 + w / 2 + stroke / 2, Bottom = y + height / 2 + h / 2 + stroke / 2 };
        }
        public void Include(Bounds b)
        {
            Left = Math.Min(Left, b.Left); Top = Math.Min(Top, b.Top);
            Right = Math.Max(Right, b.Right); Bottom = Math.Max(Bottom, b.Bottom);
        }
    }

    public static class ExportEngine
    {
        internal static void SetShapePosition(object shape, string property, double value, int slideIndex)
        {
            dynamic s = shape;
            string detail = string.Format(CultureInfo.InvariantCulture,
                "幻灯片 {0}，对象“{1}”(Id={2}, Type={3})，{4}={5:R} pt",
                slideIndex, (string)s.Name, (int)s.Id, (int)s.Type, property, value);
            if (double.IsNaN(value) || double.IsInfinity(value) || Math.Abs(value) > float.MaxValue)
                throw new InvalidOperationException("对象位置不是有效数值：" + detail);
            try
            {
                if (property == "Left") s.Left = (float)value;
                else if (property == "Top") s.Top = (float)value;
                else throw new ArgumentException("Unknown position property.", "property");
            }
            catch (Exception ex) when (ex is ArgumentException || ex is COMException)
            {
                throw new InvalidOperationException("PowerPoint 拒绝设置对象位置：" + detail + "。", ex);
            }
        }
        // Works with the user's live presentation, including unsaved changes, without editing it.
        public static void Export(object application, object presentation, int slideIndex, int[] shapeIds, string output, double marginPoints, bool keepBackground)
        {
            if (shapeIds == null || shapeIds.Length == 0) throw new ArgumentException("请先选择要导出的对象。");
            if (double.IsNaN(marginPoints) || marginPoints < 0 || marginPoints > 144) throw new ArgumentException("留白必须在 0–50.8 毫米之间。");
            string destination = Path.GetFullPath(output);
            if (!string.Equals(Path.GetExtension(destination), ".pdf", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("输出文件必须是 PDF。");
            string work = Path.Combine(Path.GetTempPath(), "Obj2Pdf-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(work);
            string staged = Path.Combine(Path.GetDirectoryName(destination), ".obj2pdf-" + Guid.NewGuid().ToString("N") + ".pdf");
            dynamic app = application, source = presentation, copy = null;
            try
            {
                string snapshot = Path.Combine(work, "snapshot.pptx");
                // Open XML presentation, no macro execution and no clipboard dependency.
                source.SaveCopyAs(snapshot, 24, 0);
                var ids = new HashSet<int>(shapeIds);
                Bounds bounds = new Bounds { Left = double.PositiveInfinity, Top = double.PositiveInfinity, Right = double.NegativeInfinity, Bottom = double.NegativeInfinity };
                dynamic originalSlide = source.Slides[slideIndex];
                int sourceFound = 0;
                for (int i = 1; i <= (int)originalSlide.Shapes.Count; i++)
                {
                    dynamic s = originalSlide.Shapes[i];
                    if (!ids.Contains((int)s.Id)) continue;
                    sourceFound++;
                    double stroke = 0;
                    try { if ((int)s.Line.Visible != 0) stroke = Math.Max(0, (double)s.Line.Weight); } catch (COMException) { }
                    Bounds current = Bounds.Rotated((double)s.Left, (double)s.Top, (double)s.Width, (double)s.Height, (double)s.Rotation, stroke);
                    if (double.IsNaN(current.Width) || double.IsInfinity(current.Width) || double.IsNaN(current.Height) || double.IsInfinity(current.Height))
                        throw new InvalidOperationException("对象边界无效：" + (string)s.Name + " (Id=" + (int)s.Id + ")。");
                    bounds.Include(current);
                }
                if (sourceFound != ids.Count) throw new InvalidOperationException("无法匹配全部对象。请选择完整对象或整个组合后重试。");
                double width = Math.Max(0.1, bounds.Width + marginPoints * 2), height = Math.Max(0.1, bounds.Height + marginPoints * 2);
                if (width > 4032 || height > 4032) throw new InvalidOperationException("所选对象超过 PowerPoint 的 56 英寸页面上限，请缩小对象后重试。");
                // Changing PageSetup through COM rescales text, even after restoring shape dimensions.
                // Edit only the temporary Open XML page definition before PowerPoint opens it.
                using (var zip = ZipFile.Open(snapshot, ZipArchiveMode.Update))
                {
                    var entry = zip.GetEntry("ppt/presentation.xml");
                    XDocument xml;
                    using (var stream = entry.Open()) xml = XDocument.Load(stream);
                    XNamespace ns = "http://schemas.openxmlformats.org/presentationml/2006/main";
                    var size = xml.Root.Element(ns + "sldSz");
                    size.SetAttributeValue("cx", (long)Math.Ceiling(Math.Max(72, width) * 12700));
                    size.SetAttributeValue("cy", (long)Math.Ceiling(Math.Max(72, height) * 12700));
                    size.SetAttributeValue("type", "custom");
                    entry.Delete();
                    using (var stream = zip.CreateEntry("ppt/presentation.xml").Open()) xml.Save(stream);
                }
                copy = app.Presentations.Open(snapshot, 0, 0, 0);
                for (int i = (int)copy.Slides.Count; i >= 1; i--) if (i != slideIndex) copy.Slides[i].Delete();
                dynamic slide = copy.Slides[1];
                int found = 0;
                for (int i = (int)slide.Shapes.Count; i >= 1; i--)
                {
                    dynamic shape = slide.Shapes[i];
                    if (!ids.Contains((int)shape.Id)) shape.Delete(); else found++;
                }
                if (found != ids.Count) throw new InvalidOperationException("无法匹配全部对象。请退出组合内部编辑，选择完整对象或整个组合后重试。");
                slide.DisplayMasterShapes = 0;
                slide.SlideShowTransition.Hidden = 0;
                if (!keepBackground)
                {
                    slide.FollowMasterBackground = 0;
                    slide.Background.Fill.Solid();
                    slide.Background.Fill.ForeColor.RGB = 0xFFFFFF;
                    try { slide.Background.Fill.Transparency = 0f; }
                    catch (Exception ex) when (ex is ArgumentException || ex is COMException)
                    { throw new InvalidOperationException("PowerPoint 拒绝设置背景透明度 Transparency=0。", ex); }
                }
                for (int i = 1; i <= (int)slide.Shapes.Count; i++)
                {
                    dynamic s = slide.Shapes[i];
                    SetShapePosition((object)s, "Left", (double)s.Left - bounds.Left + marginPoints, slideIndex);
                    SetShapePosition((object)s, "Top", (double)s.Top - bounds.Top + marginPoints, slideIndex);
                }
                string raw = Path.Combine(work, "native.pdf");
                // ppSaveAsPDF uses PowerPoint's native fixed-format exporter.
                copy.SaveAs(raw, 32, 0);
                string crop = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Obj2Pdf.Crop.exe");
                using (var process = Process.Start(new ProcessStartInfo(crop,
                    "\"" + raw + "\" \"" + staged + "\" " + width.ToString("R", CultureInfo.InvariantCulture) + " " + height.ToString("R", CultureInfo.InvariantCulture))
                    { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden }))
                {
                    if (!process.WaitForExit(60000)) { process.Kill(); throw new TimeoutException("PDF 裁切超时。"); }
                    if (process.ExitCode != 0) throw new InvalidOperationException("PDF 裁切失败。" + (File.Exists(staged + ".error") ? File.ReadAllText(staged + ".error") : ""));
                }
                // A failed export never truncates an existing destination.
                if (File.Exists(destination)) File.Replace(staged, destination, null); else File.Move(staged, destination);
            }
            finally
            {
                if (copy != null)
                {
                    try { copy.Saved = -1; copy.Close(); } catch { }
                    Marshal.ReleaseComObject((object)copy);
                }
                try { if (File.Exists(staged)) File.Delete(staged); } catch { }
                try { if (File.Exists(staged + ".error")) File.Delete(staged + ".error"); } catch { }
                try { Directory.Delete(work, true); } catch { }
            }
        }
    }
}
