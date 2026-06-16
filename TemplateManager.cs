using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace Lab26
{
    internal static class TemplateManager
    {
        
        private static readonly Dictionary<string, string[]> _templateFields = new()
        {
            ["Сертифікат_про_закінчення"] = new[] { "FullName", "Course", "Institution", "Date", "Director" },
            ["Почесна_грамота"]           = new[] { "FullName", "Achievement", "Organization", "Date" },
            ["Довідка_про_навчання"]      = new[] { "FullName", "Group", "Faculty", "Year", "Rector" },
        };

        public static void EnsureTemplates(string folder)
        {
            foreach (var (name, fields) in _templateFields)
            {
                string path = Path.Combine(folder, name + ".docx");
                
                TryCreateWithWord(path, name, fields);

                string prev = Path.Combine(folder, name + ".png");
                if (!File.Exists(prev))
                    CreatePreviewImage(prev, name, fields);
            }
        }

        public static string[] GetFields(string templatePath)
        {
            string name = Path.GetFileNameWithoutExtension(templatePath);
            return _templateFields.TryGetValue(name, out var f) ? f : Array.Empty<string>();
        }

        public static Image? GetPreview(string templatePath)
        {
            string prev = Path.ChangeExtension(templatePath, ".png");
            if (File.Exists(prev))
            {
                
                byte[] bytes = File.ReadAllBytes(prev);
                using var ms = new System.IO.MemoryStream(bytes);
                return Image.FromStream(ms);
            }
            return null;
        }

        private static void TryCreateWithWord(string path, string title, string[] fields)
        {
            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
                object missing = System.Reflection.Missing.Value;

                var doc = wordApp.Documents.Add(
                    ref missing, ref missing, ref missing, ref missing);

                
                string displayTitle = title.Replace("_", " ");
                var titlePara = doc.Paragraphs.Add(ref missing);
                titlePara.Range.Text = displayTitle;
                titlePara.Range.Font.Size = 22;
                titlePara.Range.Font.Bold = 1;
                titlePara.Range.Font.Color =
                    Microsoft.Office.Interop.Word.WdColor.wdColorDarkBlue;
                titlePara.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                titlePara.Range.InsertParagraphAfter();

                
                var linePara = doc.Paragraphs.Add(ref missing);
                linePara.Range.Text = new string('─', 55);
                linePara.Range.Font.Size = 10;
                linePara.Range.Font.Color =
                    Microsoft.Office.Interop.Word.WdColor.wdColorGold;
                linePara.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                linePara.Range.InsertParagraphAfter();

                
                doc.Paragraphs.Add(ref missing).Range.InsertParagraphAfter();

                
                foreach (var field in fields)
                {
                    var para = doc.Paragraphs.Add(ref missing);
                    var r = para.Range;

                    
                    r.Text = FriendlyLabel(field) + ":  ";
                    r.Font.Bold = 1;
                    r.Font.Size = 12;
                    r.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

                    
                    r.Collapse(
                        Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

                    
                    r.Text = $"[{FriendlyLabel(field)}]";
                    r.Font.Bold  = 0;
                    r.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorGray50;

                    
                    
                    doc.Bookmarks.Add(field, ref (object)r);

                    para.Range.InsertParagraphAfter();
                    doc.Paragraphs.Add(ref missing).Range.InsertParagraphAfter();
                }

                
                for (int i = 0; i < 2; i++)
                    doc.Paragraphs.Add(ref missing).Range.InsertParagraphAfter();

                
                var signPara = doc.Paragraphs.Add(ref missing);
                signPara.Range.Text = "М.П.                    ___________________";
                signPara.Range.Font.Size = 11;
                signPara.Range.Font.Bold = 0;
                signPara.Range.InsertParagraphAfter();

                
                object savePath = path;
                object fmt  = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument;
                object fFalse = false;
                doc.SaveAs2(ref savePath, ref fmt, ref fFalse,
                    ref missing, ref fFalse, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                doc.Close(ref fFalse, ref missing, ref missing);
                wordApp.Quit(ref fFalse, ref missing, ref missing);
            }
            catch (Exception ex)
            {
                
                System.Diagnostics.Debug.WriteLine("TemplateManager: " + ex.Message);
            }
        }

        private static void CreatePreviewImage(string path, string title, string[] fields)
        {
            int W = 280, H = 380;
            using var bmp = new Bitmap(W, H);
            using var g   = Graphics.FromImage(bmp);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            using var bg = new SolidBrush(Color.FromArgb(252, 248, 240));
            g.FillRectangle(bg, 0, 0, W, H);

            using var pen1 = new Pen(Color.FromArgb(180, 140, 60), 3f);
            g.DrawRectangle(pen1, 6, 6, W - 13, H - 13);
            using var pen2 = new Pen(Color.FromArgb(200, 170, 90), 1f);
            g.DrawRectangle(pen2, 12, 12, W - 25, H - 25);

            string display = title.Replace("_", " ");
            using var tf = new Font("Segoe UI", 11f, FontStyle.Bold);
            using var tb = new SolidBrush(Color.FromArgb(30, 30, 100));
            var sf = new StringFormat { Alignment = StringAlignment.Center };
            g.DrawString(display, tf, tb, new RectangleF(20, 22, W - 40, 68), sf);

            using var lp = new Pen(Color.FromArgb(180, 140, 60), 1f);
            g.DrawLine(lp, 20, 88, W - 20, 88);

            using var ff  = new Font("Segoe UI", 8.5f);
            using var lb  = new SolidBrush(Color.FromArgb(80, 60, 30));
            using var vb  = new SolidBrush(Color.FromArgb(140, 140, 140));
            using var lp2 = new Pen(Color.FromArgb(180, 160, 120), 0.8f);

            int y = 102;
            foreach (var field in fields)
            {
                g.DrawString(FriendlyLabel(field) + ":", ff, lb, 22, y);
                g.DrawString($"[{FriendlyLabel(field)}]", ff, vb, 22, y + 14);
                g.DrawLine(lp2, 22, y + 30, W - 22, y + 30);
                y += 44;
                if (y > H - 55) break;
            }

            using var sb2 = new SolidBrush(Color.FromArgb(50, 0, 80, 180));
            g.FillEllipse(sb2, W - 70, H - 58, 44, 44);
            using var wf  = new Font("Segoe UI", 6f);
            using var wb  = new SolidBrush(Color.White);
            g.DrawString("М.П.", wf, wb,
                new RectangleF(W - 70, H - 58, 44, 44),
                new StringFormat { Alignment = StringAlignment.Center,
                                   LineAlignment = StringAlignment.Center });

            bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        }

        private static string FriendlyLabel(string field) => field switch
        {
            "FullName"     => "ПІБ",
            "Course"       => "Курс / Програма",
            "Institution"  => "Заклад освіти",
            "Date"         => "Дата видачі",
            "Director"     => "Директор / Ректор",
            "Achievement"  => "Досягнення",
            "Organization" => "Організація",
            "Group"        => "Група",
            "Faculty"      => "Факультет",
            "Year"         => "Навчальний рік",
            "Rector"       => "Ректор",
            _ => field
        };
    }
}
