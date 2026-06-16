using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Lab26
{
    public partial class Form1 : Form
    {
        private Word.Application? _word;
        private Word.Document?    _doc;
        private string            _templatesFolder = "";

        public Form1()
        {
            InitializeComponent();
            Load += Form1_Load;
            FormClosed += Form1_FormClosed;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _templatesFolder = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Templates");
            Directory.CreateDirectory(_templatesFolder);

            TemplateManager.EnsureTemplates(_templatesFolder);
            RefreshTemplateList();
        }

        private void RefreshTemplateList()
        {
            lstTemplates.Items.Clear();
            foreach (var f in Directory.GetFiles(_templatesFolder, "*.dotx")
                                       .Concat(Directory.GetFiles(_templatesFolder, "*.docx")))
                lstTemplates.Items.Add(new TemplateItem(f));

            if (lstTemplates.Items.Count > 0)
                lstTemplates.SelectedIndex = 0;
        }

        private void lstTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstTemplates.SelectedItem is TemplateItem ti)
            {
                lblTemplateName.Text = ti.Name;
                picPreview.Image = TemplateManager.GetPreview(ti.Path);
                UpdateFieldPanel(ti.Path);
            }
        }

        private void UpdateFieldPanel(string templatePath)
        {
            var fields = TemplateManager.GetFields(templatePath);
            pnlFields.Controls.Clear();

            int y = 4;
            foreach (var field in fields)
            {
                var lbl = new Label
                {
                    Text = field, Location = new System.Drawing.Point(4, y + 2),
                    Size = new System.Drawing.Size(110, 20),
                    Font = new System.Drawing.Font("Segoe UI", 9f)
                };
                var txt = new TextBox
                {
                    Name = "field_" + field,
                    Location = new System.Drawing.Point(118, y),
                    Size = new System.Drawing.Size(220, 23),
                    Font = new System.Drawing.Font("Segoe UI", 9.5f)
                };
                pnlFields.Controls.Add(lbl);
                pnlFields.Controls.Add(txt);
                y += 30;
            }
            pnlFields.Height = Math.Max(y + 8, 60);
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (lstTemplates.SelectedItem is not TemplateItem ti)
            { MessageBox.Show("Оберіть шаблон.", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            try
            {
                CloseWordObjects();
                _word = new Word.Application();
                object templatePath = ti.Path;
                object missing      = System.Reflection.Missing.Value;

                _doc = _word.Documents.Add(
                    ref templatePath, ref missing, ref missing, ref missing);
                _doc.Activate();

                
                foreach (Word.Bookmark bm in _doc.Bookmarks)
                {
                    var ctrl = pnlFields.Controls[$"field_{bm.Name}"] as TextBox;
                    if (ctrl != null && !string.IsNullOrEmpty(ctrl.Text))
                    {
                        object bmName = bm.Name;
                        
                        if (_doc.Bookmarks.Exists(bm.Name))
                        {
                            var range = _doc.Bookmarks[ref bmName].Range;
                            range.Text = ctrl.Text;
                        }
                    }
                }

                _word.Visible = true;
                btnSave.Enabled   = true;
                btnSearch.Enabled = true;
            }
            catch (Exception ex)
            {
                CloseWordObjects();
                MessageBox.Show("Помилка: " + ex.Message, "Помилка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (_doc == null) return;
            using var dlg = new SaveFileDialog
            {
                Title  = "Зберегти документ",
                Filter = "Word документ (*.docx)|*.docx|PDF (*.pdf)|*.pdf",
                FileName = Path.GetFileNameWithoutExtension(
                    (lstTemplates.SelectedItem as TemplateItem)?.Name ?? "document")
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            try
            {
                object path = dlg.FileName;
                object miss = System.Reflection.Missing.Value;
                if (dlg.FilterIndex == 2)
                {
                    object fmt = Word.WdSaveFormat.wdFormatPDF;
                    _doc.SaveAs2(ref path, ref fmt, ref miss, ref miss, ref miss,
                        ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
                }
                else
                {
                    object fmt = Word.WdSaveFormat.wdFormatXMLDocument;
                    _doc.SaveAs2(ref path, ref fmt, ref miss, ref miss, ref miss,
                        ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
                }
                MessageBox.Show("Збережено: " + dlg.FileName, "Готово",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка збереження: " + ex.Message);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (_doc == null) return;
            using var frm = new SearchReplaceForm();
            if (frm.ShowDialog() != DialogResult.OK) return;

            try
            {
                var find = _doc.Content.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();

                object findText    = frm.FindText;
                object replaceWith = frm.ReplaceText;
                object replace     = Word.WdReplace.wdReplaceAll;
                object missing     = System.Reflection.Missing.Value;
                object wrapCont    = Word.WdFindWrap.wdFindContinue;
                object matchCase   = false;
                object matchWord   = false;
                object matchWild   = false;
                object matchSound  = false;
                object matchAll    = false;
                object format      = false;

                find.Execute(ref findText, ref matchCase, ref matchWord,
                    ref matchWild, ref matchSound, ref matchAll, ref replace,
                    ref wrapCont, ref format, ref replaceWith, ref replace,
                    ref missing, ref missing, ref missing, ref missing);

                MessageBox.Show("Заміну виконано.", "Пошук і заміна",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
        }

        private void CloseWordObjects()
        {
            try
            {
                object miss = System.Reflection.Missing.Value;
                object no   = false;
                _doc?.Close(ref no, ref miss, ref miss);
                _word?.Quit(ref no, ref miss, ref miss);
            }
            catch { }
            _doc  = null;
            _word = null;
        }

        private void Form1_FormClosed(object? sender, FormClosedEventArgs e)
            => CloseWordObjects();
    }

    internal class TemplateItem
    {
        public string Path { get; }
        public string Name => System.IO.Path.GetFileNameWithoutExtension(Path);
        public TemplateItem(string path) { Path = path; }
        public override string ToString() => Name;
    }
}
