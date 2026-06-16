namespace Lab26
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null) components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        private void InitializeComponent()
        {
            this.pnlHeader      = new System.Windows.Forms.Panel();
            this.lblTitle       = new System.Windows.Forms.Label();

            this.splitOuter     = new System.Windows.Forms.SplitContainer();
            this.splitLeft      = new System.Windows.Forms.SplitContainer();

            this.grpTemplates   = new System.Windows.Forms.GroupBox();
            this.lstTemplates   = new System.Windows.Forms.ListBox();

            this.grpPreview     = new System.Windows.Forms.GroupBox();
            this.picPreview     = new System.Windows.Forms.PictureBox();
            this.lblTemplateName= new System.Windows.Forms.Label();

            this.grpFields      = new System.Windows.Forms.GroupBox();
            this.pnlFieldsWrap  = new System.Windows.Forms.Panel();
            this.pnlFields      = new System.Windows.Forms.Panel();

            this.pnlActions     = new System.Windows.Forms.Panel();
            this.btnGenerate    = new System.Windows.Forms.Button();
            this.btnSave        = new System.Windows.Forms.Button();
            this.btnSearch      = new System.Windows.Forms.Button();

            this.pnlStatus      = new System.Windows.Forms.Panel();
            this.lblStatus      = new System.Windows.Forms.Label();

            
            this.pnlHeader.Dock      = System.Windows.Forms.DockStyle.Top;
            this.pnlHeader.Height    = 44;
            this.pnlHeader.BackColor = System.Drawing.Color.FromArgb(44, 62, 80);
            this.pnlHeader.Padding   = new System.Windows.Forms.Padding(14, 0, 0, 0);

            this.lblTitle.Text      = "Лабораторна робота №26  —  Варіант 12  (Сертифікат про закінчення)";
            this.lblTitle.Font      = new System.Drawing.Font("Segoe UI", 10.5f);
            this.lblTitle.ForeColor = System.Drawing.Color.White;
            this.lblTitle.Dock      = System.Windows.Forms.DockStyle.Fill;
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.pnlHeader.Controls.Add(this.lblTitle);

            
            this.pnlStatus.Dock      = System.Windows.Forms.DockStyle.Bottom;
            this.pnlStatus.Height    = 24;
            this.pnlStatus.BackColor = System.Drawing.Color.FromArgb(236, 240, 241);
            this.pnlStatus.Padding   = new System.Windows.Forms.Padding(8, 3, 0, 0);
            this.lblStatus.Text      = "Оберіть шаблон і заповніть поля";
            this.lblStatus.Font      = new System.Drawing.Font("Segoe UI", 8.5f);
            this.lblStatus.ForeColor = System.Drawing.Color.DimGray;
            this.lblStatus.Dock      = System.Windows.Forms.DockStyle.Fill;
            this.pnlStatus.Controls.Add(this.lblStatus);

            
            this.splitOuter.Dock          = System.Windows.Forms.DockStyle.Fill;
            this.splitOuter.Panel1MinSize = 50;
            this.splitOuter.Panel2MinSize = 50;

            
            this.splitLeft.Dock          = System.Windows.Forms.DockStyle.Fill;
            this.splitLeft.Orientation   = System.Windows.Forms.Orientation.Horizontal;
            this.splitLeft.Panel1MinSize = 50;
            this.splitLeft.Panel2MinSize = 50;

            
            this.grpTemplates.Text    = "Доступні шаблони";
            this.grpTemplates.Font    = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold);
            this.grpTemplates.Dock    = System.Windows.Forms.DockStyle.Fill;
            this.grpTemplates.Padding = new System.Windows.Forms.Padding(6, 4, 6, 6);

            this.lstTemplates.Dock          = System.Windows.Forms.DockStyle.Fill;
            this.lstTemplates.Font          = new System.Drawing.Font("Segoe UI", 9.5f);
            this.lstTemplates.BorderStyle   = System.Windows.Forms.BorderStyle.None;
            this.lstTemplates.BackColor     = System.Drawing.Color.FromArgb(248, 249, 250);
            this.lstTemplates.SelectedIndexChanged += lstTemplates_SelectedIndexChanged;
            this.grpTemplates.Controls.Add(this.lstTemplates);

            
            this.grpPreview.Text    = "Перегляд шаблону";
            this.grpPreview.Font    = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold);
            this.grpPreview.Dock    = System.Windows.Forms.DockStyle.Fill;
            this.grpPreview.Padding = new System.Windows.Forms.Padding(6, 4, 6, 24);

            this.picPreview.Dock      = System.Windows.Forms.DockStyle.Fill;
            this.picPreview.SizeMode  = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picPreview.BackColor = System.Drawing.Color.White;
            this.picPreview.BorderStyle = System.Windows.Forms.BorderStyle.None;

            this.lblTemplateName.Dock      = System.Windows.Forms.DockStyle.Bottom;
            this.lblTemplateName.Height    = 20;
            this.lblTemplateName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblTemplateName.Font      = new System.Drawing.Font("Segoe UI", 8f, System.Drawing.FontStyle.Italic);
            this.lblTemplateName.ForeColor = System.Drawing.Color.DimGray;

            this.grpPreview.Controls.Add(this.picPreview);
            this.grpPreview.Controls.Add(this.lblTemplateName);

            this.splitLeft.Panel1.Controls.Add(this.grpTemplates);
            this.splitLeft.Panel2.Controls.Add(this.grpPreview);

            this.splitOuter.Panel1.Controls.Add(this.splitLeft);

            
            
            this.grpFields.Text    = "Поля шаблону";
            this.grpFields.Font    = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold);
            this.grpFields.Dock    = System.Windows.Forms.DockStyle.Fill;
            this.grpFields.Padding = new System.Windows.Forms.Padding(6, 4, 6, 6);

            this.pnlFieldsWrap.Dock        = System.Windows.Forms.DockStyle.Fill;
            this.pnlFieldsWrap.AutoScroll  = true;

            this.pnlFields.AutoSize = true;
            this.pnlFields.Padding  = new System.Windows.Forms.Padding(4);

            this.pnlFieldsWrap.Controls.Add(this.pnlFields);
            this.grpFields.Controls.Add(this.pnlFieldsWrap);

            
            this.pnlActions.Dock      = System.Windows.Forms.DockStyle.Bottom;
            this.pnlActions.Height    = 50;
            this.pnlActions.Padding   = new System.Windows.Forms.Padding(8, 10, 8, 8);
            this.pnlActions.BackColor = System.Drawing.Color.FromArgb(245, 245, 247);

            Btn(this.btnGenerate, "▶  Створити документ", System.Drawing.Color.FromArgb(39,174,96),  8,  10, 170);
            Btn(this.btnSave,     "💾  Зберегти",          System.Drawing.Color.FromArgb(41,128,185), 184, 10, 110);
            Btn(this.btnSearch,   "🔍  Пошук і заміна",    System.Drawing.Color.FromArgb(142,68,173), 300, 10, 140);

            this.btnSave.Enabled   = false;
            this.btnSearch.Enabled = false;

            this.btnGenerate.Click += btnGenerate_Click;
            this.btnSave.Click     += btnSave_Click;
            this.btnSearch.Click   += btnSearch_Click;

            this.pnlActions.Controls.AddRange(new System.Windows.Forms.Control[]
            { this.btnGenerate, this.btnSave, this.btnSearch });

            this.splitOuter.Panel2.Controls.Add(this.grpFields);
            this.splitOuter.Panel2.Controls.Add(this.pnlActions);

            
            this.AutoScaleDimensions = new System.Drawing.SizeF(7f, 15f);
            this.AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font;
            this.Text        = "Лабораторна робота №26";
            this.Size        = new System.Drawing.Size(960, 620);
            this.MinimumSize = new System.Drawing.Size(700, 480);
            this.Font        = new System.Drawing.Font("Segoe UI", 9f);
            this.BackColor   = System.Drawing.Color.White;

            this.Controls.Add(this.splitOuter);
            this.Controls.Add(this.pnlStatus);
            this.Controls.Add(this.pnlHeader);

            this.Load += (_, _) =>
            {
                try { splitOuter.SplitterDistance = System.Math.Max(160, (int)(splitOuter.Width * 0.28)); } catch { }
                try { splitLeft.SplitterDistance  = System.Math.Max(80,  (int)(splitLeft.Height * 0.45)); } catch { }
            };
        }

        private static void Btn(System.Windows.Forms.Button b, string text,
            System.Drawing.Color bg, int x, int y, int w)
        {
            b.Text      = text;
            b.Location  = new System.Drawing.Point(x, y);
            b.Size      = new System.Drawing.Size(w, 30);
            b.BackColor = bg;
            b.ForeColor = System.Drawing.Color.White;
            b.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.Font      = new System.Drawing.Font("Segoe UI", 9f);
            b.Cursor    = System.Windows.Forms.Cursors.Hand;
        }

        #endregion

        private System.Windows.Forms.Panel          pnlHeader, pnlStatus, pnlActions;
        private System.Windows.Forms.Label          lblTitle, lblStatus, lblTemplateName;
        private System.Windows.Forms.SplitContainer splitOuter, splitLeft;
        private System.Windows.Forms.GroupBox       grpTemplates, grpPreview, grpFields;
        private System.Windows.Forms.ListBox        lstTemplates;
        private System.Windows.Forms.PictureBox     picPreview;
        private System.Windows.Forms.Panel          pnlFieldsWrap, pnlFields;
        private System.Windows.Forms.Button         btnGenerate, btnSave, btnSearch;
    }
}
