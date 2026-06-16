using System;
using System.Windows.Forms;

namespace Lab26
{
    public class SearchReplaceForm : Form
    {
        private TextBox txtFind = new(), txtReplace = new();
        private Button  btnOk = new(), btnCancel = new();

        public string FindText    => txtFind.Text;
        public string ReplaceText => txtReplace.Text;

        public SearchReplaceForm()
        {
            Text            = "Пошук і заміна";
            Size            = new System.Drawing.Size(380, 180);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition   = FormStartPosition.CenterParent;
            Font            = new System.Drawing.Font("Segoe UI", 9.5f);

            var lblF = new Label { Text = "Знайти:",   Location = new(10, 16),  Size = new(80, 20) };
            var lblR = new Label { Text = "Замінити на:", Location = new(10, 50), Size = new(80, 20) };

            txtFind.Location    = new(96, 14); txtFind.Size    = new(260, 23);
            txtReplace.Location = new(96, 48); txtReplace.Size = new(260, 23);

            btnOk.Text          = "Замінити все";
            btnOk.DialogResult  = DialogResult.OK;
            btnOk.Location      = new(130, 90); btnOk.Size = new(110, 30);
            btnOk.BackColor     = System.Drawing.Color.SteelBlue;
            btnOk.ForeColor     = System.Drawing.Color.White;
            btnOk.FlatStyle     = FlatStyle.Flat;
            btnOk.FlatAppearance.BorderSize = 0;

            btnCancel.Text         = "Скасувати";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location     = new(248, 90); btnCancel.Size = new(100, 30);
            btnCancel.FlatStyle    = FlatStyle.Flat;

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Controls.AddRange(new Control[]
            { lblF, lblR, txtFind, txtReplace, btnOk, btnCancel });
        }
    }
}
