using System;
using System.Drawing;
using System.Windows.Forms;

namespace ESCPrintApp
{
    public partial class TableInsertDialog : Form
    {
        public int Rows { get; private set; } = 2;
        public int Columns { get; private set; } = 2;

        private NumericUpDown numRows;
        private NumericUpDown numCols;

        public TableInsertDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Insert Table";
            this.Size = new Size(300, 200);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label lblRows = new Label
            {
                Text = "Rows:",
                Location = new Point(20, 30),
                Size = new Size(50, 20)
            };

            Label lblCols = new Label
            {
                Text = "Columns:",
                Location = new Point(20, 70),
                Size = new Size(60, 20)
            };

            numRows = new NumericUpDown
            {
                Location = new Point(100, 30),
                Size = new Size(60, 20),
                Minimum = 1,
                Maximum = 10,
                Value = 2
            };

            numCols = new NumericUpDown
            {
                Location = new Point(100, 70),
                Size = new Size(60, 20),
                Minimum = 1,
                Maximum = 5,
                Value = 2
            };

            Button btnOK = new Button
            {
                Text = "OK",
                Location = new Point(70, 120),
                Size = new Size(75, 30),
                DialogResult = DialogResult.OK
            };

            Button btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(160, 120),
                Size = new Size(75, 30),
                DialogResult = DialogResult.Cancel
            };

            btnOK.Click += (s, e) =>
            {
                Rows = (int)numRows.Value;
                Columns = (int)numCols.Value;
            };

            this.Controls.AddRange(new Control[]
            {
                lblRows, lblCols, numRows, numCols, btnOK, btnCancel
            });
        }
    }
}
