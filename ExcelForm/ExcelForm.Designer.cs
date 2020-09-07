namespace ExcelForm
{
    partial class ExcelForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.cboPaiXu = new System.Windows.Forms.ComboBox();
            this.button_selt = new System.Windows.Forms.Button();
            this.button_define = new System.Windows.Forms.Button();
            this.checkBox_dhcl = new System.Windows.Forms.CheckBox();
            this.label_colWidth = new System.Windows.Forms.Label();
            this.label_rowHeight = new System.Windows.Forms.Label();
            this.textBox_colWidth = new System.Windows.Forms.TextBox();
            this.textBox_rowHeight = new System.Windows.Forms.TextBox();
            this.radio_col = new System.Windows.Forms.RadioButton();
            this.radio_row = new System.Windows.Forms.RadioButton();
            this.label_col = new System.Windows.Forms.Label();
            this.textBox_col = new System.Windows.Forms.TextBox();
            this.label_row = new System.Windows.Forms.Label();
            this.textBox_row = new System.Windows.Forms.TextBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.textBox_HeaderNameEdit = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(11, 50);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(908, 400);
            this.dataGridView1.TabIndex = 26;
            this.dataGridView1.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridView1_CellBeginEdit);
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEndEdit);
            this.dataGridView1.ColumnHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_ColumnHeaderMouseDoubleClick);
            this.dataGridView1.CurrentCellChanged += new System.EventHandler(this.dataGridView1_CurrentCellChanged);
            this.dataGridView1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dataGridView1_KeyPress);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.cboPaiXu);
            this.panel1.Controls.Add(this.button_selt);
            this.panel1.Controls.Add(this.button_define);
            this.panel1.Controls.Add(this.checkBox_dhcl);
            this.panel1.Controls.Add(this.label_colWidth);
            this.panel1.Controls.Add(this.label_rowHeight);
            this.panel1.Controls.Add(this.textBox_colWidth);
            this.panel1.Controls.Add(this.textBox_rowHeight);
            this.panel1.Controls.Add(this.radio_col);
            this.panel1.Controls.Add(this.radio_row);
            this.panel1.Controls.Add(this.label_col);
            this.panel1.Controls.Add(this.textBox_col);
            this.panel1.Controls.Add(this.label_row);
            this.panel1.Controls.Add(this.textBox_row);
            this.panel1.Location = new System.Drawing.Point(11, 8);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(908, 35);
            this.panel1.TabIndex = 18;
            // 
            // cboPaiXu
            // 
            this.cboPaiXu.FormattingEnabled = true;
            this.cboPaiXu.Items.AddRange(new object[] {
            "默认排序",
            "阅读顺序",
            "行列顺序"});
            this.cboPaiXu.Location = new System.Drawing.Point(615, 6);
            this.cboPaiXu.Name = "cboPaiXu";
            this.cboPaiXu.Size = new System.Drawing.Size(104, 23);
            this.cboPaiXu.TabIndex = 27;
            // 
            // button_selt
            // 
            this.button_selt.Location = new System.Drawing.Point(814, 3);
            this.button_selt.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button_selt.Name = "button_selt";
            this.button_selt.Size = new System.Drawing.Size(91, 28);
            this.button_selt.TabIndex = 25;
            this.button_selt.Text = "选择文字";
            this.button_selt.UseVisualStyleBackColor = true;
            this.button_selt.Click += new System.EventHandler(this.button_selt_Click);
            // 
            // button_define
            // 
            this.button_define.Location = new System.Drawing.Point(726, 2);
            this.button_define.Margin = new System.Windows.Forms.Padding(4);
            this.button_define.Name = "button_define";
            this.button_define.Size = new System.Drawing.Size(81, 28);
            this.button_define.TabIndex = 24;
            this.button_define.Text = "绘制表格";
            this.button_define.UseVisualStyleBackColor = true;
            this.button_define.Click += new System.EventHandler(this.button_define_Click_1);
            // 
            // checkBox_dhcl
            // 
            this.checkBox_dhcl.AutoSize = true;
            this.checkBox_dhcl.Location = new System.Drawing.Point(359, 9);
            this.checkBox_dhcl.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.checkBox_dhcl.Name = "checkBox_dhcl";
            this.checkBox_dhcl.Size = new System.Drawing.Size(104, 19);
            this.checkBox_dhcl.TabIndex = 21;
            this.checkBox_dhcl.Text = "按单行处理";
            this.toolTip1.SetToolTip(this.checkBox_dhcl, "多行文字按照单行文字处理");
            this.checkBox_dhcl.UseVisualStyleBackColor = true;
            // 
            // label_colWidth
            // 
            this.label_colWidth.AutoSize = true;
            this.label_colWidth.Location = new System.Drawing.Point(263, 10);
            this.label_colWidth.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_colWidth.Name = "label_colWidth";
            this.label_colWidth.Size = new System.Drawing.Size(37, 15);
            this.label_colWidth.TabIndex = 26;
            this.label_colWidth.Text = "列宽";
            // 
            // label_rowHeight
            // 
            this.label_rowHeight.AutoSize = true;
            this.label_rowHeight.Location = new System.Drawing.Point(165, 10);
            this.label_rowHeight.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_rowHeight.Name = "label_rowHeight";
            this.label_rowHeight.Size = new System.Drawing.Size(37, 15);
            this.label_rowHeight.TabIndex = 24;
            this.label_rowHeight.Text = "行高";
            // 
            // textBox_colWidth
            // 
            this.textBox_colWidth.Location = new System.Drawing.Point(308, 4);
            this.textBox_colWidth.Margin = new System.Windows.Forms.Padding(4);
            this.textBox_colWidth.MaxLength = 4;
            this.textBox_colWidth.Name = "textBox_colWidth";
            this.textBox_colWidth.Size = new System.Drawing.Size(44, 25);
            this.textBox_colWidth.TabIndex = 20;
            this.textBox_colWidth.Text = "8.0";
            this.textBox_colWidth.TextChanged += new System.EventHandler(this.textBox_colWidth_TextChanged);
            this.textBox_colWidth.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_colWidth_KeyPress);
            // 
            // textBox_rowHeight
            // 
            this.textBox_rowHeight.Location = new System.Drawing.Point(212, 4);
            this.textBox_rowHeight.Margin = new System.Windows.Forms.Padding(4);
            this.textBox_rowHeight.MaxLength = 4;
            this.textBox_rowHeight.Name = "textBox_rowHeight";
            this.textBox_rowHeight.Size = new System.Drawing.Size(44, 25);
            this.textBox_rowHeight.TabIndex = 19;
            this.textBox_rowHeight.Text = "1.5";
            this.textBox_rowHeight.TextChanged += new System.EventHandler(this.textBox_rowHeight_TextChanged);
            this.textBox_rowHeight.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_rowHeight_KeyPress);
            // 
            // radio_col
            // 
            this.radio_col.AutoSize = true;
            this.radio_col.Location = new System.Drawing.Point(550, 8);
            this.radio_col.Margin = new System.Windows.Forms.Padding(4);
            this.radio_col.Name = "radio_col";
            this.radio_col.Size = new System.Drawing.Size(58, 19);
            this.radio_col.TabIndex = 23;
            this.radio_col.Text = "按列";
            this.radio_col.UseVisualStyleBackColor = true;
            // 
            // radio_row
            // 
            this.radio_row.AutoSize = true;
            this.radio_row.Checked = true;
            this.radio_row.Location = new System.Drawing.Point(475, 8);
            this.radio_row.Margin = new System.Windows.Forms.Padding(4);
            this.radio_row.Name = "radio_row";
            this.radio_row.Size = new System.Drawing.Size(58, 19);
            this.radio_row.TabIndex = 22;
            this.radio_row.TabStop = true;
            this.radio_row.Text = "按行";
            this.radio_row.UseVisualStyleBackColor = true;
            // 
            // label_col
            // 
            this.label_col.AutoSize = true;
            this.label_col.Location = new System.Drawing.Point(84, 10);
            this.label_col.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_col.Name = "label_col";
            this.label_col.Size = new System.Drawing.Size(22, 15);
            this.label_col.TabIndex = 20;
            this.label_col.Text = "列";
            // 
            // textBox_col
            // 
            this.textBox_col.Location = new System.Drawing.Point(117, 4);
            this.textBox_col.Margin = new System.Windows.Forms.Padding(4);
            this.textBox_col.MaxLength = 3;
            this.textBox_col.Name = "textBox_col";
            this.textBox_col.Size = new System.Drawing.Size(44, 25);
            this.textBox_col.TabIndex = 18;
            this.textBox_col.Text = "3";
            this.textBox_col.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_col_KeyPress);
            this.textBox_col.Leave += new System.EventHandler(this.textBox_col_Leave);
            // 
            // label_row
            // 
            this.label_row.AutoSize = true;
            this.label_row.Location = new System.Drawing.Point(1, 10);
            this.label_row.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_row.Name = "label_row";
            this.label_row.Size = new System.Drawing.Size(22, 15);
            this.label_row.TabIndex = 18;
            this.label_row.Text = "行";
            // 
            // textBox_row
            // 
            this.textBox_row.Location = new System.Drawing.Point(35, 4);
            this.textBox_row.Margin = new System.Windows.Forms.Padding(4);
            this.textBox_row.MaxLength = 4;
            this.textBox_row.Name = "textBox_row";
            this.textBox_row.Size = new System.Drawing.Size(44, 25);
            this.textBox_row.TabIndex = 17;
            this.textBox_row.Text = "3";
            this.textBox_row.Enter += new System.EventHandler(this.textBox_row_Enter);
            this.textBox_row.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_row_KeyPress);
            this.textBox_row.Leave += new System.EventHandler(this.textBox_row_Leave);
            // 
            // toolTip1
            // 
            this.toolTip1.IsBalloon = true;
            // 
            // textBox_HeaderNameEdit
            // 
            this.textBox_HeaderNameEdit.BackColor = System.Drawing.Color.Yellow;
            this.textBox_HeaderNameEdit.Location = new System.Drawing.Point(25, 64);
            this.textBox_HeaderNameEdit.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox_HeaderNameEdit.Name = "textBox_HeaderNameEdit";
            this.textBox_HeaderNameEdit.Size = new System.Drawing.Size(169, 25);
            this.textBox_HeaderNameEdit.TabIndex = 27;
            this.toolTip1.SetToolTip(this.textBox_HeaderNameEdit, "修改列名，【回车】确认修改，Esc键取消");
            this.textBox_HeaderNameEdit.Visible = false;
            this.textBox_HeaderNameEdit.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_HeaderNameEdit_KeyPress);
            this.textBox_HeaderNameEdit.Leave += new System.EventHandler(this.textBox_HeaderNameEdit_Leave);
            // 
            // ExcelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(931, 462);
            this.Controls.Add(this.textBox_HeaderNameEdit);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dataGridView1);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "ExcelForm";
            this.Text = "生成表格（售后请联系qq：985012864）";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ExcelForm_FormClosed);
            this.Load += new System.EventHandler(this.ExcelForm_Load);
            this.MouseLeave += new System.EventHandler(this.ExcelForm_MouseLeave);
            this.Resize += new System.EventHandler(this.ExcelForm_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label_colWidth;
        private System.Windows.Forms.TextBox textBox_col;
        private System.Windows.Forms.Label label_rowHeight;
        private System.Windows.Forms.TextBox textBox_rowHeight;
        private System.Windows.Forms.RadioButton radio_col;
        private System.Windows.Forms.RadioButton radio_row;
        private System.Windows.Forms.Label label_col;
        //private System.Windows.Forms.TextBox textBox_colWidth;
        private System.Windows.Forms.Label label_row;
        private System.Windows.Forms.TextBox textBox_row;
        private System.Windows.Forms.CheckBox checkBox_dhcl;
        private System.Windows.Forms.Button button_define;
        private System.Windows.Forms.TextBox textBox_colWidth;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button button_selt;
        private System.Windows.Forms.TextBox textBox_HeaderNameEdit;
        private System.Windows.Forms.ComboBox cboPaiXu;
    }
}