namespace FormExcel
{
	partial class FormExcel
	{
		/// <summary>
		///  Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		///  Clean up any resources being used.
		/// </summary>
		/// <param Name="disposing">true if managed resources should be disposed; otherwise, false.</param>
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
		///  Required method for Designer support - do not modify
		///  the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.dataGridView = new System.Windows.Forms.DataGridView();
			this.btnAddRow = new System.Windows.Forms.Button();
			this.btnAddColumn = new System.Windows.Forms.Button();
			this.btnDeleteRow = new System.Windows.Forms.Button();
			this.btnDeleteColumn = new System.Windows.Forms.Button();
			this.btnCalculate = new System.Windows.Forms.Button();
			this.textBoxFE = new System.Windows.Forms.TextBox();
			this.menuStrip = new System.Windows.Forms.MenuStrip();
			this.FileTool = new System.Windows.Forms.ToolStripMenuItem();
			this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.infoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.exitStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
			this.menuStrip.SuspendLayout();
			this.SuspendLayout();
			// 
			// dataGridView
			// 
			this.dataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView.Location = new System.Drawing.Point(13, 100);
			this.dataGridView.Name = "dataGridView";
			this.dataGridView.RowHeadersWidth = 51;
			this.dataGridView.RowTemplate.Height = 29;
			this.dataGridView.Size = new System.Drawing.Size(1289, 355);
			this.dataGridView.TabIndex = 0;
			this.dataGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellClick);
			// 
			// btnAddRow
			// 
			this.btnAddRow.BackColor = System.Drawing.SystemColors.Control;
			this.btnAddRow.Location = new System.Drawing.Point(13, 44);
			this.btnAddRow.Name = "btnAddRow";
			this.btnAddRow.Size = new System.Drawing.Size(159, 43);
			this.btnAddRow.TabIndex = 1;
			this.btnAddRow.Text = "Додати рядок";
			this.btnAddRow.UseVisualStyleBackColor = false;
			this.btnAddRow.Click += new System.EventHandler(this.btnAddRow_Click);
			// 
			// btnAddColumn
			// 
			this.btnAddColumn.BackColor = System.Drawing.SystemColors.Control;
			this.btnAddColumn.Location = new System.Drawing.Point(213, 44);
			this.btnAddColumn.Name = "btnAddColumn";
			this.btnAddColumn.Size = new System.Drawing.Size(161, 43);
			this.btnAddColumn.TabIndex = 2;
			this.btnAddColumn.Text = "Додати стовпчик";
			this.btnAddColumn.UseVisualStyleBackColor = false;
			this.btnAddColumn.Click += new System.EventHandler(this.btnAddColumn_Click);
			// 
			// btnDeleteRow
			// 
			this.btnDeleteRow.BackColor = System.Drawing.SystemColors.Control;
			this.btnDeleteRow.Location = new System.Drawing.Point(414, 44);
			this.btnDeleteRow.Name = "btnDeleteRow";
			this.btnDeleteRow.Size = new System.Drawing.Size(159, 43);
			this.btnDeleteRow.TabIndex = 3;
			this.btnDeleteRow.Text = "Видалити рядок";
			this.btnDeleteRow.UseVisualStyleBackColor = false;
			this.btnDeleteRow.Click += new System.EventHandler(this.btnDeleteRow_Click);
			// 
			// btnDeleteColumn
			// 
			this.btnDeleteColumn.BackColor = System.Drawing.SystemColors.Control;
			this.btnDeleteColumn.Location = new System.Drawing.Point(612, 44);
			this.btnDeleteColumn.Name = "btnDeleteColumn";
			this.btnDeleteColumn.Size = new System.Drawing.Size(167, 43);
			this.btnDeleteColumn.TabIndex = 4;
			this.btnDeleteColumn.Text = "Видалити стовпчик";
			this.btnDeleteColumn.UseVisualStyleBackColor = false;
			this.btnDeleteColumn.Click += new System.EventHandler(this.btnDeleteColumn_Click);
			// 
			// btnCalculate
			// 
			this.btnCalculate.BackColor = System.Drawing.SystemColors.Control;
			this.btnCalculate.Location = new System.Drawing.Point(1135, 45);
			this.btnCalculate.Name = "btnCalculate";
			this.btnCalculate.Size = new System.Drawing.Size(167, 42);
			this.btnCalculate.TabIndex = 5;
			this.btnCalculate.Text = "Порахувати";
			this.btnCalculate.UseVisualStyleBackColor = false;
			this.btnCalculate.Click += new System.EventHandler(this.btnCalculate_Click_1);
			// 
			// textBoxFE
			// 
			this.textBoxFE.Location = new System.Drawing.Point(866, 53);
			this.textBoxFE.Name = "textBoxFE";
			this.textBoxFE.Size = new System.Drawing.Size(223, 27);
			this.textBoxFE.TabIndex = 6;
			// 
			// menuStrip
			// 
			this.menuStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.FileTool});
			this.menuStrip.Location = new System.Drawing.Point(0, 0);
			this.menuStrip.Name = "menuStrip";
			this.menuStrip.Size = new System.Drawing.Size(1314, 28);
			this.menuStrip.TabIndex = 7;
			this.menuStrip.Text = "menuStrip1";
			// 
			// FileTool
			// 
			this.FileTool.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.saveToolStripMenuItem,
            this.helpToolStripMenuItem,
            this.infoToolStripMenuItem,
            this.exitStripMenuItem});
			this.FileTool.Name = "FileTool";
			this.FileTool.Size = new System.Drawing.Size(59, 24);
			this.FileTool.Text = "Файл";
			// 
			// openToolStripMenuItem
			// 
			this.openToolStripMenuItem.Name = "openToolStripMenuItem";
			this.openToolStripMenuItem.Size = new System.Drawing.Size(193, 26);
			this.openToolStripMenuItem.Text = "Відкрити";
			this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click_1);
			// 
			// saveToolStripMenuItem
			// 
			this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
			this.saveToolStripMenuItem.Size = new System.Drawing.Size(193, 26);
			this.saveToolStripMenuItem.Text = "Зберегти";
			this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click_1);
			// 
			// helpToolStripMenuItem
			// 
			this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
			this.helpToolStripMenuItem.Size = new System.Drawing.Size(193, 26);
			this.helpToolStripMenuItem.Text = "Допомога";
			this.helpToolStripMenuItem.Click += new System.EventHandler(this.helpToolStripMenuItem_Click_1);
			// 
			// infoToolStripMenuItem
			// 
			this.infoToolStripMenuItem.Name = "infoToolStripMenuItem";
			this.infoToolStripMenuItem.Size = new System.Drawing.Size(193, 26);
			this.infoToolStripMenuItem.Text = "Про програму";
			this.infoToolStripMenuItem.Click += new System.EventHandler(this.infoToolStripMenuItem_Click);
			// 
			// exitStripMenuItem
			// 
			this.exitStripMenuItem.Name = "exitStripMenuItem";
			this.exitStripMenuItem.Size = new System.Drawing.Size(193, 26);
			this.exitStripMenuItem.Text = "Вийти";
			this.exitStripMenuItem.Click += new System.EventHandler(this.exitStripMenuItem_Click);
			// 
			// openFileDialog
			// 
			this.openFileDialog.FileName = "openFileDialog";
			// 
			// FormExcel
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1314, 467);
			this.Controls.Add(this.textBoxFE);
			this.Controls.Add(this.btnCalculate);
			this.Controls.Add(this.btnDeleteColumn);
			this.Controls.Add(this.btnDeleteRow);
			this.Controls.Add(this.btnAddColumn);
			this.Controls.Add(this.btnAddRow);
			this.Controls.Add(this.dataGridView);
			this.Controls.Add(this.menuStrip);
			this.MainMenuStrip = this.menuStrip;
			this.Name = "FormExcel";
			this.Text = "Excel";
			this.Load += new System.EventHandler(this.FormExcel_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
			this.menuStrip.ResumeLayout(false);
			this.menuStrip.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.DataGridView dataGridView;
		private System.Windows.Forms.Button btnAddRow;
		private System.Windows.Forms.Button btnAddColumn;
		private System.Windows.Forms.Button btnDeleteRow;
		private System.Windows.Forms.Button btnDeleteColumn;
		private System.Windows.Forms.Button btnCalculate;
		private System.Windows.Forms.TextBox textBoxFE;
		private System.Windows.Forms.MenuStrip menuStrip;
		private System.Windows.Forms.ToolStripMenuItem FileTool;
		private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem infoToolStripMenuItem;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.SaveFileDialog saveFileDialog;
		private System.Windows.Forms.ToolStripMenuItem exitStripMenuItem;
	}
}
