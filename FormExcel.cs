using System;
using System.IO;
using System.Windows.Forms;

namespace FormExcel
{
	public partial class FormExcel : Form
	{
		Table table = new Table();
		public FormExcel()
		{
			InitializeComponent();
			WindowState = FormWindowState.Maximized;
			InitTable(table.row_count, table.col_count);
		}
        private void InitTable(int row, int column)
        {
            dataGridView.ColumnHeadersVisible = true;
            dataGridView.RowHeadersVisible = true;
            dataGridView.ColumnCount = column;
            for (int i = 0; i < column; i++)
            {
                string columnName = NumberCell.ToIndexSystem(i);
                dataGridView.Columns[i].Name = columnName;
                dataGridView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            for (int i = 0; i < row; ++i)
            {
                dataGridView.Rows.Add("");
                dataGridView.Rows[i].HeaderCell.Value = (i).ToString();
            }

            dataGridView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            table.SetTable(column, row);
        }


        private bool IsEnoughParen(string s)
        {
            int pparen = 0;
            int lparen = 0;
            foreach (char x in s)
            {
                if (x == '(') lparen++;
                if (x == ')') pparen++;
            }
            return lparen == pparen;

        }

        private void InitializeDataGridView(int rows, int columns)
        {
            dataGridView.ColumnHeadersVisible = true;
            dataGridView.RowHeadersVisible = true;
            dataGridView.ColumnCount = columns;
            for (int i = 0; i < columns; i++)
            {
                string ColumnName = NumberCell.ToIndexSystem(i);
                dataGridView.Columns[i].Name = ColumnName;
                dataGridView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            for (int i = 0; i < rows; i++)
            {
                dataGridView.Rows.Add("");
                dataGridView.Rows[i].HeaderCell.Value = (i).ToString();
            }
            dataGridView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            table.SetTable(columns, rows);
        }

        private void btnCalculate_Click_1(object sender, EventArgs e)
        {
            int column = dataGridView.SelectedCells[0].ColumnIndex;
            int row = dataGridView.SelectedCells[0].RowIndex;
            string expression = textBoxFE.Text;
            if (IsEnoughParen(expression))
            {
                if (expression == "") return;
                table.ChangeCellWithAllPointers(row, column, expression, dataGridView);
                dataGridView[column, row].Value = Table.border[row][column].Value;
            }
            else
            {
                MessageBox.Show("Помилка в кількості дужок!");
            }

        }

        /*private void dataGridView_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            int columns = dataGridView.SelectedCells[0].ColumnIndex;
            int row = dataGridView.SelectedCells[0].RowIndex;
            string expression = Table.border[row][columns].Exp;
            string value = Table.border[row][columns].Value;
            textBoxFE.Text = expression;
            textBoxFE.Focus();
        }*/

		private void btnAddRow_Click(object sender, EventArgs e)
		{
            DataGridViewRow row = new System.Windows.Forms.DataGridViewRow();
            if (dataGridView.Columns.Count == 0)
            {
                MessageBox.Show("Немає стовпчиків!");
                return;
            }
            dataGridView.Rows.Add(row);
            dataGridView.Rows[table.row_count].HeaderCell.Value = (table.row_count).ToString();
            table.AddRow(dataGridView);
        }

		private void btnAddColumn_Click(object sender, EventArgs e)
		{
            string name = NumberCell.ToIndexSystem(table.col_count);
            dataGridView.Columns.Add(name, name);
            table.AddColumn(dataGridView);
        }

		private void btnDeleteRow_Click(object sender, EventArgs e)
		{
            int curRow = table.row_count - 1;
            if (!table.DeleteRow(dataGridView))
                return;
            dataGridView.Rows.RemoveAt(curRow);
        }

		private void btnDeleteColumn_Click(object sender, EventArgs e)
		{
            int curCol = table.col_count - 1;
            if (!table.DeleteColumn(dataGridView))
                return;
            dataGridView.Columns.RemoveAt(curCol);
        }

		private void openToolStripMenuItem_Click_1(object sender, EventArgs e)
		{
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "TableFile|*.txt";
            openFileDialog.Title = "Open Table File";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;
            StreamReader sr = new StreamReader(openFileDialog.FileName);
            table.Clear();
            dataGridView.Rows.Clear();
            dataGridView.Columns.Clear();
            int row;
            int column;
            Int32.TryParse(sr.ReadLine(), out row);
            Int32.TryParse(sr.ReadLine(), out column);
            InitializeDataGridView(row, column);
            table.Open(row, column, sr, dataGridView);
            sr.Close();
        }

		private void saveToolStripMenuItem_Click_1(object sender, EventArgs e)
		{
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "TableFile|*.txt";
            saveFileDialog.Title = "Save table file";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                FileStream fs = (FileStream)saveFileDialog.OpenFile();
                StreamWriter sw = new StreamWriter(fs);
                table.Save(sw);
                sw.Close();
                fs.Close();
            }
        }

		private void helpToolStripMenuItem_Click_1(object sender, EventArgs e)
		{
            System.Windows.Forms.MessageBox.Show("Доступні функції: \n 1. + , - , * , / \n 2. mod, div \n 3. ++, -- \n 4. MAX, MIN");
        }

		private void exitToolStripMenuItem_Click_1(object sender, EventArgs e)
		{
            string errorMessage = "";
            errorMessage += "Ви впевнені, що хочете вийти?";
            System.Windows.Forms.DialogResult res = System.Windows.Forms.MessageBox.Show(errorMessage, "Вийти", System.Windows.Forms.MessageBoxButtons.YesNo);
            if (res == System.Windows.Forms.DialogResult.Yes)
            Application.Exit();
        }

		private void dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
		{
            int columns = dataGridView.SelectedCells[0].ColumnIndex;
            int row = dataGridView.SelectedCells[0].RowIndex;
            string expression = Table.border[row][columns].Exp;
            string value = Table.border[row][columns].Value;
            textBoxFE.Text = expression;
            textBoxFE.Focus();
        }

		private void FormExcel_Load(object sender, EventArgs e)
		{

		}
	}
}
