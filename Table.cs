using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace FormExcel
{
    class Table
    {
        private const int StartCol = 13;
        private const int StartRow = 10;

        public int col_count;
        public int row_count;

        public static List<List<Cell>> border = new List<List<Cell>>();
        public Dictionary<string, string> dictionary = new Dictionary<string, string>();
        public Table()
        {
            col_count = StartCol;
            row_count = StartRow;
            for (int i = 0; i < row_count; ++i)
            {
                List<Cell> NewRow = new List<Cell>();
                for (int j = 0; j < col_count; ++j)
                {
                    NewRow.Add(new Cell(i, j));
                    dictionary.Add(NewRow.Last().Name, "");
                }
                border.Add(NewRow);
            }
        }
        public void SetTable(int columns, int rows)
        {
            Clear();
            col_count = columns;
            row_count = rows;
            for (int i = 0; i < row_count; ++i)
            {
                List<Cell> NewRow = new List<Cell>();
                for (int j = 0; j < col_count; ++j)
                {
                    NewRow.Add(new Cell(i, j));
                    dictionary.Add(NewRow.Last().Name, "");
                }
                border.Add(NewRow);
            }
        }
        public void Clear()
        {
            foreach (List<Cell> list in border)
            {
                list.Clear();
            }
            border.Clear();
            dictionary.Clear();
            row_count = 0;
            col_count = 0;
        }

        public void ChangeCellWithAllPointers(int row, int column, string expression, System.Windows.Forms.DataGridView dataGridView)
        {
            border[row][column].DeletePointersAndReferences();
            border[row][column].Exp = expression;
            border[row][column].new_references.Clear();

            if (expression != "")
            {
                if (expression[0] != '=')
                {
                    border[row][column].Value = expression;
                    dictionary[FullName(row, column)] = expression;
                    foreach (Cell cell in border[row][column].pointer)
                    {
                        RefreshCellAndPointers(cell, dataGridView);
                    }
                    return;
                }
            }

            string new_expression = ConvertReferences(row, column, expression);

            if (new_expression != "")
            {
                new_expression = new_expression.Remove(0, 1);
            }

            if (!border[row][column].CheckLoop(border[row][column].new_references))
            {
                System.Windows.Forms.MessageBox.Show("Помилка! Будь ласка, змініть свій вираз.");
                border[row][column].Exp = "";
                border[row][column].Value = "0";
                dataGridView[column, row].Value = "0";
                return;
            }

            border[row][column].AddPointersAndReferences();
            string value = Calculate(new_expression);
            if (value == "помилка")
            {
                System.Windows.Forms.MessageBox.Show("Помилка в клітинці - " + FullName(row, column));
                border[row][column].Exp = "";
                border[row][column].Value = "0";
                dataGridView[column, row].Value = "0";
                return;
            }

            border[row][column].Value = value;
            dictionary[FullName(row, column)] = value;
            foreach (Cell cell in border[row][column].pointer)
                RefreshCellAndPointers(cell, dataGridView);
        }
        private string FullName(int row, int column)
        {
            Cell cell = new Cell(row, column);
            return cell.Name;
        }
        public bool RefreshCellAndPointers(Cell cell, System.Windows.Forms.DataGridView dataGridView)
        {
            cell.new_references.Clear();
            string new_expression = ConvertReferences(cell.Row, cell.Сolumn, cell.Exp);
            new_expression = new_expression.Remove(0, 1);
            string Value = Calculate(new_expression);

            if (Value == "помилка")
            {
                System.Windows.Forms.MessageBox.Show("Помилка в клітинці - " + cell.Name);
                return false;
            }

            border[cell.Row][cell.Сolumn].Value = Value;
            dictionary[FullName(cell.Row, cell.Сolumn)] = Value;
            dataGridView[cell.Сolumn, cell.Row].Value = Value;

            foreach (Cell point in cell.pointer)
            {
                if (!RefreshCellAndPointers(point, dataGridView))
                    return false;
            }
            return true;
        }
        public string ConvertReferences(int row, int column, string expression)
        {
            string cellPattern = @"[A-Z]+[0-9]+";
            Regex regex = new Regex(cellPattern, RegexOptions.IgnoreCase);
            TheIndex nums;

            foreach (Match match in regex.Matches(expression))
            {
                if (dictionary.ContainsKey(match.Value))
                {
                    nums = NumberCell.FromIndexSystem(match.Value);
                    border[row][column].new_references.Add(border[nums.row][nums.column]);
                }
            }

            MatchEvaluator evaluator = new MatchEvaluator(ReferencesToValue);
            string new_expression = regex.Replace(expression, evaluator);
            return new_expression;
        }

        public string ReferencesToValue(Match m)
        {
            if (dictionary.ContainsKey(m.Value))
                if (dictionary[m.Value] == "")
                    return "0";
                else
                    return dictionary[m.Value];
            return m.Value;
        }

        public string Calculate(string expression)
        {
            string result = null;
            try
            {
                result = Convert.ToString(Calculator.Evaluate(expression));
                if (result == "∞")
                {
                    MessageBox.Show("Помилка! Ділення на нуль!");
                    result = "";
                }
                return result;
            }
            catch
            {
                return "";
            }
        }

        public void RefreshReferences()
        {
            foreach (List<Cell> list in border)
            {
                foreach (Cell cell in list)
                {
                    if (cell.references != null)
                        cell.references.Clear();
                    if (cell.new_references != null)
                        cell.new_references.Clear();
                    if (cell.Exp == "")
                        continue;
                    string new_expression = cell.Exp;
                    if (cell.Exp[0] == '=')
                    {
                        new_expression = ConvertReferences(cell.Row, cell.Сolumn, cell.Exp);
                        cell.references.AddRange(cell.new_references);
                    }
                }
            }
        }

        public void AddRow(System.Windows.Forms.DataGridView dataGridView)
        {
            List<Cell> newRow = new List<Cell>();

            for (int i = 0; i < col_count; i++)
            {
                newRow.Add(new Cell(row_count, i));
                dictionary.Add(newRow.Last().Name, "");
            }
            border.Add(newRow);
            RefreshReferences();
            foreach (List<Cell> list in border)
            {
                foreach (Cell cell in list)
                {
                    if (cell.references != null)
                    {
                        foreach (Cell cell_ref in cell.references)
                        {
                            if (cell_ref.Row == row_count)
                            {
                                if (!cell_ref.pointer.Contains(cell))
                                    cell_ref.pointer.Add(cell);
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < col_count; i++)
            {
                ChangeCellWithAllPointers(row_count, i, "", dataGridView);
            }
            row_count++;
        }

        public void AddColumn(System.Windows.Forms.DataGridView dataGridView)
        {
            List<Cell> newCol = new List<Cell>();

            for (int i = 0; i < row_count; i++)
            {
                string name = FullName(i, col_count);
                border[i].Add(new Cell(i, col_count));
                dictionary.Add(name, "");
            }

            RefreshReferences();
            foreach (List<Cell> list in border)
            {
                foreach (Cell cell in list)
                {
                    if (cell.references != null)
                    {
                        foreach (Cell cell_ref in cell.references)
                        {
                            if (cell_ref.Сolumn == col_count)
                            {
                                if (!cell_ref.pointer.Contains(cell))
                                    cell_ref.pointer.Add(cell);
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < row_count; i++)
            {
                ChangeCellWithAllPointers(i, col_count, "", dataGridView);
            }
            col_count++;
        }

        public bool DeleteRow(System.Windows.Forms.DataGridView dataGridView)
        {
            List<Cell> lastRow = new List<Cell>();
            List<string> notEmptyCells = new List<string>();

            if (row_count == 0)
                return false;

            int curCount = row_count - 1;

            for (int i = 0; i < col_count; i++)
            {
                string name = FullName(curCount, i);

                if (dictionary[name] != "0" && dictionary[name] != "" && dictionary[name] != " ")
                    notEmptyCells.Add(name);
                if (border[curCount][i].pointer.Count != 0)
                    lastRow.AddRange(border[curCount][i].pointer);
            }

            if (lastRow.Count != 0 || notEmptyCells.Count != 0)
            {
                string errorMessage = "";

                if (notEmptyCells.Count != 0)
                {
                    errorMessage = "Немає порожніх клітинок: ";
                    errorMessage += string.Join(";", notEmptyCells.ToArray());
                    errorMessage += ' ';
                }

                if (lastRow.Count != 0)
                {
                    errorMessage += "Є клітинки, які вказують на клітинки з поточного рядка : ";
                    foreach (Cell cell in lastRow)
                    {
                        errorMessage += string.Join(";", cell.Name);
                        errorMessage += " ";
                    }
                }

                errorMessage += "Ви впевнені, що хочете видалити цей рядок?";
                System.Windows.Forms.DialogResult res = System.Windows.Forms.MessageBox.Show(errorMessage, "Увага", System.Windows.Forms.MessageBoxButtons.YesNo);

                if (res == System.Windows.Forms.DialogResult.No)
                    return false;
            }

            for (int i = 0; i < col_count; i++)
            {
                string name = FullName(curCount, i);
                dictionary.Remove(name);
            }

            foreach (Cell cell in lastRow)
                RefreshCellAndPointers(cell, dataGridView);
            border.RemoveAt(curCount);
            row_count--;
            return true;
        }

        public bool DeleteColumn(System.Windows.Forms.DataGridView dataGridView)
        {
            List<Cell> lastCol = new List<Cell>();
            List<string> notEmptyCells = new List<string>();

            if (col_count == 0)
                return false;

            int curCount = col_count - 1;

            for (int i = 0; i < row_count; i++)
            {
                string name = FullName(i, curCount);
                if (dictionary[name] != "0" && dictionary[name] != "" && dictionary[name] != " ")
                    notEmptyCells.Add(name);
                if (border[i][curCount].pointer.Count != 0)
                    lastCol.AddRange(border[i][curCount].pointer);
            }

            if (lastCol.Count != 0 || notEmptyCells.Count != 0)
            {
                string errorMessage = "";

                if (notEmptyCells.Count != 0)
                {
                    errorMessage = "Немає порожніх стовпчиків: ";
                    errorMessage += string.Join(";", notEmptyCells.ToArray());
                }

                if (lastCol.Count != 0)
                {
                    errorMessage += "Є клітинки, які вказують на клітинки з поточного стовпця: ";
                    foreach (Cell cell in lastCol)
                        errorMessage += string.Join(";", cell.Name);
                }

                errorMessage += "Ви впевнені, що хочете видалити цей стовпчик?";
                System.Windows.Forms.DialogResult res = System.Windows.Forms.MessageBox.Show(errorMessage, "Увага", System.Windows.Forms.MessageBoxButtons.YesNo);

                if (res == System.Windows.Forms.DialogResult.No)
                    return false;
            }

            for (int i = 0; i < row_count; i++)
            {
                string name = FullName(i, curCount);
                dictionary.Remove(name);
            }

            foreach (Cell cell in lastCol)
                RefreshCellAndPointers(cell, dataGridView);

            for (int i = 0; i < row_count; i++)
            {
                border[i].RemoveAt(curCount);
            }

            col_count--;
            return true;
        }
        public void Open(int r, int c, System.IO.StreamReader sr, System.Windows.Forms.DataGridView dataGridView)
        {
            for (int i = 0; i < r; i++)
            {
                for (int j = 0; j < c; j++)
                {
                    string index = sr.ReadLine();
                    string expression = sr.ReadLine();
                    string value = sr.ReadLine();

                    if (expression != "")
                        dictionary[index] = value;
                    else
                        dictionary[index] = "";

                    int refCount = Convert.ToInt32(sr.ReadLine());
                    List<Cell> newRef = new List<Cell>();
                    string refer;

                    for (int k = 0; k < refCount; k++)
                    {
                        refer = sr.ReadLine();
                        if (NumberCell.FromIndexSystem(refer).row < row_count && NumberCell.FromIndexSystem(refer).column < col_count)
                            newRef.Add(border[NumberCell.FromIndexSystem(refer).row][NumberCell.FromIndexSystem(refer).column]);
                    }

                    int pointCount = Convert.ToInt32(sr.ReadLine());
                    List<Cell> newPoint = new List<Cell>();
                    string point;

                    for (int k = 0; k < pointCount; k++)
                    {
                        point = sr.ReadLine();
                        newPoint.Add(border[NumberCell.FromIndexSystem(point).row][NumberCell.FromIndexSystem(point).column]);
                    }
                    border[i][j].SetCell(expression, value, newRef, newPoint);

                    int curCol = border[i][j].Сolumn;
                    int curRow = border[i][j].Row;
                    dataGridView[curCol, curRow].Value = dictionary[index];
                }
            }
        }
        public void Save(System.IO.StreamWriter sw)
        {
            sw.WriteLine(row_count);
            sw.WriteLine(col_count);

            foreach (List<Cell> list in border)
            {
                foreach (Cell cell in list)
                {
                    sw.WriteLine(cell.Name);
                    sw.WriteLine(cell.Exp);
                    sw.WriteLine(cell.Value);

                    if (cell.references == null)
                        sw.WriteLine("0");
                    else
                    {
                        sw.WriteLine(cell.references.Count);
                        foreach (Cell refCell in cell.references)
                            sw.WriteLine(refCell.Name);
                    }
                    if (cell.pointer == null)
                        sw.WriteLine("0");
                    else
                    {
                        sw.WriteLine(cell.pointer.Count);
                        foreach (Cell pointCell in cell.pointer)
                            sw.WriteLine(pointCell.Name);
                    }
                }
            }
        }
    }
}
