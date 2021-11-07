using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace FormExcel
{
    class Table
    {
        private const int start_col = 14;
        private const int start_row = 28;

        public int col_count;
        public int row_count;

        public static List<List<Cell>> border = new List<List<Cell>>();
        public Dictionary<string, string> dictionary = new Dictionary<string, string>();

        private const string cellPattern = @"[A-Z]+[0-9]+";
        Regex regex = new Regex(cellPattern, RegexOptions.IgnoreCase);

        public Table()
        {
            col_count = start_col;
            row_count = start_row;
            New_Row();
        }

        public void SetTable(int columns, int rows)
        {
            Clear();
            col_count = columns;
            row_count = rows;
            New_Row();
        }

        public void New_Row()
        {
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
            border.Clear();
            dictionary.Clear();
            row_count = 0;
            col_count = 0;
        }

        public void ChangeCellWithAllPointers(int row, int column, string expression, DataGridView dataGridView)
        {
            var borderi = border[row][column];
            borderi.DeletePointersAndReferences();
            borderi.Expression = expression;
            borderi.new_references.Clear();

            if (!string.IsNullOrWhiteSpace(expression))
            {
                if (expression[0] != '=')
                {
                    borderi.Value = expression;
                    dictionary[FullName(row, column)] = expression;
                    foreach (Cell cell in borderi.pointer)
                    {
                        RefreshCellAndPointers(cell, dataGridView);
                    }
                    return;
                }
            }

            string new_expression = ConvertReferences(row, column, expression);

            if (!string.IsNullOrWhiteSpace(new_expression))
            {
                new_expression = new_expression.Remove(0, 1);
            }

            if (!borderi.CheckLoop(borderi.new_references))
            {
                MessageBox.Show("Помилка! Будь ласка, змініть свій вираз.");
                borderi.Expression = "";
                borderi.Value = "0";
                dataGridView[column, row].Value = "0";
                return;
            }

            borderi.AddPointersAndReferences();
            string value = Calculate(new_expression);
            if (value == "error")
            {
                MessageBox.Show("Помилка в клітинці - " + FullName(row, column));
                borderi.Expression = "";
                borderi.Value = "0";
                dataGridView[column, row].Value = "0";
                return;
            }

            borderi.Value = value;
            dictionary[FullName(row, column)] = value;
            foreach (Cell cell in borderi.pointer)
            {
                RefreshCellAndPointers(cell, dataGridView);
            }
        }

        private string FullName(int row, int column)
        {
            Cell cell = new Cell(row, column);
            return cell.Name;
        }

        public bool RefreshCellAndPointers(Cell cell, DataGridView dataGridView)
        {
            cell.new_references.Clear();
            string new_expression = ConvertReferences(cell.Row, cell.Column, cell.Expression);
            new_expression = new_expression.Remove(0, 1);
            string value = Calculate(new_expression);

            if (value == "error")
            {
                MessageBox.Show("Помилка в клітинці - " + cell.Name);
                return false;
            }

            border[cell.Row][cell.Column].Value = value;
            dictionary[FullName(cell.Row, cell.Column)] = value;
            dataGridView[cell.Column, cell.Row].Value = value;

            foreach (Cell point in cell.pointer)
            {
                if (!RefreshCellAndPointers(point, dataGridView))
                    return false;
            }
            return true;
        }

        public string ConvertReferences(int row, int column, string expression)
        {
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
            string value;
            if (dictionary.TryGetValue(m.Value, out value))
            {
                if(value == "")
                    return "0";
                else
                    return dictionary[m.Value];
            }
            return m.Value;
        }

        public string Calculate(string expression)
        {
            try
            {
                var result = Convert.ToString(Calculator.Evaluate(expression));
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
                    if (cell.Expression == "")
                        continue;
                    if (cell.Expression[0] == '=')
                    {
                        cell.references.AddRange(cell.new_references);
                    }
                }
            }
        }

        public void AddRow(DataGridView dataGridView)
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

        public void AddColumn(DataGridView dataGridView)
        {
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
                            if (cell_ref.Column == col_count)
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

        public bool DeleteRow(DataGridView dataGridView)
        {
            List<Cell> lastRow = new List<Cell>();
            List<string> notEmptyCells = new List<string>();

            if (row_count == 0)
                return false;

            int curCount = row_count - 1;

            for (int i = 0; i < col_count; i++)
            {
                string name = FullName(curCount, i);
                var dictionary_name = dictionary[name];

                if (dictionary_name != "0" && !string.IsNullOrWhiteSpace(dictionary_name))
                    notEmptyCells.Add(name);
                if (border[curCount][i].pointer.Count != 0)
                    lastRow.AddRange(border[curCount][i].pointer);
            }

            if (lastRow.Count != 0 || notEmptyCells.Count != 0)
            {
                string errorMessage = "";

                if (notEmptyCells.Count != 0)
                {
                    errorMessage = $"Немає порожніх клітинок: {string.Join(";", notEmptyCells.ToArray())} ";
                }

                if (lastRow.Count != 0)
                {
                    errorMessage += "Є клітинки, які вказують на клітинки з поточного рядка : ";
                    foreach (Cell cell in lastRow)
                    {
                        errorMessage += string.Join(";", cell.Name, " ");
                    }
                }

                errorMessage += "Ви впевнені, що хочете видалити цей рядок?";
                DialogResult res = MessageBox.Show(errorMessage, "Увага", MessageBoxButtons.YesNo);

                if (res == DialogResult.No)
                    return false;
            }

            for (int i = 0; i < col_count; i++)
            {
                string name = FullName(curCount, i);
                dictionary.Remove(name);
            }

            foreach (Cell cell in lastRow)
            {
                RefreshCellAndPointers(cell, dataGridView);
            }
            border.RemoveAt(curCount);
            row_count--;
            return true;
        }

        public bool DeleteColumn(DataGridView dataGridView)
        {
            List<Cell> lastCol = new List<Cell>();
            List<string> notEmptyCells = new List<string>();

            if (col_count == 0)
                return false;

            int curCount = col_count - 1;

            for (int i = 0; i < row_count; i++)
            {
                string name = FullName(i, curCount);
                var dictionary_name = dictionary[name];

                if (dictionary_name != "0" && !string.IsNullOrWhiteSpace(dictionary_name))
                    notEmptyCells.Add(name);
                if (border[i][curCount].pointer.Count != 0)
                    lastCol.AddRange(border[i][curCount].pointer);
            }

            if (lastCol.Count != 0 || notEmptyCells.Count != 0)
            {
                string errorMessage = "";

                if (notEmptyCells.Count != 0)
                {
                    errorMessage = $"Немає порожніх стовпчиків: {string.Join(";", notEmptyCells.ToArray())}";
                }

                if (lastCol.Count != 0)
                {
                    errorMessage += "Є клітинки, які вказують на клітинки з поточного стовпця: ";
                    foreach (Cell cell in lastCol)
                        errorMessage += string.Join(";", cell.Name);
                }

                errorMessage += "Ви впевнені, що хочете видалити цей стовпчик?";
                DialogResult res = MessageBox.Show(errorMessage, "Увага", MessageBoxButtons.YesNo);

                if (res == DialogResult.No)
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

        public void Open(int r, int c, System.IO.StreamReader streamreader, DataGridView dataGridView)
        {
            for (int i = 0; i < r; i++)
            {
                for (int j = 0; j < c; j++)
                {
                    string index = streamreader.ReadLine();
                    string expression = streamreader.ReadLine();
                    string value = streamreader.ReadLine();

                    if (expression != "")
                        dictionary[index] = value;
                    else
                        dictionary[index] = "";

                    int refCount = Convert.ToInt32(streamreader.ReadLine());
                    List<Cell> newRef = new List<Cell>();
                    string refer;

                    for (int k = 0; k < refCount; k++)
                    {
                        refer = streamreader.ReadLine();
                        if (NumberCell.FromIndexSystem(refer).row < row_count && NumberCell.FromIndexSystem(refer).column < col_count)
                            newRef.Add(border[NumberCell.FromIndexSystem(refer).row][NumberCell.FromIndexSystem(refer).column]);
                    }

                    int pointCount = Convert.ToInt32(streamreader.ReadLine());
                    List<Cell> newPoint = new List<Cell>();
                    string point;

                    for (int k = 0; k < pointCount; k++)
                    {
                        point = streamreader.ReadLine();
                        newPoint.Add(border[NumberCell.FromIndexSystem(point).row][NumberCell.FromIndexSystem(point).column]);
                    }
                    border[i][j].SetCell(expression, value, newRef, newPoint);

                    int curCol = border[i][j].Column;
                    int curRow = border[i][j].Row;
                    dataGridView[curCol, curRow].Value = dictionary[index];
                }
            }
        }

        public void Save(System.IO.StreamWriter streamwriter)
        {
            streamwriter.WriteLine(row_count);
            streamwriter.WriteLine(col_count);

            foreach (List<Cell> list in border)
            {
                foreach (Cell cell in list)
                {
                    streamwriter.WriteLine(cell.Name);
                    streamwriter.WriteLine(cell.Expression);
                    streamwriter.WriteLine(cell.Value);

                    if (cell.references == null)
                        streamwriter.WriteLine("0");
                    else
                    {
                        streamwriter.WriteLine(cell.references.Count);
                        foreach (Cell refCell in cell.references)
                            streamwriter.WriteLine(refCell.Name);
                    }
                    if (cell.pointer == null)
                        streamwriter.WriteLine("0");
                    else
                    {
                        streamwriter.WriteLine(cell.pointer.Count);
                        foreach (Cell pointCell in cell.pointer)
                            streamwriter.WriteLine(pointCell.Name);
                    }
                }
            }
        }
    }
}
