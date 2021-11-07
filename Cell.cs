using System;
using System.Collections.Generic;

namespace FormExcel
{
	class Cell
	{
		public List<Cell> pointer = new List<Cell>();
		public List<Cell> references = new List<Cell>();
		public List<Cell> new_references = new List<Cell>();
		public Cell(int rows, int columns)
		{
			Row = rows;
			Column = columns;
			Name = NumberCell.ToIndexSystem(columns) + Convert.ToString(rows);
			Value = "0";
			Expression = "";
		}
		public string Name { get; set; }
		public string Value { get; set; }
		public string Expression { get; set; }
		public int Column { get; set; }
		public int Row { get; set; }
		public void SetCell(string exp, string val, List<Cell> a, List<Cell> b)
		{
			Value = val;
			Expression = exp;
			references.Clear();
			references.AddRange(a);
			pointer.Clear();
			pointer.AddRange(b);
		}
		public bool CheckLoop(List<Cell> list)
		{
			foreach (Cell cell in list)
			{
				if (cell.Name == Name)
					return false;
			}
			foreach (Cell point in pointer)
			{
				foreach (Cell cell in list)
				{
					if (cell.Name == point.Name)
					{
						return false;
					}
				}
				if (!point.CheckLoop(list)) return false;
			}
			return true;
		}

		public void AddPointersAndReferences()
		{
			foreach (Cell point in new_references)
			{
				point.pointer.Add(this);
			}
			references = new_references;
		}

		public void DeletePointersAndReferences()
		{
			if (references != null)
			{
				foreach (Cell point in references)
				{
					point.pointer.Remove(this);
				}
				references = null;
			}
		}
	}
}
