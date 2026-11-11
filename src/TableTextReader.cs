using System;
using System.Collections.Generic;
using System.Text;
using Tabula;
using UglyToad.PdfPig.Content;

namespace PdfReportExtractor
{
    // Compiled during packaging, not at runtime. PdfPig supplies original glyphs;
    // Tabula supplies geometry only, so no word deduplication alters source text.
    public static class TableTextReader
    {
        private sealed class Glyph
        {
            public string Text;
            public double X, Y;
            public int Order;
        }
        private sealed class Row
        {
            public double Top = double.MinValue, Bottom = double.MaxValue;
            public Cell[] Cells;
            public List<Glyph>[] Letters;
        }

        public static string[][] Read(IReadOnlyList<Letter> letters, Table table)
        {
            var rows = new Row[table.Rows.Count];
            for (int r = 0; r < rows.Length; r++)
            {
                var cells = table.Rows[r];
                var row = new Row { Cells = new Cell[cells.Count], Letters = new List<Glyph>[cells.Count] };
                for (int c = 0; c < cells.Count; c++)
                {
                    row.Cells[c] = cells[c];
                    row.Letters[c] = new List<Glyph>();
                    row.Top = Math.Max(row.Top, cells[c].Top);
                    row.Bottom = Math.Min(row.Bottom, cells[c].Bottom);
                }
                rows[r] = row;
            }
            for (int i = 0; i < letters.Count; i++)
            {
                Letter letter = letters[i];
                double x = letter.StartBaseLine.X, y = letter.StartBaseLine.Y;
                for (int r = 0; r < rows.Length; r++)
                {
                    Row row = rows[r];
                    if (y >= row.Top || y < row.Bottom) continue;
                    for (int c = 0; c < row.Cells.Length; c++)
                    {
                        Cell cell = row.Cells[c];
                        if (x >= cell.Left && x < cell.Right)
                        {
                            row.Letters[c].Add(new Glyph { Text = letter.Value, X = x, Y = y, Order = i });
                            break;
                        }
                    }
                    break;
                }
            }
            var result = new string[rows.Length][];
            for (int r = 0; r < rows.Length; r++)
            {
                result[r] = new string[rows[r].Cells.Length];
                for (int c = 0; c < result[r].Length; c++) result[r][c] = Join(rows[r].Letters[c]);
            }
            return result;
        }

        private static string Join(List<Glyph> glyphs)
        {
            glyphs.Sort(delegate(Glyph a, Glyph b) { int y = b.Y.CompareTo(a.Y); return y != 0 ? y : a.Order.CompareTo(b.Order); });
            var result = new List<string>();
            int first = 0;
            while (first < glyphs.Count)
            {
                int end = first + 1;
                while (end < glyphs.Count && Math.Abs(glyphs[end].Y - glyphs[first].Y) <= 0.8) end++;
                var line = glyphs.GetRange(first, end - first);
                line.Sort(delegate(Glyph a, Glyph b) { int x = a.X.CompareTo(b.X); return x != 0 ? x : a.Order.CompareTo(b.Order); });
                var text = new StringBuilder();
                foreach (Glyph glyph in line) text.Append(glyph.Text);
                result.Add(text.ToString().TrimEnd());
                first = end;
            }
            return String.Join("\n", result).Trim();
        }
    }
}
