using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

namespace PdfToExcelWithWatermark
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string pdfPath = "input.pdf";
                string excelPath = "output.xlsx";
                string finalPath = "output_watermarked.xlsx";

                Console.WriteLine("Step 1: Convert PDF to Excel...");
                ConvertPdfToExcel(pdfPath, excelPath);

                Console.WriteLine("Step 2: Add watermark...");

                if (File.Exists("watermark_transparent.png")) File.Delete("watermark_transparent.png");
                ProcessingImage.MakeTransparentImage("watermark.png", "watermark_transparent.png", 0.2f);

                AddSingleImageWatermark(excelPath, finalPath, "watermark_transparent.png");

                Console.WriteLine("✅ Done successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Error: " + ex.Message);
            }
        }

        sealed class TextCell
        {
            public string Text = "";
            public double Left, Right, Top, Bottom;
            public bool IsNumeric;
            public bool IsBold;
            public double FontSize;
            public bool IsMultiLine;
            public double Height => Top - Bottom;
        }

        enum RowKind { Empty, Title, Section, SubSection, ColumnHeader, Data }

        static void ConvertPdfToExcel(string pdfPath, string excelPath)
        {
            if (File.Exists(excelPath)) File.Delete(excelPath);

            using var workbook = new XLWorkbook();
            using var pdf = PdfDocument.Open(pdfPath);

            int pageNumber = 1;
            foreach (Page page in pdf.GetPages())
            {
                var sheet = workbook.AddWorksheet($"Page{pageNumber++}");

                var words = NearestNeighbourWordExtractor.Instance
                    .GetWords(page.Letters)
                    .Where(w => !string.IsNullOrWhiteSpace(w.Text))
                    .ToList();

                if (words.Count == 0) continue;

                double medianFont = MedianFontSize(words);

                var lines = GroupWordsIntoLines(words);
                var rowsOfCells = lines.Select(BuildCellsFromLine).Where(r => r.Count > 0).ToList();
                var columns = DetectColumns(rowsOfCells);
                var grid = BuildGrid(rowsOfCells, columns);
                grid = MergeMultiLineHeaders(grid);

                WriteGrid(sheet, grid, medianFont);
            }

            workbook.SaveAs(excelPath);
        }

        static List<List<Word>> GroupWordsIntoLines(IReadOnlyList<Word> words)
        {
            var lines = new List<List<Word>>();
            foreach (var word in words.OrderByDescending(w => w.BoundingBox.Bottom))
            {
                double tolerance = Math.Max(1.5, word.BoundingBox.Height * 0.4);
                double center = (word.BoundingBox.Top + word.BoundingBox.Bottom) / 2.0;

                var matched = lines.FirstOrDefault(l =>
                {
                    double lineCenter = l.Average(w => (w.BoundingBox.Top + w.BoundingBox.Bottom) / 2.0);
                    return Math.Abs(lineCenter - center) <= tolerance;
                });

                if (matched != null) matched.Add(word);
                else lines.Add(new List<Word> { word });
            }

            foreach (var line in lines) line.Sort((a, b) => a.BoundingBox.Left.CompareTo(b.BoundingBox.Left));
            return lines;
        }

        static List<TextCell> BuildCellsFromLine(List<Word> line)
        {
            var cells = new List<TextCell>();
            if (line.Count == 0) return cells;

            var current = new List<Word> { line[0] };

            for (int i = 1; i < line.Count; i++)
            {
                var prev = line[i - 1];
                var curr = line[i];
                double gap = curr.BoundingBox.Left - prev.BoundingBox.Right;
                double charWidth = prev.BoundingBox.Width / Math.Max(1, prev.Text.Length);
                double threshold = Math.Max(charWidth * 2.0, 5.0);

                if (gap > threshold)
                {
                    cells.Add(MakeCell(current));
                    current = new List<Word>();
                }
                current.Add(curr);
            }
            if (current.Count > 0) cells.Add(MakeCell(current));
            return cells;
        }

        static TextCell MakeCell(List<Word> words)
        {
            var text = string.Join(" ", words.Select(w => w.Text));
            var letters = words.SelectMany(w => w.Letters).ToList();
            double fontSize = letters.Count > 0 ? letters.Average(l => (double)l.PointSize) : 0;
            bool isBold = letters.Any(l => IsBoldFont(l.FontName));

            return new TextCell
            {
                Text = text,
                Left = words.Min(w => w.BoundingBox.Left),
                Right = words.Max(w => w.BoundingBox.Right),
                Top = words.Max(w => w.BoundingBox.Top),
                Bottom = words.Min(w => w.BoundingBox.Bottom),
                IsNumeric = LooksLikeNumber(text),
                IsBold = isBold,
                FontSize = fontSize
            };
        }

        static bool IsBoldFont(string? fontName)
        {
            if (string.IsNullOrEmpty(fontName)) return false;
            return fontName.Contains("Bold", StringComparison.OrdinalIgnoreCase)
                || fontName.Contains("Black", StringComparison.OrdinalIgnoreCase)
                || fontName.Contains("Heavy", StringComparison.OrdinalIgnoreCase);
        }

        static double MedianFontSize(IReadOnlyList<Word> words)
        {
            var sizes = words.SelectMany(w => w.Letters)
                             .Select(l => (double)l.PointSize)
                             .Where(s => s > 0)
                             .OrderBy(s => s)
                             .ToList();
            if (sizes.Count == 0) return 10;
            return sizes[sizes.Count / 2];
        }

        static bool LooksLikeNumber(string text)
        {
            var stripped = text.Trim();
            if (string.IsNullOrEmpty(stripped)) return false;
            return double.TryParse(stripped, NumberStyles.Number | NumberStyles.AllowLeadingSign,
                CultureInfo.InvariantCulture, out _);
        }

        static List<double> DetectColumns(List<List<TextCell>> rows)
        {
            const double clusterEpsilon = 8.0;
            var lefts = rows.SelectMany(r => r.Select(c => c.Left)).OrderBy(x => x).ToList();
            if (lefts.Count == 0) return new List<double>();

            var columns = new List<double>();
            double clusterAnchor = lefts[0];
            double clusterSum = lefts[0];
            int clusterCount = 1;

            for (int i = 1; i < lefts.Count; i++)
            {
                if (lefts[i] - clusterAnchor <= clusterEpsilon)
                {
                    clusterSum += lefts[i];
                    clusterCount++;
                }
                else
                {
                    columns.Add(clusterSum / clusterCount);
                    clusterAnchor = lefts[i];
                    clusterSum = lefts[i];
                    clusterCount = 1;
                }
            }
            columns.Add(clusterSum / clusterCount);
            return columns;
        }

        static List<List<TextCell?>> BuildGrid(List<List<TextCell>> rows, List<double> columns)
        {
            var grid = new List<List<TextCell?>>();
            foreach (var row in rows)
            {
                var gridRow = new List<TextCell?>(new TextCell?[columns.Count]);
                foreach (var cell in row)
                {
                    int col = NearestColumn(cell.Left, columns);
                    if (gridRow[col] == null)
                        gridRow[col] = cell;
                    else
                    {
                        gridRow[col]!.Text += " " + cell.Text;
                        gridRow[col]!.Right = Math.Max(gridRow[col]!.Right, cell.Right);
                        gridRow[col]!.IsNumeric = LooksLikeNumber(gridRow[col]!.Text);
                    }
                }
                grid.Add(gridRow);
            }
            return grid;
        }

        static int NearestColumn(double left, List<double> columns)
        {
            int best = 0;
            double bestDist = double.MaxValue;
            for (int i = 0; i < columns.Count; i++)
            {
                double d = Math.Abs(columns[i] - left);
                if (d < bestDist) { bestDist = d; best = i; }
            }
            return best;
        }

        static List<List<TextCell?>> MergeMultiLineHeaders(List<List<TextCell?>> grid)
        {
            int r = 1;
            while (r < grid.Count)
            {
                var current = grid[r];
                var above = grid[r - 1];

                var currentFilled = current.Select((c, i) => new { c, i }).Where(x => x.c != null).ToList();
                var aboveFilled = above.Select((c, i) => new { c, i }).Where(x => x.c != null).ToList();

                bool merge = currentFilled.Count > 0
                    && aboveFilled.Count > 0
                    && currentFilled.Count < aboveFilled.Count
                    && currentFilled.All(x => !x.c!.IsNumeric)
                    && currentFilled.All(x => above[x.i] != null && !above[x.i]!.IsNumeric)
                    && VerticallyClose(above, current);

                if (merge)
                {
                    foreach (var x in currentFilled)
                    {
                        above[x.i]!.Text += "\n" + x.c!.Text;
                        above[x.i]!.Bottom = x.c.Bottom;
                        above[x.i]!.IsMultiLine = true;
                        above[x.i]!.IsBold = above[x.i]!.IsBold || x.c.IsBold;
                    }
                    grid.RemoveAt(r);
                }
                else
                {
                    r++;
                }
            }
            return grid;
        }

        static bool VerticallyClose(List<TextCell?> above, List<TextCell?> current)
        {
            var aboveBottoms = above.Where(c => c != null).Select(c => c!.Bottom).ToList();
            var currentTops = current.Where(c => c != null).Select(c => c!.Top).ToList();
            var heights = above.Where(c => c != null).Select(c => c!.Height).Where(h => h > 0).ToList();
            if (aboveBottoms.Count == 0 || currentTops.Count == 0 || heights.Count == 0) return false;
            double gap = aboveBottoms.Min() - currentTops.Max();
            double avgHeight = heights.Average();
            return gap <= avgHeight * 0.8;
        }

        static void WriteGrid(IXLWorksheet sheet, List<List<TextCell?>> grid, double medianFont)
        {
            int columnCount = grid.Count > 0 ? grid[0].Count : 0;
            if (columnCount == 0) return;

            var rowKinds = grid.Select(r => ClassifyRow(r, medianFont)).ToList();

            for (int r = 0; r < grid.Count; r++)
            {
                var row = grid[r];
                var kind = rowKinds[r];

                if (kind == RowKind.Title || kind == RowKind.Section || kind == RowKind.SubSection)
                {
                    var bannerCell = row.FirstOrDefault(c => c != null);
                    if (bannerCell == null) continue;
                    var xlFirst = sheet.Cell(r + 1, 1);
                    xlFirst.SetValue(bannerCell.Text);
                    xlFirst.Style.NumberFormat.Format = "@";
                    continue;
                }

                for (int c = 0; c < row.Count; c++)
                {
                    var cell = row[c];
                    if (cell == null) continue;
                    var xlCell = sheet.Cell(r + 1, c + 1);
                    xlCell.SetValue(cell.Text);
                    xlCell.Style.NumberFormat.Format = "@";
                    if (cell.Text.Contains('\n'))
                        xlCell.Style.Alignment.WrapText = true;
                }
            }

            ApplyStyling(sheet, grid, rowKinds, columnCount);
        }

        static RowKind ClassifyRow(List<TextCell?> row, double medianFont)
        {
            var nonEmpty = row.Where(c => c != null).Cast<TextCell>().ToList();
            if (nonEmpty.Count == 0) return RowKind.Empty;

            bool hasMultiLine = nonEmpty.Any(c => c.IsMultiLine);
            bool allBold = nonEmpty.All(c => c.IsBold);
            bool hasNumeric = nonEmpty.Any(c => c.IsNumeric);
            double avgFont = nonEmpty.Average(c => c.FontSize);
            bool veryLarge = avgFont > medianFont * 1.4;
            bool largerThanBody = avgFont > medianFont * 1.1;

            if (hasMultiLine) return RowKind.ColumnHeader;
            if (veryLarge && nonEmpty.Count <= 2) return RowKind.Title;
            if (allBold && !hasNumeric && nonEmpty.Count == 1 && largerThanBody) return RowKind.Section;
            if (allBold && !hasNumeric && nonEmpty.Count == 1) return RowKind.SubSection;
            return RowKind.Data;
        }

        static void ApplyStyling(IXLWorksheet sheet, List<List<TextCell?>> grid, List<RowKind> rowKinds, int columnCount)
        {
            var titleFill = XLColor.FromHtml("#1F4E79");
            var sectionFill = XLColor.FromHtml("#2E75B6");
            var subSectionFill = XLColor.FromHtml("#BDD7EE");
            var columnHeaderFill = XLColor.FromHtml("#D9D9D9");
            var labelFill = XLColor.FromHtml("#F2F2F2");
            var borderColor = XLColor.FromHtml("#BFBFBF");

            for (int r = 0; r < grid.Count; r++)
            {
                var row = grid[r];
                var kind = rowKinds[r];
                var rowRange = sheet.Range(r + 1, 1, r + 1, columnCount);

                rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rowRange.Style.Border.OutsideBorderColor = borderColor;
                rowRange.Style.Border.InsideBorderColor = borderColor;

                switch (kind)
                {
                    case RowKind.Title:
                        rowRange.Merge();
                        rowRange.Style.Fill.BackgroundColor = titleFill;
                        rowRange.Style.Font.FontColor = XLColor.White;
                        rowRange.Style.Font.Bold = true;
                        rowRange.Style.Font.FontSize = 14;
                        rowRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        sheet.Row(r + 1).Height = 24;
                        break;

                    case RowKind.Section:
                        rowRange.Merge();
                        rowRange.Style.Fill.BackgroundColor = sectionFill;
                        rowRange.Style.Font.FontColor = XLColor.White;
                        rowRange.Style.Font.Bold = true;
                        rowRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        break;

                    case RowKind.SubSection:
                        rowRange.Merge();
                        rowRange.Style.Fill.BackgroundColor = subSectionFill;
                        rowRange.Style.Font.Bold = true;
                        rowRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        break;

                    case RowKind.ColumnHeader:
                        rowRange.Style.Fill.BackgroundColor = columnHeaderFill;
                        rowRange.Style.Font.Bold = true;
                        rowRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        for (int c = 0; c < row.Count; c++)
                            if (row[c]?.IsMultiLine == true)
                                sheet.Cell(r + 1, c + 1).Style.Alignment.WrapText = true;
                        break;

                    case RowKind.Data:
                        for (int c = 0; c < row.Count; c++)
                        {
                            var cell = row[c];
                            if (cell == null) continue;
                            var xlCell = sheet.Cell(r + 1, c + 1);
                            xlCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            xlCell.Style.Alignment.Horizontal = cell.IsNumeric
                                ? XLAlignmentHorizontalValues.Right
                                : XLAlignmentHorizontalValues.Left;
                            if (cell.IsBold && !cell.IsNumeric)
                            {
                                xlCell.Style.Font.Bold = true;
                                xlCell.Style.Fill.BackgroundColor = labelFill;
                            }
                        }
                        break;
                }
            }

            sheet.Columns().AdjustToContents();
            foreach (var col in sheet.ColumnsUsed())
            {
                if (col.Width > 30) col.Width = 30;
                if (col.Width < 8) col.Width = 8;
            }
        }

        static void AddSingleImageWatermark(string inputExcel, string outputExcel, string imagePath)
        {
            using var workbook = new XLWorkbook(inputExcel);

            foreach (var sheet in workbook.Worksheets)
            {
                var picture = sheet.AddPicture(imagePath);
                picture.Placement = XLPicturePlacement.FreeFloating;
                picture.Scale(3.0);

                int sheetWidthPx = Math.Max(1, sheet.LastColumnUsed()?.ColumnNumber() ?? 1) * 64;
                int sheetHeightPx = Math.Max(1, sheet.LastRowUsed()?.RowNumber() ?? 1) * 20;

                picture.MoveTo(
                    Math.Max(0, sheetWidthPx / 2 - picture.Width / 2),
                    Math.Max(0, sheetHeightPx / 2 - picture.Height / 2));
            }

            foreach (var sheet in workbook.Worksheets)
            {
                sheet.Cell(1, 1).Value = "";
                sheet.Protect();
            }

            if (File.Exists(outputExcel)) File.Delete(outputExcel);
            workbook.SaveAs(outputExcel);
        }
    }
}
