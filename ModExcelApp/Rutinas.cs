using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Linq;
using OfficeOpenXml.Drawing;
using System.IO;

namespace ModExcelApp
{
    public static class Rutinas
    {

        public static System.Drawing.Color title_bg = System.Drawing.Color.LightSteelBlue;
        public static System.Drawing.Color regular_bg = System.Drawing.Color.LightGoldenrodYellow;


        public static void DeleteAllExceptDespacho(ExcelWorkbook workbook)
        {
            // Disable display alerts
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Collect all sheets except the one named "Despacho"
            var sheetsToDelete = workbook.Worksheets
                .Where(sheet => sheet.Name != "Despacho")
                .ToList();

            // Delete collected sheets
            foreach (var sheet in sheetsToDelete)
            {
                workbook.Worksheets.Delete(sheet);
            }
        }

        public static ExcelWorksheet CreateSheetIfNotExist(ExcelWorkbook workbook, string sheetName)
        {
            // Check if the sheet already exists
            var existingSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (existingSheet != null)
            {
                return existingSheet;
            }

            // If the sheet does not exist, create it
            var newSheet = workbook.Worksheets.Add(sheetName);
            return newSheet;
        }


        public static void SetupColTitles(ExcelWorksheet hojaMain, ExcelWorksheet otherSheet)
        {
            if (hojaMain == null || otherSheet == null)
            {
                throw new ArgumentNullException("One or more invalid worksheet references provided.");
            }

            otherSheet.Row(1).Height = 59;

            // Set the width of columns in otherSheet
            otherSheet.Column(2).Width = 24; // foto1
            otherSheet.Column(3).Width = 24;
            otherSheet.Column(4).Width = 24;
            otherSheet.Column(5).Width = 24;
            otherSheet.Column(6).Width = 24; // col width foto5

            otherSheet.Column(10).Width = 20;
            otherSheet.Column(11).Width = 20;
            otherSheet.Column(12).Width = 20;
            otherSheet.Column(13).Width = 20;
            otherSheet.Column(14).Width = 20;
            otherSheet.Column(15).Width = 20;

            // Copy the content from specified cells in hojaMain to otherSheet
            hojaMain.Cells[3, 1].Copy(otherSheet.Cells[1, 1]); // marca
            hojaMain.Cells[3, 3].Copy(otherSheet.Cells[1, 2]); // foto1
            hojaMain.Cells[3, 4].Copy(otherSheet.Cells[1, 3]);
            hojaMain.Cells[3, 5].Copy(otherSheet.Cells[1, 4]);
            hojaMain.Cells[3, 6].Copy(otherSheet.Cells[1, 5]);
            hojaMain.Cells[3, 7].Copy(otherSheet.Cells[1, 6]); // foto5

            hojaMain.Cells[3, 21].Copy(otherSheet.Cells[1, 7]); // CTNS
            hojaMain.Cells[3, 22].Copy(otherSheet.Cells[1, 8]); // Cantidad
            otherSheet.Cells[1, 8, 1, 9].Merge = true; // Merge H1:I1
            otherSheet.Cells[1, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            hojaMain.Cells[3, 24].Copy(otherSheet.Cells[1, 10]); // Precio RMB

            hojaMain.Cells[3, 30].Copy(otherSheet.Cells[1, 11]); // Total
            hojaMain.Cells[3, 35].Copy(otherSheet.Cells[1, 12]); // CBM Caja
            hojaMain.Cells[3, 36].Copy(otherSheet.Cells[1, 13]); // CBM total

            // Copy only format
            CopyCellFormatting(hojaMain.Cells[3, 30], otherSheet.Cells[1, 14], title_bg);
            CopyCellFormatting(hojaMain.Cells[3, 30], otherSheet.Cells[1, 15], title_bg);
            CopyCellFormatting(hojaMain.Cells[3, 30], otherSheet.Cells[1, 16], title_bg);


            otherSheet.Cells[1, 14].Value = "COSTO";
            otherSheet.Cells[1, 15].Value = "SUGERIDO";
        }


        public static void CopyCellFormatting(ExcelRange sourceCell, ExcelRange targetCell, System.Drawing.Color bg_color)
        {
            targetCell.Style.Numberformat.Format = sourceCell.Style.Numberformat.Format;
            targetCell.Style.Font.Name = sourceCell.Style.Font.Name;
            targetCell.Style.Font.Size = sourceCell.Style.Font.Size;
            targetCell.Style.Font.Bold = sourceCell.Style.Font.Bold;
            targetCell.Style.Font.Italic = sourceCell.Style.Font.Italic;
            targetCell.Style.Font.UnderLine = sourceCell.Style.Font.UnderLine;
            targetCell.Style.Fill.PatternType = sourceCell.Style.Fill.PatternType;
            targetCell.Style.Fill.SetBackground(bg_color);
            targetCell.Style.Border.Left.Style = sourceCell.Style.Border.Left.Style;
            targetCell.Style.Border.Right.Style = sourceCell.Style.Border.Right.Style;
            targetCell.Style.Border.Top.Style = sourceCell.Style.Border.Top.Style;
            targetCell.Style.Border.Bottom.Style = sourceCell.Style.Border.Bottom.Style;
            targetCell.Style.HorizontalAlignment = sourceCell.Style.HorizontalAlignment;
            targetCell.Style.VerticalAlignment = sourceCell.Style.VerticalAlignment;
        }



        public static void CopiarRecord(ExcelWorksheet hoja_main, ExcelWorksheet dst_sheet, int row_orig, int row_dst)
        {
            // Set row height
            dst_sheet.Row(row_dst).Height = 116;

            // Copy marca
            CopyCell(hoja_main, dst_sheet, row_orig, 1, row_dst, 1);

            // Copy fotos
            CopyImage(hoja_main, row_orig, 3, dst_sheet, row_dst, 2);
            CopyImage(hoja_main, row_orig, 4, dst_sheet, row_dst, 3);
            CopyImage(hoja_main, row_orig, 5, dst_sheet, row_dst, 4);
            CopyImage(hoja_main, row_orig, 6, dst_sheet, row_dst, 5);
            CopyImage(hoja_main, row_orig, 7, dst_sheet, row_dst, 6);

            //CopyCell(hoja_main, dst_sheet, row_orig, 3, row_dst, 2); // foto1
            //CopyCell(hoja_main, dst_sheet, row_orig, 4, row_dst, 3);
            //CopyCell(hoja_main, dst_sheet, row_orig, 5, row_dst, 4);
            //CopyCell(hoja_main, dst_sheet, row_orig, 6, row_dst, 5);
            //CopyCell(hoja_main, dst_sheet, row_orig, 7, row_dst, 6); // foto5

            // Copy CTNS
            CopyCell(hoja_main, dst_sheet, row_orig, 21, row_dst, 7);

            // Copy Cantidad 1 (formula)
            CopyCell(hoja_main, dst_sheet, row_orig, 22, row_dst, 8);
            dst_sheet.Cells[row_dst, 8].Value = hoja_main.Cells[row_orig, 22].Value;

            // Copy Cantidad 2
            CopyCell(hoja_main, dst_sheet, row_orig, 23, row_dst, 9);

            // Copy Precio RMB
            CopyCell(hoja_main, dst_sheet, row_orig, 24, row_dst, 10);

            // Copy Total RMB (formula)
            CopyCell(hoja_main, dst_sheet, row_orig, 30, row_dst, 11);
            dst_sheet.Cells[row_dst, 11].Value = hoja_main.Cells[row_orig, 30].Value;

            // Copy CBM Caja
            CopyCell(hoja_main, dst_sheet, row_orig, 35, row_dst, 12);

            // Copy CBM total (formula)
            CopyCell(hoja_main, dst_sheet, row_orig, 36, row_dst, 13);
            dst_sheet.Cells[row_dst, 13].Value = hoja_main.Cells[row_orig, 36].Value;
        }



        public static void CopyCell(ExcelWorksheet srcSheet, ExcelWorksheet dstSheet, int srcRow, int srcCol, int dstRow, int dstCol)
        {
            var srcCell = srcSheet.Cells[srcRow, srcCol];
            var dstCell = dstSheet.Cells[dstRow, dstCol];

            dstCell.Value = srcCell.Value;
            dstCell.StyleID = srcCell.StyleID;

            if (srcCell.IsRichText)
            {
                foreach (var rt in srcCell.RichText)
                {
                    dstCell.RichText.Add(rt.Text);
                    dstCell.RichText[dstCell.RichText.Count - 1].Bold = rt.Bold;
                    dstCell.RichText[dstCell.RichText.Count - 1].Italic = rt.Italic;
                    dstCell.RichText[dstCell.RichText.Count - 1].Color = rt.Color;
                    dstCell.RichText[dstCell.RichText.Count - 1].FontName = rt.FontName;
                    dstCell.RichText[dstCell.RichText.Count - 1].Size = rt.Size;
                    //dstCell.RichText[dstCell.RichText.Count - 1].Underline = rt.Underline;
                }
            }
        }

        public static int CountMergedCells(ExcelWorksheet ws, int row, int column)
        {
            var cell = ws.Cells[row, column];

            if (cell.Merge)
            {
                var mergedRange = ws.MergedCells.FirstOrDefault(range => ws.Cells[range].Any(c => c.Start.Row == row && c.Start.Column == column));
                if (mergedRange != null)
                {
                    var range = ws.Cells[mergedRange];
                    return range.Count();
                }
            }

            return 1;
        }
        
        public static void MergeCellsInColumnA(int startRow, int endRow, ExcelWorksheet ws)
        {
            // Check if the startRow is less than or equal to endRow
            if (startRow <= endRow)
            {
                // Specify the range to merge in column A (column 1) from startRow to endRow
                var startCell = ws.Cells[startRow, 1];
                var endCell = ws.Cells[endRow, 1];
                ws.Cells[startRow, 1, endRow, 1].Merge = true;
            }
            else
            {
                // Display a message if startRow is greater than endRow
                Console.WriteLine("The start row must be less than or equal to the end row.");
            }
        }


        public static bool IsCellEmpty(ExcelRange cell)
        {
            // Check if the cell's value is null or an empty string
            return cell.Value == null || string.IsNullOrEmpty(cell.Text);
        }


        public static void CopyImage(ExcelWorksheet srcSheet, int srcRow, int srcCol, ExcelWorksheet dstSheet, int dstRow, int dstCol)
        {
            foreach (var pic in srcSheet.Drawings.OfType<ExcelPicture>())
            {
                if (pic.From.Row + 1 == srcRow && pic.From.Column + 1 == srcCol)
                {
                    using (var imageStream = new MemoryStream())
                    {
                        pic.Image.Save(imageStream, pic.Image.RawFormat);
                        imageStream.Position = 0; // Reset the stream position to the beginning

                        var newPic = dstSheet.Drawings.AddPicture(pic.Name, imageStream);
                        newPic.SetPosition(dstRow - 1, pic.From.RowOff, dstCol - 1, pic.From.ColOff);
                        newPic.SetSize(pic.Width, pic.Height);
                    }
                    break; // Assuming only one picture per cell
                }
            }
        }


    }
}
