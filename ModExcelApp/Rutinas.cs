using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace ModExcelApp
{
    public static class Rutinas
    {

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
            CopyCellFormatting(hojaMain.Cells[3, 30], otherSheet.Cells[1, 14]);
            CopyCellFormatting(hojaMain.Cells[3, 30], otherSheet.Cells[1, 15]);
            CopyCellFormatting(hojaMain.Cells[3, 30], otherSheet.Cells[1, 16]);

            //hojaMain.Cells[3, 30].Copy(otherSheet.Cells[1, 14], ExcelCopyOption.CopyFormats);
            //hojaMain.Cells[3, 30].Copy(otherSheet.Cells[1, 15], ExcelCopyOption.CopyFormats);
            //hojaMain.Cells[3, 30].Copy(otherSheet.Cells[1, 16], ExcelCopyOption.CopyFormats);

            otherSheet.Cells[1, 14].Value = "COSTO";
            otherSheet.Cells[1, 15].Value = "SUGERIDO";
        }




        private static void CopyCellFormatting(ExcelRange sourceCell, ExcelRange targetCell)
        {
            targetCell.Style.Numberformat.Format = sourceCell.Style.Numberformat.Format;
            targetCell.Style.Font.Name = sourceCell.Style.Font.Name;
            targetCell.Style.Font.Size = sourceCell.Style.Font.Size;
            targetCell.Style.Font.Bold = sourceCell.Style.Font.Bold;
            targetCell.Style.Font.Italic = sourceCell.Style.Font.Italic;
            targetCell.Style.Font.UnderLine = sourceCell.Style.Font.UnderLine;
            targetCell.Style.Fill.PatternType = sourceCell.Style.Fill.PatternType;

            Color bgcolor = ColorTranslator.FromHtml(sourceCell.Style.Fill.BackgroundColor.Rgb);
            targetCell.Style.Fill.BackgroundColor.SetColor(bgcolor);

            targetCell.Style.Border.Left.Style = sourceCell.Style.Border.Left.Style;
            targetCell.Style.Border.Right.Style = sourceCell.Style.Border.Right.Style;
            targetCell.Style.Border.Top.Style = sourceCell.Style.Border.Top.Style;
            targetCell.Style.Border.Bottom.Style = sourceCell.Style.Border.Bottom.Style;
            targetCell.Style.HorizontalAlignment = sourceCell.Style.HorizontalAlignment;
            targetCell.Style.VerticalAlignment = sourceCell.Style.VerticalAlignment;
        }
    }





}
