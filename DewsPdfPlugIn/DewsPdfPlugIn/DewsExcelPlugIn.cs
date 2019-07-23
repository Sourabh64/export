using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DewsPdfPlugIn
{
    public class DewsExcelPlugIn:IDewsExport
    {
        public void Export(IDictionary<string, string> ProjectDetails, IDictionary<string, string> Metrics, IDictionary<string, Dictionary<string, string>> ProjectMetricValues, string outpath)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook wb = oXL.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (Worksheet)oXL.ActiveSheet;
            Range tRange = ws.get_Range("H5");
            tRange.Interior.Color = XlRgbColor.rgbGreen;
            
            int row = 4;
            int col = 5;
            foreach (var a in ProjectDetails)
            {
                ws.Cells[row, col] = a.Key;
                ws.Cells[row, col].Interior.Color = XlRgbColor.rgbGrey;
                ws.Cells[row, col].Borders.Weight = Excel.XlBorderWeight.xlMedium;
                ws.Cells[row, col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                col++;
            }
            ws.Rows[6].Orientation = Excel.XlOrientation.xlUpward;
            row++;
            col = 5;
            foreach (var b in ProjectDetails)
            {
                ws.Cells[row, col] = b.Value;
                ws.Cells[row, col].Interior.Color = XlRgbColor.rgbGrey;
                ws.Cells[row, col].Borders.Weight = Excel.XlBorderWeight.xlMedium;
                ws.Cells[row, col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                col++;
            }
            row++;
            col = 2;
            foreach (var a in Metrics)
            {
                ws.Cells[row, col] = a.Key;
                ws.Cells[row, col].Interior.Color = XlRgbColor.rgbDarkGrey;
                ws.Cells[row, col].Borders.Weight = Excel.XlBorderWeight.xlMedium;
                ws.Cells[row, col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ws.Cells[row, col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                if (col == 6 || col == 11 || col == 12)
                {
                    ws.Columns[col].ColumnWidth = 20;                   
                }
                else
                {
                    ws.Columns[col].ColumnWidth = 15;
                }
                col++;
            }
            row++;
            col = 2;
            foreach (var b in Metrics)
            {
                ws.Cells[row, col] = b.Value;
                ws.Cells[row, col].Interior.Color = XlRgbColor.rgbLightGrey;
                ws.Cells[row, col].Borders.Weight = Excel.XlBorderWeight.xlMedium;
                ws.Cells[row, col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                col++;
            }
            row = 8;
            col = 1;
            foreach (var a in ProjectMetricValues)
            {
                ws.Cells[row, col] = a.Key;
                ws.Cells[row, col].Interior.Color = XlRgbColor.rgbDarkGrey;
                ws.Cells[row, col].Borders.Weight = Excel.XlBorderWeight.xlMedium;
                ws.Cells[row, col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                row++;
            }
            row = 8;
            col = 2;
            foreach (var b in ProjectMetricValues.Keys)
            {
                var val = ProjectMetricValues[b];
                foreach (var c in val.Keys)
                {
                    ws.Cells[row, col] = val[c];
                    ws.Cells[row, col].Borders.Weight = Excel.XlBorderWeight.xlMedium;
                    ws.Cells[row, col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    col++;
                }
                row++;
                col = 2;
            }
            ws.Cells[7, 1] = "Goal";
            ws.Cells[7, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ws.Cells[7, 1].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            ws.Cells[6, 1].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            ws.Cells[7, 1].Interior.Color = XlRgbColor.rgbDarkGrey;
            ws.Cells[6, 1].Interior.Color = XlRgbColor.rgbDarkGrey;
            ws.Cells[5, 8].Interior.Color = XlRgbColor.rgbGreen;
            ws.Columns[1].ColumnWidth = 15;
            Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)ws.Cells[1, 1];
            float Left = (float)((double)oRange.Left);
            float Top = (float)((double)oRange.Top);
            const float ImageSize = 140;
            const float ImageSize1 = 40;
            ws.Cells[1, 6] = "DEWS Report";
            ws.Cells[1, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ws.Cells[1, 6].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            ws.Cells[1, 6].Font.Size = 20;
            ws.Range["F1", "F3"].Merge();
            ws.Shapes.AddPicture("D:\\siemens.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, ImageSize, ImageSize1);
            ws.Cells[1, 11] = DateTime.Now.ToOADate().ToString();
            ws.Cells[1, 11].NumberFormat = "mm/dd/yyyy HH:mm:ss";

            //ws.UsedRange.Columns.AutoFit();
            oXL.Visible = true;
            wb.SaveCopyAs(outpath);
        }
    }
}
