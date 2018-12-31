using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

class ExcelHandler
{
    public enum StyleNames
    {
        Header,
        BorderBox
    }
    private ExcelPackage package;
    public ExcelWorksheet worksheet;

    public ExcelHandler()
    {
        package = new ExcelPackage();
        // Creating Styles
    }

    public void CreateSheet(string sheetName)
    {
        worksheet = package.Workbook.Worksheets.Add(sheetName);
    }

    public void SetStyle(StyleNames currentStyle, int fromRow, int fromCol, int toRow, int toCol)
    {
        using (var range = worksheet.Cells[fromRow, fromCol, toRow, toCol])
        {
            switch (currentStyle)
            {
                case StyleNames.Header:
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    range.AutoFitColumns();
                    break;
                case StyleNames.BorderBox:
                    range.Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    range.Style.Border.DiagonalUp = true;
                    range.Style.Border.DiagonalDown = true;
                    break;
                default:
                    break;
            }
        }
    }

    public void SaveSheet(string fileName)
    {
        if(package != null)
        {
            FileInfo xlFile = new FileInfo(string.Format("{0}_{1}-{2}-{3}.{4}", fileName,
                                 DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, ".xlsx"));
            package.SaveAs(xlFile);
        }
    }
}
