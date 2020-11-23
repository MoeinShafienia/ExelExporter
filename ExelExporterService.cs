using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;
using System;

namespace ExelExporter
{
    public class ExelExporterService
    {
        public static void Export()
        {
            var fileInfo = new FileInfo("Inventory.xlsx");
            using (var package = new ExcelPackage(fileInfo))
            {
                //getting data
                var factory = new ProductFactory();
                var products =factory.GetListOfProdutsByGenfu();

                //Add a new worksheet to the empty workbook
                var worksheet = package.Workbook.Worksheets.Add("Inventory");

                // uncomment next line if you want Right to left support
                //worksheet.View.RightToLeft = true;

                var type = products[0].GetType();
                var propertyInfo = type.GetProperties();

                //Adding headers
                var column = 1;
                foreach(var property in propertyInfo)
                {
                    worksheet.Cells[1, column].Value = property.Name;
                    column++;
                }

                //Adding rows
                var row = 2;
                var dataColumn = 'A';
                foreach(var item in products)
                {
                    dataColumn = 'A';
                    foreach(var property in propertyInfo)
                    {
                        worksheet.Cells[$"{dataColumn}{row}"].Value = property.GetValue(item);
                        dataColumn++;
                    }
                    row++;
                }

                //Create an autofilter for the range
                worksheet.Cells.AutoFilter = true;

                //Ok now format the values;
                using (var range = worksheet.Cells[1, 1, 1, propertyInfo.Length]) 
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                    range.Style.Font.Color.SetColor(Color.White);
                }

                worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells

                package.Save();
            }
        }       
    }
}