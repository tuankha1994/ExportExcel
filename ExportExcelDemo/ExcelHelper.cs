using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcelDemo
{
    public class ExcelHelper
    {
        //export simple 1 sheet data horizontal
        public static bool ExportExcelFile<T>(string fileName, List<T> listData)
        {
            try
            {
                var file = new FileInfo(fileName);
                if (file.Exists)
                {
                    file.Delete();
                }
                using (ExcelPackage pck = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add("List " + typeof(T).Name);
                    var propsName = typeof(T).GetProperties().Select(x => x.Name).ToList();
                    AddTableHeader(worksheet, propsName);
                    AddTableContent(worksheet, propsName, listData);
                    pck.Save();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void AddTableHeader(ExcelWorksheet worksheet, List<string> listHeaders)
        {
            int startRowIndex = 1;
            for (int columnIndex = 1; columnIndex <= listHeaders.Count; columnIndex++)
            {
                var cell = worksheet.Cells[startRowIndex, columnIndex];
                var fill = cell.Style.Fill;
                fill.PatternType = ExcelFillStyle.Solid;
                fill.BackgroundColor.SetColor(Color.LightGray);
                cell.Value = listHeaders[columnIndex - 1];
            }
        }

        private static void AddTableContent<T>(ExcelWorksheet worksheet, List<string> listHeaders, List<T> listContent)
        {
            int startRowIndex = 2;
            for (int row = startRowIndex, index = 0; index < listContent.Count; row++, index++)
            {
                for (int columnIndex = 1, colIndex = 0; colIndex < typeof(T).GetProperties().Length; columnIndex++, colIndex++)
                {
                    var value = listContent[index].GetType().GetProperty(listHeaders[colIndex]).GetValue(listContent[index]);
                    var dataType = value.GetType();

                    if (dataType == typeof(DateTime))
                    {
                        worksheet.Cells[row, columnIndex].Value = ((DateTime)value).ToString("yyyy-MM-dd hh:mm tt");
                    }
                    else if (dataType == typeof(bool))
                    {
                        worksheet.Cells[row, columnIndex].Value = ((bool)value) ? "TRUE" : "FALSE";
                    }
                    else
                    {
                        worksheet.Cells[row, columnIndex].Value = value;
                    }
                }
            }
        }
    }
}
