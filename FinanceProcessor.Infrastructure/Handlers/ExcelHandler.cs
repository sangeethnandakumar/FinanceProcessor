using FinanceProcessor.Application.Handlers;
using NPOI.XSSF.UserModel;
using System.Data;
using System.IO;

namespace FinanceProcessor.Infrastructure.Handlers
{
    public class ExcelHandler : IExcelHandler
    {
        public DataTable ParseExcelDataIntoDataTable(string excelFile, int sheetIndex)
        {
            var dtTable = new DataTable();
            using (var rstream = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
            {
                rstream.Position = 0;
                var xssWorkbook = new XSSFWorkbook(rstream);
                var sheet = xssWorkbook.GetSheetAt(sheetIndex);

                // Create DataTable and add columns
                var headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int j = 0; j < cellCount; j++)
                {
                    var cell = headerRow.GetCell(j);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    dtTable.Columns.Add(cell.ToString());
                }

                // Add data rows to DataTable
                for (var i = 1; i < sheet.PhysicalNumberOfRows; i++)
                {
                    var dataRow = dtTable.NewRow();
                    for (int j = 0; j < cellCount; j++)
                    {
                        var cell = sheet.GetRow(i).GetCell(j);
                        if (cell == null) continue;
                        dataRow[j] = cell.ToString();
                    }
                    dtTable.Rows.Add(dataRow);
                }
                rstream.Close();
            }
            return dtTable;
        }
    }
}
