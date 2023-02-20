using System.Data;

namespace FinanceProcessor.Application.Handlers
{
    /// <summary>
    /// Handles Excel related operations
    /// </summary>
    public interface IExcelHandler
    {
        DataTable ParseExcelDataIntoDataTable(string excelFile, int sheetIndex);
    }
}
