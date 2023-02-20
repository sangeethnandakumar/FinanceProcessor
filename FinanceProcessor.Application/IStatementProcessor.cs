using FinanceProcessor.Core.Stetement;
using System.Collections.Generic;

namespace FinanceProcessor.Application
{
    /// <summary>
    /// Handles statement related operations
    /// </summary>
    public interface IStatementProcessor
    {
        List<FinancialStatement> MakeStatements(string detailExcelFile, string groupExcelFile);
        void PrintSinglePageStatements();
        void PrintMultiPageStatements();
        void MergePDFPages();
    }
}
