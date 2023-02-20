using FinanceProcessor.Application;
using FinanceProcessor.Application.Handlers;
using FinanceProcessor.Core.Stetement;
using System.Collections.Generic;

namespace FinanceProcessor.Infrastructure
{
    public class StatementProcessor : IStatementProcessor
    {
        private readonly IExcelHandler excelHandler;
        private readonly IPDFHandler pdfHandler;

        public StatementProcessor(IExcelHandler excelHandler, IPDFHandler pdfHandler)
        {
            this.excelHandler = excelHandler;
            this.pdfHandler = pdfHandler;
        }

        public List<FinancialStatement> MakeStatements(string detailExcelFile, string groupExcelFile)
        {
            var detailData = excelHandler.ParseExcelDataIntoDataTable(detailExcelFile, 0);
            var groupData = excelHandler.ParseExcelDataIntoDataTable(groupExcelFile, 0);
            return null;
        }

        public void MergePDFPages()
        {
            throw new System.NotImplementedException();
        }

        public void PrintMultiPageStatements()
        {
            throw new System.NotImplementedException();
        }

        public void PrintSinglePageStatements()
        {
            throw new System.NotImplementedException();
        }
    }
}
