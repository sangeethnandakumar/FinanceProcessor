using FinanceProcessor.Core.Stetement;
using System.Collections.Generic;
using System.Data;

namespace FinanceProcessor.Application.Handlers
{
    /// <summary>
    /// Handles PDF related operations
    /// </summary>
    public interface IPDFHandler
    {
        void MakePDF(string fileName, string html);
    }
}
