using FinanceProcessor.Core.Stetement;
using System.Collections.Generic;

namespace FinanceProcessor.Core.Statement
{
    public class PageResources
    {
        public Dictionary<string, string> Images { get; set; }
        public List<string> Styles { get; set; } = new List<string>();
        public FinancialStatement FinancialStatement { get; set; }
    }
}
