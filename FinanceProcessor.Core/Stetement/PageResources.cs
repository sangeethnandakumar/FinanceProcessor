using System.Collections.Generic;

namespace FinanceProcessor.Core.Stetement
{
    public class PageResources
    {
        public Dictionary<string, string> Images { get; set; }
        public List<string> Styles { get; set; } = new List<string>();
        public FinancialStatement FinancialStatement { get; set; }
    }
}
