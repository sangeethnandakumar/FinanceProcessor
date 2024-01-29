using System;

namespace FinanceProcessor.Core.Stetement
{
    public class Payment
    {
        public string ReceiptID { get; set; }
        public DateTime Date { get; set; }
        public string Check { get; set; }
        public decimal Amount { get; set; }
    }
}
