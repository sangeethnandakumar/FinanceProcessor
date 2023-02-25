using System.Collections.Generic;

namespace FinanceProcessor.Core.Stetement
{
    public class FinancialStatement
    {
        public string ReceiptID { get; set; }
        public string FullName { get; set; }
        public string AddressLine1 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZIPCode { get; set; }
        public string IMBarcode { get; set; }
        public string QRContent { get; set; }
        public string TraySort { get; set; }
        public int Pages { get; set; }
        public decimal Total { get; set; }
        public int FicialYear { get; set; }
        public List<Payment> Payments { get; set; }
    }
}
