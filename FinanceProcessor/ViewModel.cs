using System.Data;

namespace FinanceProcessor
{
    public class ViewModel
    {
        public FirstPage FirstPage { get; set; } = new FirstPage();
        public SecondPage SecondPage { get; set; } = new SecondPage();
    }

    public class FirstPage
    {
        public string DetailExcelLoc { get; set; } = string.Empty;
        public string GroupExcelLoc { get; set; } = string.Empty;
        public string SinglePDFLoc { get; set; } = string.Empty;
        public string MultiPDFLoc { get; set; } = string.Empty;
        public DataTable DetailTable { get; set; } = new DataTable();
        public DataTable GroupTable { get; set; } = new DataTable();
    }

    public class SecondPage
    {
        public bool NotifyWhenCompleted { get; set; } = true;
        public bool OpenWhenCompleted { get; set; } = true;
        public int CPUCores { get; set; } = Environment.ProcessorCount;
    }
}