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
        public string DetailExcelLoc { get; set; }
        public string GroupExcelLoc { get; set; }
        public string SinglePDFLoc { get; set; }
        public string MultiPDFLoc { get; set; }
        public DataTable DetailTable { get; set; }
        public DataTable GroupTable { get; set; }
    }

    public class SecondPage
    {
        public bool NotifyWhenCompleted { get; set; } = true;
        public bool OpenWhenCompleted { get; set; } = true;
        public int CPUCores { get; set; } = Environment.ProcessorCount;
    }
}