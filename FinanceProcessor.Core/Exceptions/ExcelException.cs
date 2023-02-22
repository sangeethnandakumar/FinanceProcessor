using System;
using System.Collections.Generic;
using System.Linq;

namespace FinanceProcessor.Core.Exceptions
{
    public class GenericException : Exception
    {
        public string Heading { get; set; }
        public string SubHeading { get; set; }
        public List<string> Options { get; set; } = new List<string>();

        public GenericException(string heading, string subHeading, string[] options)
        {
            Heading = heading;
            SubHeading = subHeading;
            Options = options.ToList();
        }
    }

    public class ExcelException : GenericException
    {
        public ExcelException(string heading, string subHeading, params string[] options) : base(heading, subHeading, options) 
        {
        }
    }
}
