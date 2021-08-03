using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GoogleSheetsToDocs.Models
{
    public class SheetVM
    {
        public SheetVM()
        {
            ClassificationCodes = new();
            Claims = new();
            ApplicationEvents = new();
        }

        public string SheetId { get; set; }
        public string DocId { get; set; }
        public CoverPage CoverPage { get; set; }
        public string Abstract { get; set; }
        public List<ClassificationCode> ClassificationCodes { get; set; }
        public List<Claim> Claims { get; set; }
        public List<ApplicationEvent> ApplicationEvents { get; set; }

        public string Error { get; set; }

    }


    public class CoverPage
    {
        public string Title { get; set; }
        public string PreparedBy { get; set; }
        public string Inventor { get; set; }
        public string PatentNo { get; set; }
    }
    public class ClassificationCode
    {
        public string USC_Code { get; set; }
        public string Definition { get; set; }
    }
    public class Claim
    {
        public IndentLevel Level { get; set; }
        public string Content { get; set; }
        public List<Claim> Children { get; set; } = new List<Claim>();

        public enum IndentLevel
        {
            L1, L2, L3
        };

    }

    public class ApplicationEvent
    {
        public DateTime Date { get; set; }
        public string Status { get; set; }
        public string Hyperlink { get; set; }
    }

}
