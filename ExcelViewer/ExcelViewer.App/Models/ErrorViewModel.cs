using System;

namespace ExcelViewer.App.Models
{
    public class ErrorViewModel
    {
        public string RequestId { get; set; }

        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
    }
    public class Sheet
    {
        public string SheetName { get; set; }
        public string Path { get; set; }
    }
}
