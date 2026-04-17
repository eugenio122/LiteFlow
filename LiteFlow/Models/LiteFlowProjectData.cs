using DocumentFormat.OpenXml.Drawing.Diagrams;
using System.Collections.Generic;

namespace LiteFlow.Models
{
    public class LiteFlowProjectData
    {
        public string TemplatePath { get; set; } = "";
        public string FilePrefix { get; set; } = "";
        public string FileName { get; set; } = "";
        public string TestCaseName { get; set; } = "";
        public string QAName { get; set; } = "";
        public string TestDate { get; set; } = "";
        public string Comments { get; set; } = "";

        public LayoutMode ReportLayout { get; set; } = LayoutMode.Padrao;
        public int MobileColumns { get; set; } = 2;

        public List<EvidenceData> Steps { get; set; } = new();
    }

}