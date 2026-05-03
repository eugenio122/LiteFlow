namespace LiteFlow.Models
{
    public class EvidenceData
    {
        public string StepId { get; set; } = Guid.NewGuid().ToString("N");
        public string ImageDataBase64 { get; set; } = "";
        public string Note { get; set; } = "";
        public bool TextBelowImage { get; set; } = false; 
        public bool IsEvidenceOnly { get; set; } = false;
    }
}