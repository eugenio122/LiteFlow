namespace LiteFlow.Models
{
    public class EvidenceData
    {
        public string ImageDataBase64 { get; set; } = "";
        public string Note { get; set; } = "";
        public bool TextBelowImage { get; set; } = false; 
        public bool IsEvidenceOnly { get; set; } = false;
    }
}