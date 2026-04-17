using System.Drawing;
using System.Windows.Forms;

namespace LiteFlow.Models
{
    public class EvidenceItem
    {
        public Bitmap Image { get; set; } = null!;
        public string Note { get; set; } = "";
        public bool TextBelowImage { get; set; } = false;
        public bool IsEvidenceOnly { get; set; } = false;
        public PictureBox Thumbnail { get; set; } = null!;
        public string DiskPath { get; set; } = ""; // Caminho físico para Lazy Load na exportação
    }
}