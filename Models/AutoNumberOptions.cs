namespace BoekSolutions.SheetSetEditor.Models
{
    public class AutoNumberOptions
    {
        public int StartNumber { get; set; }
        public int Increment { get; set; } = 1;
        public string Prefix { get; set; } = "";
        public string Suffix { get; set; } = "";
        public bool ApplyPageNumber { get; set; } = true;

    }
}
