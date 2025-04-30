namespace NoteQuickFormatter.Models
{
    class Notebook
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public bool IsCurrentlyViewed { get; set; }
    }
}
