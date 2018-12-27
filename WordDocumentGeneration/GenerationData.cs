namespace WordDocumentGeneration
{
    public class GenerationData
    {
        public DocumentProperties DocumentProperties { get; set; }
        public TitleArea TitleArea { get; set; }
    }

    public class TitleArea
    {
        public string Title { get; set; }
        public string Name { get; set; }
        public string Date { get; set; }
    }

    public class DocumentProperties
    {
        public string Creator { get; set; }
        public string Title { get; set; }
        public string Subject { get; set; }
        public string Category { get; set; }
        public string Keywords { get; set; }
        public string Description { get; set; }
    }
}
