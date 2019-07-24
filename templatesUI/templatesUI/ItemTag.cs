using System.IO;

namespace templatesUI
{
    public class ItemTag
    {
        public FileSystemInfo FileInfo { get; set; }
        public MicrosoftItem MicrosoftItem { get; set; }
        public bool Selected { get; set; }
        public string Color { get; set; }
    }
}