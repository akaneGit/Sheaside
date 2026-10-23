using System;
using System.IO;
using System.Xml;

namespace Sheaside.Services
{
    internal sealed class UserPreferences
    {
        private readonly string path;
        public bool ShowVeryHidden { get; private set; }
        public string LoadError { get; private set; }
        public event EventHandler Changed;
        public UserPreferences(string path = null)
        {
            this.path = path;
            if (path == null || !File.Exists(path)) return;
            try
            {
                var document = new XmlDocument { XmlResolver = null };
                using (var reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit }))
                    document.Load(reader);
                ShowVeryHidden = document.DocumentElement?.GetAttribute("showVeryHidden") == "true";
            }
            catch (Exception e) when (e is IOException || e is UnauthorizedAccessException || e is XmlException)
            { LoadError = e.Message; }
        }
        public void SetShowVeryHidden(bool value)
        {
            if (value == ShowVeryHidden) return;
            if (path != null)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
                string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    using (var writer = XmlWriter.Create(temporary, new XmlWriterSettings { Indent = true }))
                    {
                        writer.WriteStartElement("Sheaside");
                        writer.WriteAttributeString("showVeryHidden", value ? "true" : "false");
                        writer.WriteEndElement();
                    }
                    if (File.Exists(path)) File.Replace(temporary, path, null);
                    else File.Move(temporary, path);
                }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
            }
            ShowVeryHidden = value;
            Changed?.Invoke(this, EventArgs.Empty);
        }
    }
}
