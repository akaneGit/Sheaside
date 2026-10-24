using System;
using System.IO;
using System.Xml;

namespace Sheaside.Services
{
    internal sealed class UserPreferences
    {
        private readonly string path;
        public bool ShowVeryHidden { get; private set; }
        public bool ShowHidden { get; private set; }
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
                ShowHidden = document.DocumentElement?.GetAttribute("showHidden") == "true";
            }
            catch (Exception e) when (e is IOException || e is UnauthorizedAccessException || e is XmlException)
            { LoadError = e.Message; }
        }
        public void SetShowVeryHidden(bool value)
        { Save(value, ShowHidden); }
        public void SetShowHidden(bool value) { Save(ShowVeryHidden, value); }
        private void Save(bool value, bool showHidden)
        {
            if (value == ShowVeryHidden && showHidden == ShowHidden) return;
            if (path != null)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
                string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    using (var writer = XmlWriter.Create(temporary, new XmlWriterSettings { Indent = true }))
                    {
                        writer.WriteStartElement("Sheaside");
                        writer.WriteAttributeString("showHidden", showHidden ? "true" : "false");
                        writer.WriteAttributeString("showVeryHidden", value ? "true" : "false");
                        writer.WriteEndElement();
                    }
                    if (File.Exists(path)) File.Replace(temporary, path, null);
                    else File.Move(temporary, path);
                }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
            }
            ShowVeryHidden = value;
            ShowHidden = showHidden;
            Changed?.Invoke(this, EventArgs.Empty);
        }
    }
}
