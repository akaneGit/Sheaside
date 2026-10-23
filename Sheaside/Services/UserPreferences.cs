using System;
using System.IO;
using System.Xml;

namespace Sheaside.Services
{
    internal enum PaneTheme { System, Light, Dark }
    internal sealed class UserPreferences
    {
        private readonly string path;
        public bool ShowVeryHidden { get; private set; }
        public PaneTheme Theme { get; private set; }
        public bool StartupVisible { get; private set; } = true;
        public bool DockLeft { get; private set; }
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
                PaneTheme theme;
                if (Enum.TryParse(document.DocumentElement?.GetAttribute("theme"), out theme) && Enum.IsDefined(typeof(PaneTheme), theme)) Theme = theme;
                StartupVisible = document.DocumentElement?.GetAttribute("startupVisible") != "false";
                DockLeft = document.DocumentElement?.GetAttribute("dockLeft") == "true";
            }
            catch (Exception e) when (e is IOException || e is UnauthorizedAccessException || e is XmlException)
            { LoadError = e.Message; }
        }
        public void SetShowVeryHidden(bool value)
        { SetOptions(value, Theme, StartupVisible, DockLeft); }
        public void SetOptions(bool value, PaneTheme theme, bool startupVisible, bool dockLeft)
        {
            if (!Enum.IsDefined(typeof(PaneTheme), theme)) throw new InvalidOperationException("未対応のテーマです。");
            if (value == ShowVeryHidden && theme == Theme && startupVisible == StartupVisible && dockLeft == DockLeft) return;
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
                        writer.WriteAttributeString("theme", theme.ToString());
                        writer.WriteAttributeString("startupVisible", startupVisible ? "true" : "false");
                        writer.WriteAttributeString("dockLeft", dockLeft ? "true" : "false");
                        writer.WriteEndElement();
                    }
                    if (File.Exists(path)) File.Replace(temporary, path, null);
                    else File.Move(temporary, path);
                }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
            }
            ShowVeryHidden = value;
            Theme = theme; StartupVisible = startupVisible; DockLeft = dockLeft;
            Changed?.Invoke(this, EventArgs.Empty);
        }
    }
}
