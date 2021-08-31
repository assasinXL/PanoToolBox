using System.Windows;
using System.IO;
using System.Xml.Linq;

namespace Pano_ToolBox
{
    /// <summary>
    /// Interaction logic for MapWindow.xaml
    /// </summary>
    public partial class MapWindow : Window
    {
        public MapWindow()
        {
            InitializeComponent();
        }

        private void SaveXML()
        {
            string projname = this.Name.Content.ToString();
            var file = $"{Directory.GetCurrentDirectory()}/output/{projname}/resources/showmap.xml";
            XElement root = XElement.Load(file);
            foreach (var layer in root.Elements("layer"))
            {
                if (layer.Attribute("name").Value == "mapcontainer")
                {
                    layer.SetAttributeValue("scale.normal", this.mapsize_pc_value.Text);
                    layer.SetAttributeValue("scale.mobile", this.mapsize_mb_value.Text);
                }
            }
            foreach (var style in root.Elements("style"))
            {
                if (style.Attribute("name").Value == "spot")
                {
                    style.SetAttributeValue("scale.normal", this.spotsize_pc_value.Text);
                    style.SetAttributeValue("scale.mobile", this.spotsize_mb_value.Text);
                }
            }
            root.Save(file);
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.browser.Dispose();
            this.Close();
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            this.SaveXML();
            this.browser.Refresh();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            this.SaveXML();
            this.browser.Dispose();
            this.Close();
        }
    }
}
