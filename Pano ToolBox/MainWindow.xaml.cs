using System.Collections.Generic;
using System.Windows;
using System.IO;

namespace Pano_ToolBox
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void xlsxlist_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded)
            {
                var lsFile = new GetFileList();
                var xlsxfiles = lsFile.GetList();
                if(xlsxfiles.Count == 0)
                {
                    this.msgbox.AppendText("no xlsx files.");
                    return;
                }
                foreach(var file in xlsxfiles)
                {
                    this.xlsxlist.Items.Add(file);
                }
            }
        }
    }

    /// <summary>
    /// Glob files
    /// </summary>
    public partial class GetFileList
    {
        public List<string> GetList()
        {
            var xlsxfiles = new DirectoryInfo($"{Directory.GetCurrentDirectory()}/xlsx").GetFiles($"*.xlsx");
            var filelist = new List<string>();
            foreach (var file in xlsxfiles)
            {
                filelist.Add(new FileInfo(file.FullName).Name);
            }
            return filelist;
        }
    }
}
