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
            var lsFiles = new GetFileList();
            var xlsxfiles = lsFiles.GetList("txt");
            foreach(var file in xlsxfiles)
            {
                MessageBox.Show(file);
            }
        }
    }

    /// <summary>
    /// Glob files
    /// </summary>
    public partial class GetFileList
    {
        private string cwd = Directory.GetCurrentDirectory();
        public List<string> GetList(string ext)
        {
            var xlsxroot = new DirectoryInfo($"{this.cwd}/xlsx");
            var xlsxfiles = xlsxroot.GetFiles($"*.{ext}");
            var filelist = new List<string>();
            foreach (var file in xlsxfiles)
            {
                filelist.Add(file.ToString());
            }
            return filelist;
        }
    }
}
