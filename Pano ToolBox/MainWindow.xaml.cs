using System.Collections.Generic;
using System.Windows;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Diagnostics;
using GlobExpressions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using HtmlAgilityPack;

namespace Pano_ToolBox
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private CoreProgram program = null;
        private List<string> xlsxfiles = null;
        public MainWindow()
        {
            this.program = new CoreProgram();
            this.xlsxfiles = new ExcelReader().GetList();
            InitializeComponent();
        }

        private void xlsxlist_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded)
            {
                if(this.xlsxfiles.Count == 0)
                {
                    this.msgbox.AppendText("没有检测到Excel文件。\n");
                    return;
                }
                foreach(var file in this.xlsxfiles)
                {
                    this.xlsxlist.Items.Add(new FileInfo(file).Name);
                }
            }
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in this.xlsxfiles)
            {
                // Read Excel
                var projname = this.program.ExcelToMemory(item);
                this.msgbox.AppendText($"---------- {projname} ----------\n");
                this.msgbox.AppendText($"Excel文件读取完成。\n");
                // Sort Pano
                if (this.program.SortPano() == 1)
                {
                    this.msgbox.AppendText($"检测到项目\"{projname}\"中存在src文件夹，将不会对此项目进行图片分拣而直接使用已有的全景图片资源进行生成任务，请自行确保src文件夹中的图片数据的有效性。\n");
                }
                else if (this.program.SortPano() == 2)
                {
                    this.msgbox.AppendText($"请检查data文件夹内容和时间线内容是否正确。");
                }
                else
                {
                    this.msgbox.AppendText($"全景图片分拣完成。\n");
                }
                // Make Pano
                this.program.MakePano();
                this.msgbox.AppendText("全景项目生成完成。\n");
                // Parse .html
                this.program.ParseHTML();
                this.msgbox.AppendText("tour.html解析完成，现已更名为index.html。\n");
            }
        }
    }

    /// <summary>
    /// Excel Reader
    /// </summary>
    public partial class ExcelReader
    {
        public List<string> GetList()
        {
            var xlsxpath = $"{Directory.GetCurrentDirectory()}/xlsx";
            if (!Directory.Exists(xlsxpath))
            {
                Directory.CreateDirectory(xlsxpath);
            }
            var xlsxfiles = new DirectoryInfo(xlsxpath).GetFiles($"*.xlsx");
            var filelist = new List<string>();
            foreach (var file in xlsxfiles)
            {
                filelist.Add(file.FullName);
            }
            return filelist;
        }
        public DataTable Read(string filename)
        {
            var data = new DataTable();
            var fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(0);
            int startRow = 0;
            if (sheet != null)
            {
                IRow firstRow = sheet.GetRow(0);
                int cellCount = firstRow.LastCellNum;
                for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                {
                    ICell cell = firstRow.GetCell(i);
                    if (cell != null)
                    {
                        string cellValue = cell.StringCellValue;
                        if (cellValue != null)
                        {
                            DataColumn column = new DataColumn(cellValue);
                            data.Columns.Add(column);
                        }
                    }
                }
                startRow = sheet.FirstRowNum + 1;
                int rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        if (row.GetCell(j) != null)
                            dataRow[j] = row.GetCell(j).ToString();
                    }
                    data.Rows.Add(dataRow);
                }
            }
            fs.Close();
            return data;
        }
    }
    /// <summary>
    /// Core Program
    /// </summary>
    public partial class CoreProgram
    {
        private List<int> pointnum = new List<int>();
        private List<string> pointname = new List<string>();
        private List<string> capturetime = new List<string>();
        private List<string> voicelist = new List<string>();
        private string projname = null;
        private int totalnum = 0;
        private string phonenum = null;
        private string voicemode = null;
        private string projdir = null;
        private bool Checkup()
        {
            return !Directory.Exists($"{this.projdir}/vtour");
        }
        private void CopyFile(string src, string dst)
        {
            if (!File.Exists(src))
                throw new FileNotFoundException($"找不到{src}");
            if (File.Exists(dst))
                File.Delete(dst);
            else
            {
                var dstpath = Directory.GetParent(dst).ToString();
                Directory.CreateDirectory(dstpath);
            }
            File.Copy(src, dst);
        }
        private void MoveDir(string src, string dst)
        {
            DirectoryInfo srcdir = new DirectoryInfo(src);
            var dstdir = $"{dst}/{srcdir.Name}";
            if (Directory.Exists(dstdir))
                Directory.Delete(dstdir, true);
            Directory.Move(src, dst);
        }
        public string ExcelToMemory(string file)
        {
            DataTable data = new ExcelReader().Read(file);
            this.pointnum.Clear();
            this.pointname.Clear();
            this.capturetime.Clear();
            this.voicelist.Clear();
            this.projname = data.Rows[0][0].ToString();
            this.totalnum = int.Parse(data.Rows[0][5].ToString());
            this.phonenum = data.Rows[2][5].ToString();
            this.voicemode = data.Rows[4][5].ToString();

            for (var i = 0; i < this.totalnum; i++)
            {
                this.pointnum.Add((int.Parse((string)data.Rows[i][1])));
                this.pointname.Add(data.Rows[i][2].ToString());
                this.capturetime.Add((data.Rows[i][3]).ToString().Replace(":", "_"));
                this.voicelist.Add(data.Rows[i][4].ToString());
            }

            this.projdir = $"{Directory.GetCurrentDirectory()}/output/{this.projname}";
            this.exportSpotInfoXML($"{Directory.GetCurrentDirectory()}/spot/spotinfo.xml");
            return this.projname;
        }
        public void exportSpotInfoXML(string file)
        {
            XDocument doc = new XDocument();
            XElement root = new XElement("project");
            root.SetAttributeValue("name", this.projname);
            XElement pointlist = new XElement("pointlist", new XAttribute("totalnum", this.totalnum));
            for (var index = 0; index < this.totalnum; index++)
            {
                XElement point = new XElement("point",
                    new XAttribute("num", this.pointnum[index]),
                    new XAttribute("name", this.pointname[index]));
                pointlist.Add(point);
            }
            root.Add(pointlist);
            root.Add(new XElement("spotlist"));
            doc.AddFirst(root);
            if (!Directory.Exists(new FileInfo(file).Directory.ToString()))
            {
                Directory.CreateDirectory(new FileInfo(file).Directory.ToString());
            }
            doc.Save(file);
        }
        public int SortPano()
        {
            if (Directory.Exists($"{this.projdir}/src"))
            {
                return 1;
            }
            for (var i = 0; i < this.totalnum; i++)
            {
                var srcroot = new DirectoryInfo($"{Directory.GetCurrentDirectory()}/data/{this.projname}");
                var srcpath = srcroot.GlobFiles($"*{this.capturetime[i]}/*.jpg").ToArray();
                if (srcpath.Count() == 0)
                    return 2;
                var dstpath = $"{Directory.GetCurrentDirectory()}/output/{this.projname}/src/{this.pointnum[i]}.jpg";
                this.CopyFile(srcpath[0].ToString(), dstpath);
            }
            return 0;
        }
        public void MakePano()
        {
            var krpanodir = $"{Directory.GetCurrentDirectory()}/krpano";
            var krpano = new Process()
            {
                StartInfo = new ProcessStartInfo($"{krpanodir}/krpanotools.exe",
                $"makepano \"{krpanodir}/templates/vtour-multires.config\" \"{this.projdir}/src/*.jpg\"")
        };
            krpano.Start();
            krpano.WaitForExit();
            krpano.Close();
            this.MoveDir($"{this.projdir}/src/vtour", $"{this.projdir}/vtour");
        }
        public void ParseHTML()
        {
            var file = $"{this.projdir}/vtour/tour.html";
            var htmlpath = (new DirectoryInfo(file)).Parent;
            HtmlDocument doc = new HtmlDocument();
            doc.Load(file);
            HtmlTextNode node = doc.DocumentNode.SelectSingleNode("//head//title//text()") as HtmlTextNode;
            node.Text = this.projname;
            File.Delete(file);
            doc.Save($"{htmlpath}/index.html");
        }
    }
}
