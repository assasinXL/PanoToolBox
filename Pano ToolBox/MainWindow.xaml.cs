using System.Collections.Generic;
using System.Windows;
using System.Data;
using System.IO;
using System.Xml.Linq;
using System.Diagnostics;
using System.Threading;
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
            InitializeComponent();
            if (!Directory.Exists($"{Directory.GetCurrentDirectory()}/input"))
            {
                Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/input");
            }
            if (!Directory.Exists($"{Directory.GetCurrentDirectory()}/xlsx"))
            {
                Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}/xlsx");
            }
        }
        private void xlsxlist_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded)
            {
                this.xlsxfiles = new ExcelReader().GetList();
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
        private void Run_Generate(object sender, RoutedEventArgs e)
        {
            var items = this.xlsxlist.SelectedItems;
            if (items.Count == 0)
            {
                this.msgbox.AppendText("请点选xlsx文件后再执行操作。\n");
                return;
            }
            foreach (var item in items)
            {
                // Read Excel
                var projname = this.program.ExcelToMemory(item.ToString());
                this.msgbox.AppendText($"---------- {projname} ----------\n");
                this.msgbox.AppendText($"Excel文件读取完成。\n");
                // Make Pano
                this.program.MakePano();
                this.msgbox.AppendText("全景项目生成完成。\n");
                // Import Template
                this.program.ImportTemplate(this.map.IsChecked.Value, this.voice.IsChecked.Value, this.contact.IsChecked.Value);
                this.msgbox.AppendText("模板文件导入完成。\n");
                // Parse .xml
                this.program.ParseTOUR(this.map.IsChecked.Value, this.voice.IsChecked.Value, this.contact.IsChecked.Value);
                this.msgbox.AppendText("配置文件修改完成。\n");
                if (this.contact.IsChecked.Value)
                {
                    this.program.ParseContact();
                    this.msgbox.AppendText("联系方式导入完成。\n");
                }
                // Parse .html
                this.program.ParseHTML();
                this.msgbox.AppendText("tour.html解析完成，现已更名为index.html。\n");
            }
        }
        private void Run_Refresh(object sender, RoutedEventArgs e)
        {
            this.xlsxlist.Items.Clear();
            this.xlsxlist_Loaded(sender, e);
        }
        private void Run_Cursor(object sender, RoutedEventArgs e)
        {
            var items = this.xlsxlist.SelectedItems;
            if (items.Count == 0)
            {
                this.msgbox.AppendText("请点选xlsx文件后再执行操作。\n");
                return;
            }
            foreach(var item in items)
            {
                var krpanodir = $"{Directory.GetCurrentDirectory()}/tool/krpano";
                var krpano = new Process()
                {
                    StartInfo = new ProcessStartInfo($"\"{krpanodir}/krpano Tools.exe\"",
                    $"\"{Directory.GetCurrentDirectory()}/output/{item.ToString().Split(".")[0]}/tour.xml\"")
                };
                krpano.Start();
                krpano.WaitForExit();
                krpano.Close();
            }

        }
        private void Run_Krpano(object sender, RoutedEventArgs e)
        {
            Thread ts_thread;
            ts_thread = new Thread(RunKrpano);
            ts_thread.IsBackground = true;
            ts_thread.Start();
        }
        private void Run_Server(object sender, RoutedEventArgs e)
        {
            Thread ts_thread;
            ts_thread = new Thread(RunServer);
            ts_thread.IsBackground = true;
            ts_thread.Start();
        }
        static void RunKrpano()
        {
            var ts = new Process() { StartInfo = new ProcessStartInfo($"{Directory.GetCurrentDirectory()}/tool/krpano/krpano Tools.exe") };
            ts.Start();
        }
        static void RunServer()
        {
            var ts = new Process() { StartInfo = new ProcessStartInfo($"{Directory.GetCurrentDirectory()}/tool/krpano/krpano Testing Server.exe") };
            ts.Start();
        }
        private void Run_Quit(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
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
        private List<string> voicelist = new List<string>();
        private string projname = null;
        private int totalnum = 0;
        private string phonenum = null;
        private string voicemode = null;
        private string inputdir = null;
        private string outputdir = null;
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
            if (Directory.Exists(dst))
                Directory.Delete(dst, true);
            else
            {
                var dstpath = Directory.GetParent(dst).ToString();
                Directory.CreateDirectory(dstpath);
            }
            Directory.Move(src, dst);
        }
        public string ExcelToMemory(string file)
        {
            DataTable data = new ExcelReader().Read(file);
            this.pointnum.Clear();
            this.pointname.Clear();
            this.voicelist.Clear();
            this.projname = data.Rows[0][0].ToString();
            this.totalnum = int.Parse(data.Rows[2][0].ToString());
            this.phonenum = data.Rows[4][0].ToString();
            this.voicemode = data.Rows[6][0].ToString();

            for (var i = 0; i < this.totalnum; i++)
            {
                this.pointnum.Add((int.Parse((string)data.Rows[i][1])));
                this.pointname.Add(data.Rows[i][2].ToString());
                this.voicelist.Add(data.Rows[i][3].ToString());
            }

            this.inputdir = $"{Directory.GetCurrentDirectory()}/input/{this.projname}";
            this.outputdir = $"{Directory.GetCurrentDirectory()}/output/{this.projname}";
            this.exportSpotInfoXML($"{this.outputdir}/spotinfo.xml");
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
        public void MakePano()
        {
            var krpanodir = $"{Directory.GetCurrentDirectory()}/tool/krpano";
            var krpano = new Process()
            {
                StartInfo = new ProcessStartInfo($"{krpanodir}/krpanotools.exe",
                $"makepano \"{krpanodir}/templates/vtour-multires.config\" \"{this.inputdir}/*.jpg\"")
        };
            krpano.Start();
            krpano.WaitForExit();
            krpano.Close();
            File.Delete($"{this.inputdir}/vtour/tour_testingserver.exe");
            File.Delete($"{this.inputdir}/vtour/tour_testingserver_macos");
            this.MoveDir($"{this.inputdir}/vtour", $"{Directory.GetCurrentDirectory()}/output/{this.projname}");
        }
        public void ParseHTML()
        {
            var file = $"{Directory.GetCurrentDirectory()}/output/{this.projname}/tour.html";
            var htmlpath = (new DirectoryInfo(file)).Parent;
            HtmlDocument doc = new HtmlDocument();
            doc.Load(file);
            HtmlTextNode node = doc.DocumentNode.SelectSingleNode("//head//title//text()") as HtmlTextNode;
            node.Text = this.projname;
            File.Delete(file);
            doc.Save($"{htmlpath}/index.html");
        }
        public void ImportTemplate(bool map, bool voice, bool contact)
        {
            var vtourdir = $"{Directory.GetCurrentDirectory()}/output/{this.projname}";
            var templatedir = new DirectoryInfo($"{Directory.GetCurrentDirectory()}/tool/template");
            this.CopyFile($"{templatedir}/vtourskin.xml", $"{vtourdir}/skin/vtourskin.xml");
            this.CopyFile($"{templatedir}/showbutton.xml", $"{vtourdir}/resources/showbutton.xml");
            if (map)
                this.ImportMap();
            if (voice)
                this.ImportVoice();
            if (contact)
                this.ImportContact();
        }
        public void ImportMap()
        {
            var templatedir = new DirectoryInfo($"{Directory.GetCurrentDirectory()}/tool/template");
            this.CopyFile($"{this.inputdir}/map.png", $"{outputdir}/resources/map.png");
            this.CopyFile($"{templatedir}/showmap.xml", $"{outputdir}/resources/showmap.xml");
            this.CopyFile($"{templatedir}/mapon.png", $"{outputdir}/resources/mapon.png");
            this.CopyFile($"{templatedir}/mapoff.png", $"{outputdir}/resources/mapoff.png");
        }
        public void ImportVoice()
        {
            var templatedir = new DirectoryInfo($"{Directory.GetCurrentDirectory()}/tool/template");
            this.CopyFile($"{templatedir}/showvoice.xml", $"{outputdir}/resources/showvoice.xml");
            this.CopyFile($"{templatedir}/voice.png", $"{outputdir}/resources/voice.png");
            this.CopyFile($"{templatedir}/mute.png", $"{outputdir}/resources/mute.png");
            if (this.voicemode == "是")
            {
                this.CopyFile($"{this.inputdir}/{this.voicelist[0]}.mp3", $"{this.outputdir}/resources/voice.mp3");
            }
            else
            {
                for (var index = 0; index < this.totalnum; index++)
                    this.CopyFile($"{this.inputdir}/{this.voicelist[index]}.mp3", $"{this.outputdir}/resources/{this.pointnum[index]}.mp3");
            }
        }
        public void ImportContact()
        {
            var templatedir = new DirectoryInfo($"{Directory.GetCurrentDirectory()}/tool/template");
            this.CopyFile($"{templatedir}/showcontact.xml", $"{outputdir}/resources/showcontact.xml");
            this.CopyFile($"{templatedir}/contact.png", $"{outputdir}/resources/contact.png");
            this.CopyFile($"{templatedir}/copy.png", $"{outputdir}/resources/copy.png");
            this.CopyFile($"{templatedir}/copytext.png", $"{outputdir}/resources/copytext.png");
        }
        public void ParseTOUR(bool map, bool voice, bool contact)
        {
            var file = $"{this.outputdir}/tour.xml";
            XElement root = XElement.Load(file);
            root.Attribute("title").SetValue(this.projname);
            root.Element("skin_settings").Remove();
            XElement skin_settings = new XElement("skin_settings",
                new XAttribute("thumbs_opened", "true"),
                new XAttribute("thumbs_text", "true"),
                new XAttribute("loadingtext", "加载中..."));
            root.Element("action").AddAfterSelf(skin_settings);
            if (map)
            {
                XElement includemap = new XElement("include",
                    new XAttribute("url", "resources/showmap.xml"));
                root.AddFirst(includemap);
            }
            if (voice)
            {
                XElement includevoice = new XElement("include",
                    new XAttribute("url", "resources/showvoice.xml"));
                root.AddFirst(includevoice);
            }
            if (contact)
            {
                XElement includecontact = new XElement("include",
                    new XAttribute("url", "resources/showcontact.xml"));
                root.AddFirst(includecontact);
            }
            var scenelist = root.Elements("scene");
            if (this.voicemode == "是")
                root.SetElementValue("action", "if(startscene === null OR !scene[get(startscene)], copy(startscene,scene[0].name); ); loadscene(get(startscene), null, MERGE); if (startactions !== null, startactions() ); init_voice(vc1, 'resources/1.mp3');");
            foreach (var scene in scenelist)
            {
                int index = int.Parse(scene.Attribute("title").Value);
                scene.SetAttributeValue("title", this.pointname[index - 1]);
                if (this.voicemode == "是")
                    scene.SetAttributeValue("onstart", "activatespot();");
                else
                {
                    if(this.voicemode == "是")
                        scene.SetAttributeValue("onstart", $"activatespot(); init_voice(vc{index}, 'resources/voice.mp3');");
                    else
                        scene.SetAttributeValue("onstart", $"activatespot(); init_voice(vc{index}, 'resources/{index}.mp3');");
                }
                scene.Add(new XElement("layer",
                    new XAttribute("name", "voice_btn"),
                    new XAttribute("style", "voice_btn_style"),
                    new XAttribute("keep", "true")));
                scene.Add(new XElement("layer",
                    new XAttribute("name", "mute_btn"),
                    new XAttribute("style", "mute_btn_style"),
                    new XAttribute("keep", "true")));
            }
            root.Save(file);
        }
        public void ParseContact()
        {
            var file = $"{this.outputdir}/resources/showcontact.xml";
            XElement root = XElement.Load(file);
            foreach (XElement elem in root.Elements("layer"))
            {
                if (elem.Attribute("name").Value.ToString() == "contact_btn")
                    elem.SetAttributeValue("onclick.mobile", $"phonecall({this.phonenum})");
                if (elem.Attribute("name").Value.ToString() == "copy_btn")
                    elem.SetAttributeValue("onclick", $"copytext({this.phonenum}); show_copytext();");
                if (elem.Attribute("name").Value.ToString() == "phonelist")
                {
                    elem.SetAttributeValue("html", this.phonenum);
                    elem.SetAttributeValue("onclick", $"copytext({this.phonenum}); show_copytext();");
                }
            }
            root.Save(file);
        }
    }
}
