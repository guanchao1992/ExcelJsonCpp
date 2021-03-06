using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using tablegen2.common;
using tablegen2.layouts;
using tablegen2.logic;

namespace tablegen2
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            AppData.MainWindow = this;

            InitializeComponent();
            
            setting.ExcelDirChanged += () => tree.refreshExcelPath(AppData.Config.ExcelDir);
            setting.ExportFormatChanged += () => refreshButtonGenAll();
            tree.OpenExcelRequest += () => setting.browseExcelDirectory();

            if (AppData.Config != null)
                refreshButtonGenAll();
        }

        #region 事件处理
        private void Window_PreviewKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                _flipHelpPanel();
                e.Handled = true;
            }
        }

        private void btnHelp_Clicked(object sender, RoutedEventArgs e)
        {
            _flipHelpPanel();
        }

        private void btnGenAll_Clicked(object sender, RoutedEventArgs e)
        {
            string excelDir = AppData.Config.ExcelDir;
            string exportDir = AppData.Config.ExportDir;
            TableExportFormat fmt = AppData.Config.ExportFormat;

            if (string.IsNullOrEmpty(excelDir) || !Directory.Exists(excelDir))
            {
                Log.Err("请选择合法的Excel配置目录！");
                return;
            }

            if (string.IsNullOrEmpty(exportDir) || !Directory.Exists(exportDir))
            {
                Log.Err("请选择合法的导出目录！");
                return;
            }

            if (fmt == TableExportFormat.Unknown)
            {
                Log.Err("请选择导出数据格式！");
                return;
            }

            var excels = tree.AllExcels.ToList();
            if (excels.Count == 0)
            {
                Log.Wrn("您选择的配置目录中不包含任何Excel文件！ 目录：{0}", excelDir);
                return;
            }

            Log.Msg("=================================================");
            foreach (var filePath in excels)
            {
                _genSingleFileImpl(filePath, exportDir, fmt);
            }
        }
        #endregion

        public void addMessage(string msg, Color color)
        {
            if (console != null)
                console.addMessage(msg, color);
        }

        public void refreshButtonGenAll()
        {
            switch (AppData.Config.ExportFormat)
            {
                case TableExportFormat.Unknown:
                    btnGenAll.IsEnabled = false;
                    btnGenAll.Content = "请选择生成的数据格式";
                    break;
                default:
                    btnGenAll.IsEnabled = true;
                    btnGenAll.Content = string.Format(btnGenAll.Tag as string, AppData.Config.ExportFormat.ToString());
                    break;
            }
        }

        public void genSingleFile(string filePath, string exportDir, TableExportFormat fmt)
        {
            if (!File.Exists(filePath))
            {
                Log.Err("源文件不存在！ {0}", filePath);
                return;
            }

            if (!Directory.Exists(exportDir))
            {
                Log.Err("导出目录不存在！ {0}", exportDir);
                return;
            }

            if (fmt == TableExportFormat.Unknown)
            {
                Log.Err("导出数据格式不合法！");
                return;
            }

            Log.Msg("=================================================");
            _genSingleFileImpl(filePath, exportDir, fmt);
        }
        
        private void _genSingleFileImpl(string filePath, string exportDir, TableExportFormat fmt)
        {
            Log.Msg("正在分析 {0}", filePath);
            try
            {
                TableExcelData data = TableExcelReader.loadFromExcel(filePath);
                string errmsg;
                if (!data.checkUnique(out errmsg))
                    Log.Wrn(errmsg);
                switch (fmt)
                {
                    case TableExportFormat.Json:
                        {
                            var exportPath = Path.Combine(exportDir, string.Format("{0}.json", Path.GetFileNameWithoutExtension(filePath)));
                            TableExcelExportJson.exportExcelFile(data, exportPath);
                        }
                        break;
                        /*
                    case TableExportFormat.Dat:
                        {
                            var exportPath = Path.Combine(exportDir, string.Format("{0}.exdat", Path.GetFileNameWithoutExtension(filePath)));
                            TableExcelExportDat.exportExcelFile(data, exportPath);
                        }
                        break;
                    case TableExportFormat.Xml:
                        {
                            var exportPath = Path.Combine(exportDir, string.Format("{0}.xml", Path.GetFileNameWithoutExtension(filePath)));
                            TableExcelExportXml.exportExcelFile(data, exportPath);
                        }
                        break;
                    case TableExportFormat.Lua:
                        {
                            var exportPath = Path.Combine(exportDir, string.Format("{0}.lua", Path.GetFileNameWithoutExtension(filePath)));
                            TableExcelExportLua.exportExcelFile(data, exportPath);
                        }
                        break;
                        */
                }
                //生成.h文件
                {
                    var exportPath = Path.Combine(exportDir, string.Format("{0}.h", Path.GetFileNameWithoutExtension(filePath)));
                    var outCppData = CreateTableCpp.toFileData(Path.GetFileNameWithoutExtension(filePath), data.Headers);
                    File.WriteAllBytes(exportPath, Encoding.UTF8.GetBytes(outCppData));
                }

                Log.Msg("生成成功");
            }
            catch (System.Exception ex)
            {
                Log.Err(ex.Message);
            }
        }

        private void _flipHelpPanel()
        {
            var hp = new HelperPanel();
            var pw = new PopupWindow(hp);
            pw.Owner = Window.GetWindow(this);
            pw.Title = "使用说明";
            pw.ShowDialog();
        }
    }
}
