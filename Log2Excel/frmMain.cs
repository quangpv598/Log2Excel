using ClosedXML;
using ClosedXML.Excel;
using ExcelDataReader;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;

namespace Log2Excel
{
    public partial class frmMain : Form
    {
        /// <summary>
        /// Path of config file
        /// </summary>
        private readonly string ConfigPath = "Config.xlsx";

        /// <summary>
        /// Save general environment variables for application
        /// </summary>
        private Hashtable htbEnvironmentVariables = new();

        /// <summary>
        /// Map command between DZS vs CISCO
        /// </summary>
        private List<(string, string)> mapCommands = new();

        // Các Device Prefix và Suffix sẽ được lưu dưới dạng key trong Dic
        // Mỗi key sẽ chứa một tập các String
        public Dictionary<string, Queue<string>> dicDeviceLog = new Dictionary<string, Queue<string>>();
        public frmMain()
        {
            InitializeComponent();

            // Get full config
            this.ConfigPath = Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName), ConfigPath);
            this.ReadConfigFile();
        }

        private void ReadConfigFile()
        {
            DataSet dataSet = new DataSet();

            using (var stream = File.Open(this.ConfigPath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                //reader = ExcelReaderFactory.CreateBinaryReader(stream);
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                reader.Close();

            }

            // Environment Variables
            DataTable dtEnvironmentVariables = dataSet.Tables[0];
            foreach (DataRow dr in dtEnvironmentVariables.Rows)
            {
                htbEnvironmentVariables.Add(dr["NAME"], dr["VALUE"]);
            }

            // DZS vs CISCO map command
            DataTable dtMapCommand = dataSet.Tables[1];
            foreach (DataRow dr in dtMapCommand.Rows)
            {
                mapCommands.Add((dr["DZS"].ToString(), dr["CISCO"].ToString()));
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty("txtLogPath.Text"))
            {
                MessageBox.Show("Please select a log file.");
                return;
            }

            //1. Xoá dữ liệu
            dicDeviceLog.Clear();

            //2. Nạp dữ liệu khởi tạo cho từ điển chứa dữ liệu
            var deviceSuffix = "txtDeviceSuffix.Text".Split(';');
            foreach (string suffix in deviceSuffix)
            {
                string fullDeviceName = "txtDevicePrefix.Text" + suffix;
                // Nếu tên thiết bị chưa có trong từ điển, thì thêm vào từ điển
                if (!dicDeviceLog.ContainsKey(fullDeviceName))
                {
                    dicDeviceLog.Add(fullDeviceName, new Queue<string>());
                }
            }

            //3. Duyệt các dòng, ghi dữ liệu vào từ điển
            List<string> lines = new();

            //foreach(var file in fdl.FileNames)
            //{
            //    var currentLines = File.ReadAllLines(file);

            //    lines = lines.Concat(currentLines).ToList();
            //}

            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];

                // Kiểm tra nếu dòng chứa lệnh show
                if (line.Replace(" ", "").ToLower().Contains("#show"))
                {
                    // duyệt qua tên các thiết bị đã lưu trước đó
                    foreach (string fullDeviceName in dicDeviceLog.Keys)
                    {
                        // Nếu dòng mà bắt đầu bằng tên thiết bị
                        if (line.StartsWith(fullDeviceName))
                        {
                            var queueString = dicDeviceLog[fullDeviceName];

                            // Thêm dòng hiện tại
                            queueString.Enqueue(line);

                            // Thêm các dòng kết quả của lệnh show cho tới khi gặp lại tên thiết bị

                            i++;
                            line = lines[i];
                            while (dicDeviceLog.Keys.Where(deviceName => line.StartsWith(deviceName)).Count() == 0)
                            {
                                queueString.Enqueue(line);
                                i++;
                                line = lines[i];
                            }

                            // Nếu ngay dòng tiếp theo sau kết quả của lệnh show tiếp đó
                            // mà là 1 lệnh show mới thì tiến hành trừ i đi 1 để đọc lại đoạn show ở trên
                            if (line.Replace(" ", "").ToLower().Contains("#show"))
                            {
                                i--;
                            }

                            break;
                        }
                    }
                }
            }

            //4. Xuất excel từ bộ dữ liệu từ điển
            DataTable dtResult = new DataTable();
            foreach (string columnName in dicDeviceLog.Keys)
            {
                dtResult.Columns.Add(columnName, typeof(string));
            }

            int maxRow = dicDeviceLog.Values.Max(queue => queue.Count);
            for (int i = 0; i < maxRow; i++)
            {
                DataRow drNewRow = dtResult.NewRow();

                foreach (string columnName in dicDeviceLog.Keys)
                {
                    var queue = dicDeviceLog[columnName];
                    if (queue.Count > 0)
                    {
                        drNewRow[columnName] = queue.Dequeue();
                    }
                }

                dtResult.Rows.Add(drNewRow);
            }

            string fileName = "";//"txtLogPath.Text".Replace(Path.GetExtension(fdl.FileName), ".xlsx");
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(dtResult, "Sample Sheet");
                workbook.SaveAs(fileName);
            }

            MessageBox.Show($"Convert Success.\nFile : {fileName}");
        }

        private void lblOpenConfigFile_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(ConfigPath)
            {
                UseShellExecute = true
            };
            p.Start();
        }

        private void lblConvert_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }
    }

    public class ExportExcelHandler
    {
        public void Export(List<CommandBlockMapper> commandBlockMappers)
        {

        }
    }

    public class CommandBlockManager
    {
        public List<CommandBlockMapper> LoadDataLogFile(string[] files)
        {
            foreach (string file in files)
            {

            }

            return new List<CommandBlockMapper>();
        }
    }

    public class CommandBlockMapper
    {
        public Guid Id = new Guid();
        public readonly CommandBlock DzsCommand;
        public readonly CommandBlock CiscoCommand;

        public string Router
        {
            get
            {
                if (DzsCommand != null
                    && CiscoCommand != null)
                {
                    if (DzsCommand.Router == CiscoCommand.Router)
                    {
                        return DzsCommand.Router;
                    }

                    throw new Exception("Router not match!!!");
                }

                return string.Empty;
            }
        }
        public CommandBlockMapper(CommandBlock dzsCommand, CommandBlock ciscoCommand)
        {
            DzsCommand = dzsCommand;
            CiscoCommand = ciscoCommand;
        }
    }

    public class CommandBlock
    {
        public Guid Id = new Guid();
        public readonly string Router;
        public readonly Queue<string> Contents;
        public readonly string Command;
        public CommandBlock(string router, Queue<string> contents, string command)
        {
            Router = router;
            Contents = contents;
            Command = command;
        }

        public int Length => Contents == null ? 0 : Contents.Count;
    }
}