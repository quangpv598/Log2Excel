using Accessibility;
using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
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
        private List<CommandMapper> mapCommands = new();


        private List<DeviceConfig> deviceConfigs = new();

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
                mapCommands.Add(new CommandMapper
                {
                    Key = dr["KEY"].ToString(),
                    DzsCommand = dr["DZS"].ToString(),
                    CiscoCommnad = dr["CISCO"].ToString()
                });
            }

            // Load device config
            var routers = htbEnvironmentVariables["ROUTERS"].ToString().Split(';');
            var dszPrefixDevices = htbEnvironmentVariables["DZS_PREFIX"].ToString().Split(';');
            var ciscoPrefixDevices = htbEnvironmentVariables["CISCO_PREFIX"].ToString().Split(';');
            for (int i = 0; i < routers.Length; i++)
            {
                string router = routers[i];
                string dzsPrefix = dszPrefixDevices[i];
                string ciscoPrefix = ciscoPrefixDevices[i];

                deviceConfigs.Add(new DeviceConfig
                {
                    Prefix = dzsPrefix,
                    Router = router,
                    DeviceType = DeviceType.DZS
                });

                deviceConfigs.Add(new DeviceConfig
                {
                    Prefix = ciscoPrefix,
                    Router = router,
                    DeviceType = DeviceType.CISCO
                });
            }
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
            CommandBlockManager commandBlockManager = new CommandBlockManager(
                rounters: htbEnvironmentVariables["ROUTERS"].ToString().Split(';'),
                deviceConfigs: deviceConfigs,
                mapCommands: mapCommands);

            string[] files = Directory.GetFiles(htbEnvironmentVariables["INPUT_PATH"].ToString());
            var dicData = commandBlockManager.LoadDataLogFile(files);

            ExportExcelHandler exportExcelHandler = new();
            exportExcelHandler.Export(dicData: dicData,
                deviceConfigs: deviceConfigs,
                rounters: htbEnvironmentVariables["ROUTERS"].ToString().Split(';'),
                commentName: htbEnvironmentVariables["COMMENT_NAME"].ToString(),
                mapCommands: mapCommands,
                outputPath: htbEnvironmentVariables["OUTPUT_PATH"].ToString());
        }
    }

    public class ExportExcelHandler
    {
        public void Export(Dictionary<string, Dictionary<string, Dictionary<string, Queue<string>>>> dicData,
            List<DeviceConfig> deviceConfigs,
            string[] rounters,
            string commentName,
            List<CommandMapper> mapCommands,
            string outputPath)
        {
            string fileName = Path.Combine(outputPath, $"output_{Guid.NewGuid()}.xlsx");
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("RESULT");

                worksheet.Style.Font.FontSize = 13;
                worksheet.Style.Font.FontName = "Consolas";

                // Generate Header
                int headerRowIndex = 1;
                for (int i = 0; i < rounters.Length; i++)
                {
                    // 3 = 2 cột cho router (dzs vs cisco) + 1 cột cho comment
                    int cellIndex = i * 3 + 1;

                    var cell = worksheet.Cell(headerRowIndex, cellIndex);

                    cell.Value = rounters[i];
                    cell.Style.Fill.BackgroundColor = XLColor.Black;
                    cell.Style.Font.FontColor = XLColor.White;
                    worksheet.Range(headerRowIndex, cellIndex, headerRowIndex, cellIndex + 1).Row(1).Merge();

                    var commentCell = worksheet.Cell(headerRowIndex, cellIndex + 2);
                    commentCell.Value = commentName;
                    commentCell.Style.Fill.BackgroundColor = XLColor.LightCyan;
                    commentCell.Style.Font.FontColor = XLColor.Black;

                    // Next cell 
                    var secondHeaderCellDZS = worksheet.Cell(headerRowIndex + 1, cellIndex);
                    var secondHeaderCellCISCO = worksheet.Cell(headerRowIndex + 1, cellIndex + 1);
                    var secondHeaderCellComment = worksheet.Cell(headerRowIndex + 1, cellIndex + 2);
                    secondHeaderCellDZS.Value = "DZS";
                    secondHeaderCellCISCO.Value = "CISCO";
                    secondHeaderCellDZS.Style.Fill.BackgroundColor = XLColor.DarkRed;
                    secondHeaderCellDZS.Style.Font.FontColor = XLColor.White;
                    secondHeaderCellCISCO.Style.Fill.BackgroundColor = XLColor.DarkRed;
                    secondHeaderCellCISCO.Style.Font.FontColor = XLColor.White;

                    secondHeaderCellComment.Style.Fill.BackgroundColor = XLColor.LightCyan;
                    worksheet.Range(commentCell, secondHeaderCellComment).Column(1).Merge();

                    cell.WorksheetColumn().Width = 160;
                    worksheet.Cell(headerRowIndex, cellIndex + 1).WorksheetColumn().Width = 160;
                    commentCell.WorksheetColumn().Width = 30;
                }

                worksheet.Row(headerRowIndex).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Row(headerRowIndex + 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Row(headerRowIndex).Style.Font.FontSize = 20;
                worksheet.Row(headerRowIndex + 1).Style.Font.FontSize = 20;
                worksheet.Row(headerRowIndex).Style.Alignment.WrapText = true;
                worksheet.Row(headerRowIndex + 1).Style.Alignment.WrapText = true;


                // Generate body
                int currentRowIndex = 3;

                for (int i = 0; i < mapCommands.Count; i++)
                {
                    var key = mapCommands[i].Key;
                    var dicRouters = dicData[key];

                    int maxRouterLines = 0;
                    int routerRowIndex = currentRowIndex;
                    worksheet.Range(currentRowIndex, 1, currentRowIndex, rounters.Length * 3).Style.Fill.BackgroundColor = XLColor.Yellow;

                    for (int j = 0; j < rounters.Length; j++)
                    {
                        var rounter = rounters[j];
                        var devices = dicRouters[rounter];

                        var devicesName = deviceConfigs.Where(d => d.Router == rounter)
                            .OrderBy(d => d.DeviceType);

                        for (int k = 0; k < devicesName.Count(); k++)
                        {
                            currentRowIndex = routerRowIndex;

                            var queueLog = devices[devicesName.ElementAt(k).Name];

                            if (maxRouterLines < queueLog.Count())
                            {
                                maxRouterLines = queueLog.Count();
                            }

                            while (queueLog.Count() > 0)
                            {
                                string content = queueLog.Dequeue();

                                int cellIndex = j * 3 + 1 + k;
                                worksheet.Cell(currentRowIndex, cellIndex).Value = content;
                                currentRowIndex++;
                            }

                        }
                    }

                    currentRowIndex = routerRowIndex + maxRouterLines + 2; // 2 là 2 dòng trống

                    worksheet.Range(currentRowIndex, 1, currentRowIndex, rounters.Length * 3).Style.Fill.BackgroundColor = XLColor.Black;
                    currentRowIndex++;
                }

                workbook.SaveAs(fileName);
            }

            MessageBox.Show($"Convert Success.");

            var p = new Process();
            p.StartInfo = new ProcessStartInfo(fileName)
            {
                UseShellExecute = true
            };
            p.Start();
        }
    }

    public class CommandBlockManager
    {
        private readonly string[] _routers;
        private readonly List<CommandMapper> _mapCommands;
        private readonly string[] _combieCommands;
        private readonly List<DeviceConfig> _deviceConfigs;

        public CommandBlockManager(string[] rounters,
            List<DeviceConfig> deviceConfigs,
            List<CommandMapper> mapCommands)
        {
            _routers = rounters;
            _mapCommands = mapCommands;
            _combieCommands = _mapCommands.Select(c => c.DzsCommand)
                                    .Concat(_mapCommands.Select(c => c.CiscoCommnad))
                                    .Where(c => !string.IsNullOrEmpty(c))
                                    .ToArray();
            _deviceConfigs = deviceConfigs;
        }


        /// <summary>
        /// Đọc file, chuyển sang một mảng 2 chiều các phần tử
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        public Dictionary<string, Dictionary<string, Dictionary<string, Queue<string>>>> LoadDataLogFile(string[] files)
        {
            // 2. Tạo bảng lưu trữ
            /* Cấu trúc lưu trữ data
            - INDEX : 1
                + R1 
                    + DZS_DEVICE_NAME - Queue<string>
                    + CISO_DEVICE_NAME - Queue<string>
                + R2
                    + DZS_DEVICE_NAME - Queue<string>
                    + CISO_DEVICE_NAME - Queue<string>
                ....
            - INDEX : 2
                + R1 
                    + DZS_DEVICE_NAME - Queue<string>
                    + CISO_DEVICE_NAME - Queue<string>
                + R2
                    + DZS_DEVICE_NAME - Queue<string>
                    + CISO_DEVICE_NAME - Queue<string>
                ....
            - INDEX : n
            */
            Dictionary<string, Dictionary<string, Dictionary<string, Queue<string>>>> dicData = new();
            foreach (var mapCommand in _mapCommands)
            {
                string index = mapCommand.Key;

                Dictionary<string, Dictionary<string, Queue<string>>> dicRouters = new();
                foreach (var router in _routers)
                {
                    Dictionary<string, Queue<string>> dicDevices = new();
                    foreach (var deviceConfig in _deviceConfigs.Where(d => d.Router == router))
                    {
                        dicDevices.Add($"{deviceConfig.Prefix}{deviceConfig.Router}", new Queue<string>());
                    }

                    dicRouters.Add(router, dicDevices);
                }

                dicData.Add(index, dicRouters);
            }


            //3. Duyệt các dòng, ghi dữ liệu vào Hashtable
            List<string> lines = new();

            foreach (var file in files)
            {
                var currentLines = File.ReadAllLines(file);

                lines = lines.Concat(currentLines).ToList();
            }

            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];

                foreach (var mapCommand in _mapCommands)
                {
                    bool isBreakManually = false;
                    string key = mapCommand.Key;

                    string dzsCommand = mapCommand.DzsCommand;
                    string ciscoCommand = mapCommand.CiscoCommnad;

                    foreach (string command in new string[] { dzsCommand, ciscoCommand })
                    {
                        var elements = line.Split('#');
                        if (elements.Length < 2)
                        {
                            continue;
                        }

                        string commandInLine = elements[1];

                        if (!string.IsNullOrEmpty(command) && commandInLine.Trim() == command.Trim())
                        {
                            var currentDeviceConfig = _deviceConfigs
                                .FirstOrDefault(d => line.StartsWith(d.Prefix) && line.Contains(d.Router));

                            if (currentDeviceConfig != null)
                            {
                                var dicRouter = dicData[key][currentDeviceConfig.Router];
                                var queueLog = dicRouter[$"{currentDeviceConfig.Prefix}{currentDeviceConfig.Router}"] as Queue<string>;

                                // Thêm dòng hiện tại
                                queueLog.Enqueue(line);

                                // Thêm các dòng kết quả của lệnh show cho tới khi gặp lại tên thiết bị

                                i++;
                                line = lines[i];
                                while (_deviceConfigs.Count(d => line.StartsWith(d.Prefix) && line.Contains(d.Router)) == 0)
                                {
                                    queueLog.Enqueue(line);
                                    i++;
                                    line = lines[i];
                                }

                                // Nếu ngay dòng tiếp theo sau kết quả của lệnh show tiếp đó
                                // mà là 1 lệnh show mới thì tiến hành trừ i đi 1 để đọc lại đoạn show ở trên
                                foreach (var commandString in _combieCommands)
                                {
                                    if (line.ToLower().Contains(commandString))
                                    {
                                        i--;
                                    }
                                }

                                isBreakManually = true;
                                break;
                            }
                        }
                    }

                    if (isBreakManually)
                    {
                        break;
                    }
                }
            }

            return dicData;
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

    public class DeviceConfig
    {
        public string Prefix { get; set; }
        public string Router { get; set; }
        public DeviceType DeviceType { get; set; }
        public string Name => Prefix + Router;
    }

    public class CommandMapper
    {
        public string Key;
        public string DzsCommand { get; set; }
        public string CiscoCommnad { get; set; }
    }

    public enum DeviceType
    {
        DZS,
        CISCO
    }
}