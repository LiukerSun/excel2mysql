using System;
using System.Windows.Forms;
using System.Data;
using MySql.Data.MySqlClient;
using CsvHelper;
using System.Globalization;
using System.IO;
using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text.RegularExpressions;
using Color = System.Drawing.Color;
using Font = System.Drawing.Font;
using Control = System.Windows.Forms.Control;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using System.Threading;
using System.Text;

namespace WindowsFormsApp
{
    public class MainForm : Form
    {
        private readonly Button btnSelectFile;
        private readonly TextBox txtFilePath;
        private readonly Label lblFilePath;
        private readonly OpenFileDialog openFileDialog;

        // 数据库选择控件
        private readonly Label lblDatabase;
        private readonly ComboBox cmbDatabase;
        private readonly Button btnRefreshDatabases;
        private readonly Button btnDatabaseSettings;

        // 导入按钮
        private readonly Button btnImport;

        // 状态标签
        private readonly Label lblStatus;

        // 日志和进度显示
        private readonly RichTextBox txtLog;
        private readonly ProgressBar progressBar;
        private readonly Label lblProgress;

        private readonly DatabaseConfig dbConfig;
        private bool isDatabaseConnected = false;

        public MainForm()
        {
            // 初始化所有控件
            btnSelectFile = new Button();
            txtFilePath = new TextBox();
            lblFilePath = new Label();
            openFileDialog = new OpenFileDialog();
            
            lblDatabase = new Label();
            cmbDatabase = new ComboBox();
            btnRefreshDatabases = new Button();
            btnDatabaseSettings = new Button();
            
            btnImport = new Button();
            lblStatus = new Label();

            txtLog = new RichTextBox();
            progressBar = new ProgressBar();
            lblProgress = new Label();
            
            dbConfig = DatabaseConfig.Load();

            InitializeComponents();
            LoadDatabaseConfig();
            UpdateControlStates();
        }

        private void InitializeComponents()
        {
            // 窗体基本设置
            this.Text = "Excel/CSV导入工具";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Padding = new Padding(10);

            // 配置文件选择部分
            lblFilePath.Text = "文件路径：";
            lblFilePath.Location = new System.Drawing.Point(20, 20);
            lblFilePath.AutoSize = true;

            txtFilePath.Location = new System.Drawing.Point(20, 50);
            txtFilePath.Width = 600;
            txtFilePath.ReadOnly = true;

            btnSelectFile.Text = "选择文件";
            btnSelectFile.Location = new System.Drawing.Point(640, 48);
            btnSelectFile.Width = 100;
            btnSelectFile.Click += BtnSelectFile_Click;

            openFileDialog.Filter = "Excel文件|*.xlsx;*.xls|CSV文件|*.csv|所有文件|*.*";
            openFileDialog.Title = "选择Excel或CSV文件";

            // 数据库选择部分
            lblDatabase.Text = "数据库：";
            lblDatabase.Location = new System.Drawing.Point(20, 100);
            lblDatabase.AutoSize = true;

            cmbDatabase.Location = new System.Drawing.Point(20, 130);
            cmbDatabase.Width = 300;
            cmbDatabase.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbDatabase.SelectedIndexChanged += CmbDatabase_SelectedIndexChanged;

            btnRefreshDatabases.Text = "刷新";
            btnRefreshDatabases.Location = new System.Drawing.Point(330, 129);
            btnRefreshDatabases.Width = 80;
            btnRefreshDatabases.Click += BtnRefreshDatabases_Click;

            btnDatabaseSettings.Text = "数据库设置";
            btnDatabaseSettings.Location = new System.Drawing.Point(420, 129);
            btnDatabaseSettings.Width = 100;
            btnDatabaseSettings.Click += BtnDatabaseSettings_Click;

            // 状态标签
            lblStatus.Location = new System.Drawing.Point(20, 170);
            lblStatus.AutoSize = true;
            lblStatus.ForeColor = System.Drawing.Color.Gray;

            // 导入按钮
            btnImport.Text = "导入数据";
            btnImport.Location = new System.Drawing.Point(20, 200);
            btnImport.Width = 100;
            btnImport.Click += BtnImport_Click;
            btnImport.Enabled = false;

            // 进度条
            lblProgress.Text = "进度：";
            lblProgress.Location = new System.Drawing.Point(20, 250);
            lblProgress.AutoSize = true;

            progressBar.Location = new System.Drawing.Point(20, 280);
            progressBar.Size = new System.Drawing.Size(740, 23);
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Value = 0;
            progressBar.Style = ProgressBarStyle.Continuous;

            // 日志框
            txtLog.Location = new System.Drawing.Point(20, 320);
            txtLog.Size = new System.Drawing.Size(740, 220);
            txtLog.ReadOnly = true;
            txtLog.BackColor = Color.White;
            txtLog.Font = new Font("Consolas", 9F);
            txtLog.ScrollBars = RichTextBoxScrollBars.Both;
            txtLog.WordWrap = false;

            // 添加控件到窗体
            this.Controls.AddRange(new Control[] {
                lblFilePath, txtFilePath, btnSelectFile,
                lblDatabase, cmbDatabase, btnRefreshDatabases, btnDatabaseSettings,
                lblStatus,
                btnImport,
                lblProgress, progressBar,
                txtLog
            });

            // 初始日志
            LogMessage("程序已启动");
        }

        private void LoadDatabaseConfig()
        {
            if (!string.IsNullOrEmpty(dbConfig.Database))
            {
                LoadDatabases(selectSavedDatabase: true);
            }
        }

        private void UpdateControlStates()
        {
            bool hasFile = !string.IsNullOrEmpty(txtFilePath.Text);
            bool hasDatabase = cmbDatabase.SelectedItem != null;

            // 更新按钮状态
            btnImport.Enabled = hasFile && hasDatabase && isDatabaseConnected;

            // 更新状态显示
            if (!hasFile)
            {
                lblStatus.Text = "请选择要导入的文件";
                lblStatus.ForeColor = System.Drawing.Color.Gray;
            }
            else if (!hasDatabase || !isDatabaseConnected)
            {
                lblStatus.Text = "请选择数据库";
                lblStatus.ForeColor = System.Drawing.Color.Gray;
            }
            else
            {
                lblStatus.Text = "就绪";
                lblStatus.ForeColor = System.Drawing.Color.Green;
            }
        }

        private void LoadDatabases(bool selectSavedDatabase = false)
        {
            try
            {
                LogMessage("开始加载数据库列表...");
                Cursor = Cursors.WaitCursor;
                cmbDatabase.Items.Clear();
                isDatabaseConnected = false;

                var connectionStringBuilder = new MySqlConnectionStringBuilder
                {
                    Server = dbConfig.Server,
                    Port = (uint)dbConfig.Port,
                    UserID = dbConfig.Username,
                    Password = dbConfig.Password
                };

                LogMessage($"正在连接到服务器 {dbConfig.Server}:{dbConfig.Port}...");
                using var connection = new MySqlConnection(connectionStringBuilder.ConnectionString);
                connection.Open();
                LogMessage("服务器连接成功");

                using var command = new MySqlCommand("SHOW DATABASES", connection);
                using var reader = command.ExecuteReader();

                int databaseCount = 0;
                while (reader.Read())
                {
                    string dbName = reader.GetString(0);
                    if (!dbName.Equals("information_schema", StringComparison.OrdinalIgnoreCase) &&
                        !dbName.Equals("mysql", StringComparison.OrdinalIgnoreCase) &&
                        !dbName.Equals("performance_schema", StringComparison.OrdinalIgnoreCase) &&
                        !dbName.Equals("sys", StringComparison.OrdinalIgnoreCase))
                    {
                        cmbDatabase.Items.Add(dbName);
                        databaseCount++;
                    }
                }

                LogMessage($"成功加载 {databaseCount} 个数据库");

                if (cmbDatabase.Items.Count > 0)
                {
                    if (selectSavedDatabase && !string.IsNullOrEmpty(dbConfig.Database))
                    {
                        int index = cmbDatabase.Items.IndexOf(dbConfig.Database);
                        if (index >= 0)
                        {
                            cmbDatabase.SelectedIndex = index;
                            LogMessage($"已选择上次使用的数据库: {dbConfig.Database}");
                        }
                        else
                        {
                            cmbDatabase.SelectedIndex = 0;
                            LogMessage($"未找到上次使用的数据库 {dbConfig.Database}，已选择第一个数据库");
                        }
                    }
                    else
                    {
                        cmbDatabase.SelectedIndex = 0;
                        LogMessage("已选择第一个数据库");
                    }
                }

                isDatabaseConnected = true;
                lblStatus.Text = "数据库连接成功";
                lblStatus.ForeColor = Color.Green;
                LogMessage("数据库连接状态：已连接", false);
            }
            catch (Exception ex)
            {
                LogMessage($"加载数据库列表失败: {ex.Message}", true);
                LogMessage($"详细错误: {ex}", true);
                MessageBox.Show($"加载数据库列表失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "数据库连接失败";
                lblStatus.ForeColor = Color.Red;
            }
            finally
            {
                Cursor = Cursors.Default;
                UpdateControlStates();
            }
        }

        private void CmbDatabase_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cmbDatabase.SelectedItem != null)
            {
                dbConfig.Database = cmbDatabase.SelectedItem.ToString() ?? "";
            }
            UpdateControlStates();
        }

        private void BtnSelectFile_Click(object? sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = openFileDialog.FileName;
                UpdateControlStates();
            }
        }

        private void BtnDatabaseSettings_Click(object? sender, EventArgs e)
        {
            using var settingsForm = new DatabaseSettingsForm(dbConfig);
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                isDatabaseConnected = false;
                LoadDatabases(selectSavedDatabase: true);
            }
        }

        private void BtnRefreshDatabases_Click(object? sender, EventArgs e)
        {
            LoadDatabases();
        }

        private async Task<DataTable> ReadExcelFileAsync(string filePath, ImportSettings settings, CancellationToken cancellationToken = default)
        {
            var dt = new DataTable();
            LogMessage($"开始读取Excel文件: {filePath}");

            try
            {
                using var document = SpreadsheetDocument.Open(filePath, false);
                var workbookPart = document.WorkbookPart ?? throw new Exception("工作簿部分不能为空");
                
                // 获取所有工作表
                var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList() 
                    ?? throw new Exception("未找到工作表");
                
                if (!sheets.Any())
                {
                    throw new Exception("Excel文件中没有工作表");
                }
                
                LogMessage($"Excel文件包含 {sheets.Count} 个工作表");

                // 选择要导入的工作表
                Sheet selectedSheet;
                if (sheets.Count == 1)
                {
                    selectedSheet = sheets[0];
                    if (selectedSheet == null)
                    {
                        throw new Exception("无法获取工作表");
                    }
                    LogMessage($"使用唯一的工作表: {selectedSheet.Name?.Value ?? "(未命名)"}");
                }
                else
                {
                    // 创建工作表选择对话框
                    using var sheetDialog = new Form
                    {
                        Text = "选择工作表",
                        Size = new System.Drawing.Size(300, 200),
                        StartPosition = FormStartPosition.CenterParent,
                        FormBorderStyle = FormBorderStyle.FixedDialog,
                        MaximizeBox = false,
                        MinimizeBox = false
                    };

                    var listBox = new ListBox
                    {
                        Dock = DockStyle.Top,
                        Height = 120
                    };

                    // 处理可能为null的工作表名称
                    var sheetNames = sheets.Select(s => s.Name?.Value ?? "(未命名)").ToArray();
                    listBox.Items.AddRange(sheetNames.Cast<object>().ToArray());
                    listBox.SelectedIndex = 0;

                    var btnOK = new Button
                    {
                        Text = "确定",
                        DialogResult = DialogResult.OK,
                        Dock = DockStyle.Bottom
                    };

                    sheetDialog.Controls.AddRange(new Control[] { listBox, btnOK });

                    if (sheetDialog.ShowDialog() != DialogResult.OK)
                    {
                        throw new Exception("用户取消了工作表选择");
                    }

                    if (listBox.SelectedIndex < 0 || listBox.SelectedIndex >= sheets.Count)
                    {
                        throw new Exception("无效的工作表选择");
                    }

                    selectedSheet = sheets[listBox.SelectedIndex];
                    if (selectedSheet == null)
                    {
                        throw new Exception("无法获取所选工作表");
                    }
                    LogMessage($"用户选择了工作表: {selectedSheet.Name?.Value ?? "(未命名)"}");
                }

                // 获取选定工作表的数据
                if (selectedSheet.Id == null)
                {
                    throw new Exception($"工作表 '{selectedSheet.Name?.Value ?? "(未命名)"}' 的ID为空");
                }

                // 验证工作表ID的有效性
                string worksheetId = selectedSheet.Id.Value;
                if (string.IsNullOrWhiteSpace(worksheetId))
                {
                    throw new Exception($"工作表 '{selectedSheet.Name?.Value ?? "(未命名)"}' 的ID无效");
                }

                // 尝试获取工作表部分
                OpenXmlPart? worksheetPart;
                try
                {
                    worksheetPart = workbookPart.GetPartById(worksheetId);
                }
                catch (ArgumentException ex)
                {
                    throw new Exception($"无法获取工作表 '{selectedSheet.Name?.Value ?? "(未命名)"}' 的内容: {ex.Message}");
                }

                if (worksheetPart == null)
                {
                    throw new Exception($"工作表 '{selectedSheet.Name?.Value ?? "(未命名)"}' 的内容为空");
                }

                if (!(worksheetPart is WorksheetPart))
                {
                    throw new Exception($"工作表 '{selectedSheet.Name?.Value ?? "(未命名)"}' 的类型不正确");
                }

                var worksheet = (WorksheetPart)worksheetPart;

                // 预先缓存共享字符串表
                var sharedStringCache = new string[0];
                var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (sharedStringTable != null)
                {
                    sharedStringCache = sharedStringTable
                        .Elements<SharedStringItem>()
                        .Select(item => item.InnerText ?? string.Empty)
                        .ToArray();
                }

                // 使用OpenXMLReader流式读取数据
                using var reader = OpenXmlReader.Create(worksheet) 
                    ?? throw new Exception("无法创建XML读取器");

                bool isFirstRow = true;
                var columnNames = new Dictionary<string, int>();
                Row? currentRow = null;
                int rowCount = 0;
                int processedRows = 0;
                
                // 首先计算总行数
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Row))
                    {
                        rowCount++;
                    }
                }
                
                LogMessage($"检测到总行数: {rowCount}");
                reader.Close();

                // 重新打开reader开始处理数据
                using var reader2 = OpenXmlReader.Create(worksheet)
                    ?? throw new Exception("无法创建XML读取器");
                
                // 创建一个缓冲区来存储待处理的行
                var rowBuffer = new List<DataRow>();
                const int bufferSize = 1000; // 缓冲区大小

                while (reader2.Read())
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        throw new OperationCanceledException();
                    }

                    if (reader2.ElementType == typeof(Row))
                    {
                        var element = reader2.LoadCurrentElement();
                        if (element == null)
                        {
                            LogMessage("警告：无法加载行元素", true);
                            continue;
                        }

                        currentRow = element as Row;
                        if (currentRow == null)
                        {
                            LogMessage("警告：跳过无效的行数据", true);
                            continue;
                        }
                        
                        if (isFirstRow)
                        {
                            // 处理表头
                            var headerCells = currentRow.Elements<Cell>().ToList();
                            if (!headerCells.Any())
                            {
                                throw new Exception("表头行为空");
                            }

                            foreach (var cell in headerCells)
                            {
                                string columnName = GetCellValue(cell, sharedStringCache);
                                if (string.IsNullOrEmpty(columnName))
                                {
                                    columnName = $"Column{dt.Columns.Count + 1}";
                                }
                                else if (settings.TrimStrings)
                                {
                                    columnName = columnName.Trim();
                                }

                                if (columnNames.ContainsKey(columnName))
                                {
                                    columnNames[columnName]++;
                                    columnName = $"{columnName}_{columnNames[columnName]}";
                                }
                                else
                                {
                                    columnNames[columnName] = 1;
                                }

                                dt.Columns.Add(columnName);
                            }
                            isFirstRow = false;
                            LogMessage($"成功读取表头，共 {dt.Columns.Count} 列");
                        }
                        else
                        {
                            // 处理数据行
                            var dataRow = dt.NewRow();
                            int currentColumn = 0;

                            foreach (var cell in currentRow.Elements<Cell>())
                            {
                                string? cellReference = cell.CellReference?.Value;
                                if (cellReference != null)
                                {
                                    int columnIndex = GetColumnIndexFromReference(cellReference);
                                    while (currentColumn < columnIndex && currentColumn < dt.Columns.Count)
                                    {
                                        dataRow[currentColumn] = DBNull.Value;
                                        currentColumn++;
                                    }
                                }

                                if (currentColumn < dt.Columns.Count)
                                {
                                    string value = GetCellValue(cell, sharedStringCache);
                                    if (settings.TrimStrings)
                                    {
                                        value = value.Trim();
                                    }
                                    dataRow[currentColumn] = value;
                                    currentColumn++;
                                }
                            }

                            while (currentColumn < dt.Columns.Count)
                            {
                                dataRow[currentColumn] = DBNull.Value;
                                currentColumn++;
                            }

                            // 检查是否所有列都为空
                            bool isRowEmpty = true;
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                var value = dataRow[i];
                                if (value != DBNull.Value && value != null && 
                                    !(value is string strValue && string.IsNullOrWhiteSpace(strValue)))
                                {
                                    isRowEmpty = false;
                                    break;
                                }
                            }

                            if (!isRowEmpty)
                            {
                                rowBuffer.Add(dataRow);
                                processedRows++;
                                if (processedRows % 1000 == 0)
                                {
                                    int progress = (int)((double)processedRows / rowCount * 100);
                                    UpdateProgress(progress, $"正在读取Excel数据: {processedRows}/{rowCount}");
                                    await Task.Delay(1); // 让UI有机会响应
                                }
                            }
                            else
                            {
                                LogMessage($"跳过空行 {currentRow.RowIndex}");
                            }

                            // 当缓冲区满时，批量添加到DataTable
                            if (rowBuffer.Count >= bufferSize)
                            {
                                foreach (var row in rowBuffer)
                                {
                                    dt.Rows.Add(row);
                                }
                                rowBuffer.Clear();
                                
                                // 释放一些内存
                                if (processedRows % (bufferSize * 10) == 0)
                                {
                                    GC.Collect(0, GCCollectionMode.Optimized, true);
                                }
                            }
                        }
                    }
                }

                // 添加剩余的行
                if (rowBuffer.Count > 0)
                {
                    foreach (var row in rowBuffer)
                    {
                        dt.Rows.Add(row);
                    }
                }

                if (dt.Rows.Count == 0)
                {
                    throw new Exception("Excel文件中没有有效数据行");
                }

                LogMessage($"Excel数据读取完成，共读取 {dt.Rows.Count} 行有效数据");
                return dt;
            }
            catch (OperationCanceledException)
            {
                LogMessage("用户取消了数据读取操作");
                throw;
            }
            catch (Exception ex)
            {
                LogMessage($"读取Excel文件时发生错误: {ex.Message}", true);
                LogMessage($"详细错误: {ex}", true);
                throw;
            }
            finally
            {
                UpdateProgress(0);
            }
        }

        private string GetCellValue(Cell cell, string[] sharedStringCache)
        {
            if (cell.CellValue == null)
            {
                return string.Empty;
            }

            string value = cell.CellValue.Text ?? string.Empty;

            // 使用缓存的共享字符串
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringCache.Length > 0)
            {
                if (int.TryParse(value, out int index) && index >= 0 && index < sharedStringCache.Length)
                {
                    return sharedStringCache[index];
                }
                return string.Empty;
            }

            // 处理数字格式
            if (cell.DataType == null || cell.DataType.Value == CellValues.Number)
            {
                // 移除尾部的.0
                if (value.EndsWith(".0"))
                {
                    return value[..^2];
                }
                // 处理科学计数法
                if (decimal.TryParse(value, System.Globalization.NumberStyles.Float, 
                    System.Globalization.CultureInfo.InvariantCulture, out decimal decimalValue))
                {
                    return decimalValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            return value;
        }

        private int GetColumnIndexFromReference(string? cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return 0;
            }

            // 从单元格引用中提取列字母（例如从"A1"中提取"A"）
            string columnReference = Regex.Match(cellReference, @"[A-Za-z]+").Value;
            
            // 将列字母转换为索引（A=0, B=1, ..., Z=25, AA=26, ...）
            int index = 0;
            for (int i = 0; i < columnReference.Length; i++)
            {
                index = index * 26 + (columnReference[i] - 'A' + 1);
            }
            return index - 1; // 转换为0基索引
        }

        private DataTable ReadCsvFile(string filePath, ImportSettings settings)
        {
            var dt = new DataTable();
            using var reader = new StreamReader(filePath);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

            // 读取列头
            csv.Read();
            csv.ReadHeader();
            foreach (string header in csv.HeaderRecord ?? Array.Empty<string>())
            {
                dt.Columns.Add(header.Trim());
            }

            // 读取数据
            while (csv.Read())
            {
                var row = dt.NewRow();
                bool hasData = false;

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    string value = csv.GetField(i) ?? "";
                    if (settings.TrimStrings)
                    {
                        value = value.Trim();
                    }
                    if (!string.IsNullOrEmpty(value))
                    {
                        hasData = true;
                    }
                    row[i] = value;
                }

                if (!settings.SkipEmptyRows || hasData)
                {
                    dt.Rows.Add(row);
                }
            }

            return dt;
        }

        private async void BtnImport_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text))
            {
                MessageBox.Show("请先选择要导入的文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbDatabase.SelectedItem == null)
            {
                MessageBox.Show("请先选择数据库", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using var settingsForm = new ImportSettingsForm(dbConfig.GetConnectionString());
                if (settingsForm.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                Cursor = Cursors.WaitCursor;
                var settings = settingsForm.Settings;

                // 创建取消令牌
                using var cts = new CancellationTokenSource();

                // 添加取消按钮
                var btnCancel = new Button
                {
                    Text = "取消导入",
                    Location = new System.Drawing.Point(btnImport.Right + 10, btnImport.Top),
                    Width = 100
                };
                btnCancel.Click += (s, e) => cts.Cancel();
                this.Controls.Add(btnCancel);
                btnImport.Enabled = false;

                try
                {
                    // 读取数据
                    DataTable data;
                    string extension = Path.GetExtension(txtFilePath.Text).ToLower();
                    if (extension == ".csv")
                    {
                        data = ReadCsvFile(txtFilePath.Text, settings);
                    }
                    else
                    {
                        data = await ReadExcelFileAsync(txtFilePath.Text, settings, cts.Token);
                    }

                    if (data.Rows.Count == 0)
                    {
                        MessageBox.Show("没有找到可导入的数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 导入数据
                    ImportData(data, settings);
                }
                finally
                {
                    this.Controls.Remove(btnCancel);
                    btnCancel.Dispose();
                    btnImport.Enabled = true;
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("导入操作已取消", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void ImportData(DataTable data, ImportSettings settings)
        {
            using var connection = new MySqlConnection(dbConfig.GetConnectionString());
            connection.Open();
            LogMessage($"开始导入数据到表 {settings.TableName}");

            // 检查表名是否合法
            if (string.IsNullOrWhiteSpace(settings.TableName))
            {
                throw new Exception("表名不能为空");
            }

            if (!Regex.IsMatch(settings.TableName, "^[a-zA-Z0-9_]+$"))
            {
                throw new Exception("表名只能包含字母、数字和下划线");
            }

            // 检查表是否存在
            bool tableExists = CheckTableExists(connection, settings.TableName);
            LogMessage($"检查表 {settings.TableName} 是否存在: {(tableExists ? "存在" : "不存在")}");

            // 声明表结构信息变量
            Dictionary<string, string> columnTypes;

            if (tableExists)
            {
                switch (settings.Mode)
                {
                    case ImportMode.ErrorIfExists:
                        LogMessage($"表 {settings.TableName} 已存在，根据设置终止导入", true);
                        throw new Exception($"表 {settings.TableName} 已存在");
                    case ImportMode.ClearAndImport:
                        LogMessage($"清空表 {settings.TableName} 中的数据");
                        ExecuteNonQuery(connection, $"TRUNCATE TABLE {settings.TableName}");
                        break;
                    case ImportMode.Append:
                        LogMessage("使用追加模式导入数据");
                        // 验证表结构
                        ValidateTableSchema(connection, settings.TableName, data);
                        break;
                }
                // 获取现有表的结构信息
                columnTypes = GetTableColumnTypes(connection, settings.TableName);
            }
            else if (settings.CreateTableIfNotExists)
            {
                try
                {
                    LogMessage($"创建新表 {settings.TableName}");
                    CreateTable(connection, settings.TableName, data);
                    // 获取新建表的结构信息
                    columnTypes = GetTableColumnTypes(connection, settings.TableName);
                    LogMessage("表创建成功");
                }
                catch (Exception ex)
                {
                    LogMessage($"创建表失败: {ex.Message}", true);
                    throw new Exception($"创建表 {settings.TableName} 失败: {ex.Message}");
                }
            }
            else
            {
                LogMessage($"表 {settings.TableName} 不存在且未设置自动创建", true);
                throw new Exception($"表 {settings.TableName} 不存在");
            }

            // 验证列数是否匹配
            if (data.Columns.Count != columnTypes.Count)
            {
                throw new Exception($"数据列数({data.Columns.Count})与表结构列数({columnTypes.Count})不匹配");
            }

            // 验证列名是否匹配
            foreach (DataColumn column in data.Columns)
            {
                if (!columnTypes.ContainsKey(column.ColumnName))
                {
                    throw new Exception($"数据列 '{column.ColumnName}' 在表结构中不存在");
                }
            }

            var columns = string.Join(", ", data.Columns.Cast<DataColumn>().Select(c => $"`{c.ColumnName}`"));
            var values = string.Join(", ", data.Columns.Cast<DataColumn>().Select(c => "?"));
            var sql = $"INSERT INTO {settings.TableName} ({columns}) VALUES ({values})";
            LogMessage($"SQL语句: {sql}");

            MySqlTransaction? transaction = null;
            try
            {
                transaction = connection.BeginTransaction();
                using var cmd = new MySqlCommand(sql, connection, transaction);
                var parameters = data.Columns.Cast<DataColumn>()
                    .Select(c => new MySqlParameter())
                    .ToArray();
                cmd.Parameters.AddRange(parameters);

                int batchCount = 0;
                int totalCount = 0;
                int totalRows = data.Rows.Count;
                int errorCount = 0;
                const int maxErrors = 10; // 最大错误数，超过此数则停止导入

                foreach (DataRow row in data.Rows)
                {
                    try
                    {
                        for (int i = 0; i < parameters.Length; i++)
                        {
                            var columnName = data.Columns[i].ColumnName;
                            var value = row[i];
                            
                            // 处理空值
                            if (value == DBNull.Value || value == null || (value is string str && string.IsNullOrWhiteSpace(str)))
                            {
                                parameters[i].Value = DBNull.Value;
                                continue;
                            }

                            // 根据MySQL列类型处理数据
                            if (columnTypes.TryGetValue(columnName, out var columnType))
                            {
                                var upperColumnType = columnType.ToUpper();
                                try
                                {
                                    if (upperColumnType.Contains("DECIMAL") || upperColumnType.Contains("DOUBLE") || 
                                        upperColumnType.Contains("FLOAT") || upperColumnType == "REAL")
                                    {
                                        // 处理数值类型
                                        if (decimal.TryParse(value.ToString(), out decimal decimalValue))
                                        {
                                            parameters[i].Value = decimalValue;
                                        }
                                        else
                                        {
                                            parameters[i].Value = DBNull.Value;
                                            LogMessage($"警告: 行 {totalCount + 1}, 列 '{columnName}' 的值 '{value}' 无法转换为数值，已设为NULL");
                                        }
                                    }
                                    else if (upperColumnType.Contains("INT"))
                                    {
                                        // 处理整数类型
                                        if (long.TryParse(value.ToString(), out long intValue))
                                        {
                                            parameters[i].Value = intValue;
                                        }
                                        else
                                        {
                                            parameters[i].Value = DBNull.Value;
                                            LogMessage($"警告: 行 {totalCount + 1}, 列 '{columnName}' 的值 '{value}' 无法转换为整数，已设为NULL");
                                        }
                                    }
                                    else if (upperColumnType.Contains("DATE") || upperColumnType.Contains("TIME"))
                                    {
                                        // 处理日期时间类型
                                        if (DateTime.TryParse(value.ToString(), out DateTime dateValue))
                                        {
                                            parameters[i].Value = dateValue;
                                        }
                                        else
                                        {
                                            parameters[i].Value = DBNull.Value;
                                            LogMessage($"警告: 行 {totalCount + 1}, 列 '{columnName}' 的值 '{value}' 无法转换为日期，已设为NULL");
                                        }
                                    }
                                    else
                                    {
                                        // 其他类型按字符串处理
                                        parameters[i].Value = value;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"警告: 行 {totalCount + 1}, 列 '{columnName}' 的值 '{value}' 转换失败: {ex.Message}", true);
                                    parameters[i].Value = DBNull.Value;
                                }
                            }
                            else
                            {
                                // 如果找不到列类型信息，按原值处理
                                parameters[i].Value = value;
                            }
                        }

                        cmd.ExecuteNonQuery();
                        batchCount++;
                        totalCount++;

                        // 更新进度
                        if (totalCount % 2000 == 0 || totalCount == totalRows)
                        {
                            int progressValue = (int)((double)totalCount / totalRows * 100);
                            UpdateProgress(progressValue, $"已处理 {totalCount}/{totalRows} 条记录");
                        }

                        if (batchCount >= settings.BatchSize)
                        {
                            transaction.Commit();
                            transaction.Dispose();
                            transaction = connection.BeginTransaction();
                            cmd.Transaction = transaction;
                            batchCount = 0;
                        }
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        LogMessage($"导入第 {totalCount + 1} 行数据时出错: {ex.Message}", true);
                        
                        if (errorCount >= maxErrors)
                        {
                            throw new Exception($"导入过程中出现了 {errorCount} 个错误，已超过最大错误数限制，导入终止");
                        }
                        
                        // 继续处理下一行
                        continue;
                    }
                }

                if (transaction != null && batchCount > 0)
                {
                    transaction.Commit();
                    LogMessage($"已提交最后 {batchCount} 条记录");
                }

                string resultMessage = errorCount > 0 
                    ? $"导入完成，成功导入 {totalCount} 条数据，失败 {errorCount} 条"
                    : $"成功导入 {totalCount} 条数据";
                
                LogMessage($"数据导入完成，{resultMessage}");
                MessageBox.Show(resultMessage, errorCount > 0 ? "导入完成(有错误)" : "成功", 
                    MessageBoxButtons.OK, errorCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LogMessage($"导入数据时发生错误: {ex.Message}", true);
                LogMessage($"详细错误: {ex}", true);
                transaction?.Rollback();
                throw;
            }
            finally
            {
                transaction?.Dispose();
                UpdateProgress(0);
            }
        }

        private Dictionary<string, string> GetTableColumnTypes(MySqlConnection connection, string tableName)
        {
            var columnTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using var cmd = new MySqlCommand($"SHOW COLUMNS FROM {tableName}", connection);
            using var reader = cmd.ExecuteReader();
            
            while (reader.Read())
            {
                string columnName = reader.GetString("Field");
                string columnType = reader.GetString("Type");
                columnTypes[columnName] = columnType;
            }
            
            return columnTypes;
        }

        private bool CheckTableExists(MySqlConnection connection, string tableName)
        {
            using var cmd = new MySqlCommand(
                "SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = @database AND table_name = @tableName",
                connection);
            cmd.Parameters.AddWithValue("@database", connection.Database);
            cmd.Parameters.AddWithValue("@tableName", tableName);
            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }

        private void CreateTable(MySqlConnection connection, string tableName, DataTable data)
        {
            // 分析每列的最大长度和数据类型
            var columnDefinitions = new List<string>();
            foreach (DataColumn column in data.Columns)
            {
                int maxLength = 1; // 默认最小长度为1
                bool hasNumericData = true;
                bool hasDecimalPoint = false;
                bool hasDateData = true;
                bool allEmpty = true;
                
                // 分析列数据
                foreach (DataRow row in data.Rows)
                {
                    var value = row[column];
                    if (value == DBNull.Value || value == null)
                    {
                        continue;
                    }

                    string strValue = value.ToString() ?? "";
                    if (!string.IsNullOrEmpty(strValue))
                    {
                        allEmpty = false;
                        // 检查长度
                        maxLength = Math.Max(maxLength, strValue.Length);
                        
                        // 检查是否为数字
                        if (hasNumericData)
                        {
                            hasNumericData = decimal.TryParse(strValue, out decimal numericValue);
                            if (hasNumericData && strValue.Contains('.'))
                            {
                                hasDecimalPoint = true;
                            }
                        }
                        
                        // 检查是否为日期
                        if (hasDateData)
                        {
                            hasDateData = DateTime.TryParse(strValue, out _);
                        }
                    }
                }
                
                // 确定列类型
                string columnType;
                if (allEmpty)
                {
                    // 如果列全是空值，使用默认的VARCHAR(255)
                    columnType = "VARCHAR(255)";
                    LogMessage($"列 {column.ColumnName} 全为空值，使用默认类型 VARCHAR(255)");
                }
                else if (hasDateData && maxLength <= 30)
                {
                    columnType = "DATETIME NULL";
                }
                else if (hasNumericData)
                {
                    if (hasDecimalPoint)
                    {
                        columnType = "DECIMAL(18,2) NULL";
                    }
                    else
                    {
                        columnType = (maxLength <= 9 ? "INT" : "BIGINT") + " NULL";
                    }
                }
                else
                {
                    // 对于文本类型，根据最大长度选择合适的类型
                    if (maxLength <= 255)
                    {
                        columnType = $"VARCHAR({maxLength}) NULL";
                    }
                    else if (maxLength <= 65535)
                    {
                        columnType = "TEXT NULL";
                    }
                    else if (maxLength <= 16777215)
                    {
                        columnType = "MEDIUMTEXT NULL";
                    }
                    else
                    {
                        columnType = "LONGTEXT NULL";
                    }
                }
                
                columnDefinitions.Add($"`{column.ColumnName}` {columnType}");
            }
            
            // 创建表
            var sql = $"CREATE TABLE {tableName} ({string.Join(", ", columnDefinitions)}) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci";
            LogMessage($"创建表SQL: {sql}");
            ExecuteNonQuery(connection, sql);
        }

        private void ValidateTableSchema(MySqlConnection connection, string tableName, DataTable data)
        {
            // 获取表结构
            using var cmd = new MySqlCommand($"SHOW COLUMNS FROM {tableName}", connection);
            using var reader = cmd.ExecuteReader();
            var tableColumns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var nullableColumns = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            
            while (reader.Read())
            {
                string columnName = reader.GetString("Field");
                string columnType = reader.GetString("Type");
                string isNullable = reader.GetString("Null");
                tableColumns[columnName] = columnType;
                nullableColumns[columnName] = isNullable.Equals("YES", StringComparison.OrdinalIgnoreCase);
            }
            reader.Close();

            // 检查每个列的类型是否兼容
            foreach (DataColumn column in data.Columns)
            {
                if (!tableColumns.ContainsKey(column.ColumnName))
                {
                    LogMessage($"警告: 表中不存在列 {column.ColumnName}", true);
                    throw new Exception($"表结构不匹配: 缺少列 {column.ColumnName}");
                }

                string existingType = tableColumns[column.ColumnName].ToUpper();
                bool isNullable = nullableColumns[column.ColumnName];

                // 检查是否需要允许NULL
                bool hasNullValues = data.AsEnumerable()
                    .Any(r => r[column] == DBNull.Value || r[column] == null);

                if (hasNullValues && !isNullable)
                {
                    LogMessage($"列 {column.ColumnName} 包含空值，需要修改为允许NULL");
                    string baseType = existingType;
                    ExecuteNonQuery(connection, $"ALTER TABLE {tableName} MODIFY COLUMN `{column.ColumnName}` {baseType} NULL");
                }

                // 检查文本类型的长度是否足够
                if (existingType.StartsWith("VARCHAR"))
                {
                    int currentLength = int.Parse(existingType.Split('(', ')')[1]);
                    int maxLength = 1; // 默认最小长度为1
                    
                    foreach (DataRow row in data.Rows)
                    {
                        if (row[column] != DBNull.Value && row[column] != null)
                        {
                            string value = row[column].ToString() ?? "";
                            if (!string.IsNullOrEmpty(value))
                            {
                                maxLength = Math.Max(maxLength, value.Length);
                            }
                        }
                    }

                    if (maxLength > currentLength)
                    {
                        if (maxLength <= 255)
                        {
                            LogMessage($"需要扩展列 {column.ColumnName} 的长度从 {currentLength} 到 {maxLength}");
                            ExecuteNonQuery(connection, $"ALTER TABLE {tableName} MODIFY COLUMN `{column.ColumnName}` VARCHAR({maxLength}) {(isNullable ? "NULL" : "NOT NULL")}");
                        }
                        else if (maxLength <= 65535)
                        {
                            LogMessage($"需要将列 {column.ColumnName} 转换为 TEXT 类型");
                            ExecuteNonQuery(connection, $"ALTER TABLE {tableName} MODIFY COLUMN `{column.ColumnName}` TEXT {(isNullable ? "NULL" : "NOT NULL")}");
                        }
                        else if (maxLength <= 16777215)
                        {
                            LogMessage($"需要将列 {column.ColumnName} 转换为 MEDIUMTEXT 类型");
                            ExecuteNonQuery(connection, $"ALTER TABLE {tableName} MODIFY COLUMN `{column.ColumnName}` MEDIUMTEXT {(isNullable ? "NULL" : "NOT NULL")}");
                        }
                        else
                        {
                            LogMessage($"需要将列 {column.ColumnName} 转换为 LONGTEXT 类型");
                            ExecuteNonQuery(connection, $"ALTER TABLE {tableName} MODIFY COLUMN `{column.ColumnName}` LONGTEXT {(isNullable ? "NULL" : "NOT NULL")}");
                        }
                    }
                }
            }
        }

        private void ExecuteNonQuery(MySqlConnection connection, string sql)
        {
            using var cmd = new MySqlCommand(sql, connection);
            cmd.ExecuteNonQuery();
        }

        // 日志记录方法
        private void LogMessage(string message, bool isError = false)
        {
            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new Action(() => LogMessage(message, isError)));
                return;
            }

            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string logMessage = $"[{timestamp}] {message}{Environment.NewLine}";

            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.SelectionLength = 0;

            if (isError)
            {
                txtLog.SelectionColor = Color.Red;
            }
            else
            {
                txtLog.SelectionColor = Color.Black;
            }

            txtLog.AppendText(logMessage);
            txtLog.ScrollToCaret();
        }

        // 更新进度条方法
        private void UpdateProgress(int value, string message = "")
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() => UpdateProgress(value, message)));
                return;
            }

            progressBar.Value = Math.Min(100, Math.Max(0, value));
            if (!string.IsNullOrEmpty(message))
            {
                LogMessage($"进度 {value}%: {message}");
            }
        }
    }
} 