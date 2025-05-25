using System;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Data;

namespace WindowsFormsApp
{
    public class ImportSettingsForm : Form
    {
        private readonly Label lblTableName = new();
        private readonly ComboBox txtTableName = new();
        private readonly Label lblExistingTables = new();
        private readonly ListBox lstExistingTables = new();
        private readonly Label lblImportMode = new();
        private readonly ComboBox cmbImportMode = new();
        private readonly CheckBox chkCreateTable = new();
        private readonly CheckBox chkTrimStrings = new();
        private readonly CheckBox chkSkipEmptyRows = new();
        private readonly Label lblBatchSize = new();
        private readonly NumericUpDown numBatchSize = new();
        private readonly Button btnOK = new();
        private readonly Button btnCancel = new();

        public ImportSettings Settings { get; private set; }
        private readonly string connectionString;

        public ImportSettingsForm(string connectionString)
        {
            this.connectionString = connectionString;
            Settings = new ImportSettings();
            InitializeComponents();
            LoadExistingTables();
        }

        private void InitializeComponents()
        {
            this.Text = "导入设置";
            this.Size = new System.Drawing.Size(600, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Padding = new Padding(10);

            // 表名
            lblTableName.Text = "表名：";
            lblTableName.Location = new System.Drawing.Point(20, 20);
            lblTableName.AutoSize = true;

            txtTableName.Location = new System.Drawing.Point(120, 17);
            txtTableName.Width = 200;
            txtTableName.AutoCompleteMode = AutoCompleteMode.Suggest;
            txtTableName.AutoCompleteSource = AutoCompleteSource.ListItems;

            // 现有表列表
            lblExistingTables.Text = "现有表：";
            lblExistingTables.Location = new System.Drawing.Point(340, 20);
            lblExistingTables.AutoSize = true;

            lstExistingTables.Location = new System.Drawing.Point(340, 40);
            lstExistingTables.Size = new System.Drawing.Size(200, 200);
            lstExistingTables.SelectionMode = SelectionMode.One;
            lstExistingTables.SelectedIndexChanged += LstExistingTables_SelectedIndexChanged;

            // 导入模式
            lblImportMode.Text = "导入模式：";
            lblImportMode.Location = new System.Drawing.Point(20, 60);
            lblImportMode.AutoSize = true;

            cmbImportMode.Location = new System.Drawing.Point(120, 57);
            cmbImportMode.Width = 200;
            cmbImportMode.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbImportMode.Items.AddRange(new object[] {
                "追加到现有数据",
                "清空表后导入",
                "如果表存在则报错"
            });
            cmbImportMode.SelectedIndex = 0;

            // 选项
            chkCreateTable.Text = "如果表不存在则创建";
            chkCreateTable.Location = new System.Drawing.Point(120, 100);
            chkCreateTable.AutoSize = true;
            chkCreateTable.Checked = true;

            chkTrimStrings.Text = "删除字符串前后空格";
            chkTrimStrings.Location = new System.Drawing.Point(120, 130);
            chkTrimStrings.AutoSize = true;
            chkTrimStrings.Checked = true;

            chkSkipEmptyRows.Text = "跳过空行";
            chkSkipEmptyRows.Location = new System.Drawing.Point(120, 160);
            chkSkipEmptyRows.AutoSize = true;
            chkSkipEmptyRows.Checked = true;

            // 批量大小
            lblBatchSize.Text = "批量大小：";
            lblBatchSize.Location = new System.Drawing.Point(20, 200);
            lblBatchSize.AutoSize = true;

            numBatchSize.Location = new System.Drawing.Point(120, 197);
            numBatchSize.Width = 100;
            numBatchSize.Minimum = 1;
            numBatchSize.Maximum = 10000;
            numBatchSize.Value = 1000;

            // 按钮
            btnOK.Text = "确定";
            btnOK.Location = new System.Drawing.Point(120, 400);
            btnOK.Width = 90;
            btnOK.Click += BtnOK_Click;

            btnCancel.Text = "取消";
            btnCancel.Location = new System.Drawing.Point(220, 400);
            btnCancel.Width = 90;
            btnCancel.Click += BtnCancel_Click;

            // 添加控件
            this.Controls.AddRange(new Control[] {
                lblTableName, txtTableName,
                lblExistingTables, lstExistingTables,
                lblImportMode, cmbImportMode,
                chkCreateTable, chkTrimStrings, chkSkipEmptyRows,
                lblBatchSize, numBatchSize,
                btnOK, btnCancel
            });
        }

        private void LoadExistingTables()
        {
            try
            {
                using var connection = new MySqlConnection(connectionString);
                connection.Open();

                // 获取所有表名
                var dt = connection.GetSchema("Tables");
                var tables = new List<string>();
                foreach (DataRow row in dt.Rows)
                {
                    string? tableName = row["TABLE_NAME"]?.ToString();
                    if (!string.IsNullOrEmpty(tableName))
                    {
                        // 排除系统表
                        if (!tableName.Equals("information_schema", StringComparison.OrdinalIgnoreCase) &&
                            !tableName.Equals("mysql", StringComparison.OrdinalIgnoreCase) &&
                            !tableName.Equals("performance_schema", StringComparison.OrdinalIgnoreCase) &&
                            !tableName.Equals("sys", StringComparison.OrdinalIgnoreCase))
                        {
                            tables.Add(tableName);
                        }
                    }
                }

                // 按字母顺序排序
                tables.Sort();

                // 更新列表框和自动完成
                lstExistingTables.Items.Clear();
                txtTableName.Items.Clear();

                if (tables.Count > 0)
                {
                    foreach (var table in tables)
                    {
                        lstExistingTables.Items.Add(table);
                        txtTableName.Items.Add(table);
                    }
                }
                else
                {
                    lstExistingTables.Items.Add("(当前数据库没有表)");
                }
            }
            catch (Exception ex)
            {
                lstExistingTables.Items.Clear();
                lstExistingTables.Items.Add("(加载表列表失败)");
                MessageBox.Show($"加载表列表失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LstExistingTables_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (lstExistingTables.SelectedItem != null && 
                lstExistingTables.SelectedItem.ToString() != "(当前数据库没有表)")
            {
                txtTableName.Text = lstExistingTables.SelectedItem.ToString();
            }
        }

        private void BtnOK_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTableName.Text))
            {
                MessageBox.Show("请输入表名", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Settings.TableName = txtTableName.Text.Trim();
            Settings.Mode = cmbImportMode.SelectedIndex switch
            {
                0 => ImportMode.Append,
                1 => ImportMode.ClearAndImport,
                2 => ImportMode.ErrorIfExists,
                _ => ImportMode.Append
            };
            Settings.CreateTableIfNotExists = chkCreateTable.Checked;
            Settings.TrimStrings = chkTrimStrings.Checked;
            Settings.SkipEmptyRows = chkSkipEmptyRows.Checked;
            Settings.BatchSize = (int)numBatchSize.Value;

            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancel_Click(object? sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
} 