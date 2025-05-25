using System;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace WindowsFormsApp
{
    public class DatabaseSettingsForm : Form
    {
        private readonly Label lblServer = new();
        private readonly TextBox txtServer = new();
        private readonly Label lblPort = new();
        private readonly NumericUpDown numPort = new();
        private readonly Label lblUsername = new();
        private readonly TextBox txtUsername = new();
        private readonly Label lblPassword = new();
        private readonly TextBox txtPassword = new();
        private readonly Button btnSave = new();
        private readonly Button btnTest = new();
        private readonly Button btnCancel = new();
        private readonly CheckBox chkShowPassword = new();

        private readonly DatabaseConfig dbConfig;

        public DatabaseSettingsForm(DatabaseConfig config)
        {
            dbConfig = config;
            InitializeComponents();
            LoadSettings();
        }

        private void InitializeComponents()
        {
            this.Text = "数据库连接设置";
            this.Size = new System.Drawing.Size(400, 300);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Padding = new Padding(10);

            // 服务器
            lblServer.Text = "服务器：";
            lblServer.Location = new System.Drawing.Point(20, 20);
            lblServer.AutoSize = true;

            txtServer.Location = new System.Drawing.Point(120, 17);
            txtServer.Width = 200;

            // 端口
            lblPort.Text = "端口：";
            lblPort.Location = new System.Drawing.Point(20, 60);
            lblPort.AutoSize = true;

            numPort.Location = new System.Drawing.Point(120, 57);
            numPort.Width = 100;
            numPort.Minimum = 1;
            numPort.Maximum = 65535;
            numPort.Value = 3306;

            // 用户名
            lblUsername.Text = "用户名：";
            lblUsername.Location = new System.Drawing.Point(20, 100);
            lblUsername.AutoSize = true;

            txtUsername.Location = new System.Drawing.Point(120, 97);
            txtUsername.Width = 200;

            // 密码
            lblPassword.Text = "密码：";
            lblPassword.Location = new System.Drawing.Point(20, 140);
            lblPassword.AutoSize = true;

            txtPassword.Location = new System.Drawing.Point(120, 137);
            txtPassword.Width = 200;
            txtPassword.UseSystemPasswordChar = true;

            // 显示密码复选框
            chkShowPassword.Text = "显示密码";
            chkShowPassword.Location = new System.Drawing.Point(120, 170);
            chkShowPassword.AutoSize = true;
            chkShowPassword.CheckedChanged += ChkShowPassword_CheckedChanged;

            // 按钮
            btnTest.Text = "测试连接";
            btnTest.Location = new System.Drawing.Point(20, 210);
            btnTest.Width = 90;
            btnTest.Click += BtnTest_Click;

            btnSave.Text = "保存";
            btnSave.Location = new System.Drawing.Point(120, 210);
            btnSave.Width = 90;
            btnSave.Click += BtnSave_Click;

            btnCancel.Text = "取消";
            btnCancel.Location = new System.Drawing.Point(220, 210);
            btnCancel.Width = 90;
            btnCancel.Click += BtnCancel_Click;

            // 添加控件
            this.Controls.AddRange(new Control[] {
                lblServer, txtServer,
                lblPort, numPort,
                lblUsername, txtUsername,
                lblPassword, txtPassword,
                chkShowPassword,
                btnTest, btnSave, btnCancel
            });
        }

        private void LoadSettings()
        {
            txtServer.Text = dbConfig.Server;
            numPort.Value = dbConfig.Port;
            txtUsername.Text = dbConfig.Username;
            txtPassword.Text = dbConfig.Password;
        }

        private void ChkShowPassword_CheckedChanged(object? sender, EventArgs e)
        {
            txtPassword.UseSystemPasswordChar = !chkShowPassword.Checked;
        }

        private void BtnTest_Click(object? sender, EventArgs e)
        {
            try
            {
                var tempConfig = new DatabaseConfig
                {
                    Server = txtServer.Text,
                    Port = (int)numPort.Value,
                    Username = txtUsername.Text,
                    Password = txtPassword.Text
                };

                using var connection = new MySqlConnection(tempConfig.GetConnectionString());
                connection.Open();
                MessageBox.Show("数据库连接成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (MySqlException ex)
            {
                string errorMessage = ex.Number switch
                {
                    1042 => "无法连接到数据库服务器，请检查服务器地址和端口是否正确。",
                    1045 => "用户名或密码错误。",
                    _ => $"连接失败：{ex.Message}"
                };
                MessageBox.Show(errorMessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生未知错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            dbConfig.Server = txtServer.Text;
            dbConfig.Port = (int)numPort.Value;
            dbConfig.Username = txtUsername.Text;
            dbConfig.Password = txtPassword.Text;
            dbConfig.Save();
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