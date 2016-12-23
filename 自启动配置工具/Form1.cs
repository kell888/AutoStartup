using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace 自启动配置工具
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 读取注册表
        /// </summary>
        /// <returns></returns>
        public static string ReadRegistry(string name)
        {
            try
            {
                RegistryKey rk = Registry.LocalMachine;
                RegistryKey softWare = rk.OpenSubKey("Software");
                RegistryKey microsoft = softWare.OpenSubKey("Microsoft");
                RegistryKey windows = microsoft.OpenSubKey("Windows");
                RegistryKey current = windows.OpenSubKey("CurrentVersion");
                RegistryKey run = current.OpenSubKey(@"Run", true);
                return run.GetValue(name, "", RegistryValueOptions.None).ToString();
            }
            catch// (Exception ex)
            {
                return "";
            }
        }

        /// <summary>
        /// 设置开机自启动-写入注册表
        /// </summary>
        /// <param name="exePath">可执行文件的完整路径</param>
        /// <returns></returns>
        public static bool SetSelfStarting(string exePath)
        {
            try
            {
                string name = Path.GetFileName(exePath);
                RegistryKey rk = Registry.LocalMachine;
                RegistryKey softWare = rk.OpenSubKey("SOFTWARE");
                RegistryKey microsoft = softWare.OpenSubKey("Microsoft");
                RegistryKey windows = microsoft.OpenSubKey("Windows");
                RegistryKey current = windows.OpenSubKey("CurrentVersion");
                RegistryKey run = current.OpenSubKey(@"Run", true);//这里必须加true就是得到写入权限 
                run.SetValue(name, exePath);
                run.Close();
                return true;
            }
            catch// (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// 取消开机自启-删除注册表
        /// </summary>
        /// <param name="exePath">可执行文件的完整路径</param>
        /// <returns></returns>
        public static bool CancelSelfStarting(string exePath)
        {
            try
            {
                string name = Path.GetFileName(exePath);
                RegistryKey rk = Registry.LocalMachine;
                RegistryKey softWare = rk.OpenSubKey("Software");
                RegistryKey microsoft = softWare.OpenSubKey("Microsoft");
                RegistryKey windows = microsoft.OpenSubKey("Windows");
                RegistryKey current = windows.OpenSubKey("CurrentVersion");
                RegistryKey run = current.OpenSubKey(@"Run", true);
                run.DeleteValue(name);//删除注册表的值
                run.Close();
                return true;
            }
            catch// (Exception ex)
            {
                return false;
            }
        }

        private void AutoStartup(string executablePath, bool flag)
        {
            string filename = Path.GetFileName(executablePath);
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string path = Environment.GetFolderPath(Environment.SpecialFolder.System).Substring(0, 1) + @":\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup";
            if (!Directory.Exists(path))
                path = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string file = path + "\\" + filename + ".lnk";
            if (flag)
            {
                if (!IsAutoStartupAllUsers(filename))
                {
                    try
                    {
                        CreateShortcut(executablePath, path);
                    }
                    catch (Exception e)
                    {
                        if (!IsAutoStartupCurrentUser(filename))
                            CreateShortcut(executablePath, Environment.GetFolderPath(Environment.SpecialFolder.Startup));
                    }
                }
                else if (!IsAutoStartupCurrentUser(filename))
                {
                    CreateShortcut(executablePath, Environment.GetFolderPath(Environment.SpecialFolder.Startup));
                }
            }
            else
            {
                if (IsAutoStartupAllUsers(filename))
                {
                    File.Delete(file);
                }
                if (IsAutoStartupCurrentUser(filename))
                {
                    string file2 = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\" + filename + ".lnk";
                    File.Delete(file2);
                }
            }
        }

        public void SetAutoStartupAllUsers(string executablePath)
        {
            string filename = Path.GetFileName(executablePath);
            string path = Environment.GetFolderPath(Environment.SpecialFolder.System).Substring(0, 1) + @":\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup";
            if (!Directory.Exists(path))
                path = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string file = path + "\\" + filename + ".lnk";
            if (!IsAutoStartupAllUsers(filename))
            {
                try
                {
                    CreateShortcut(executablePath, path);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }

        public void SetAutoStartupCurrentUser(string executablePath)
        {
            string filename = Path.GetFileName(executablePath);
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string file = path + "\\" + filename + ".lnk";
            if (!IsAutoStartupAllUsers(filename))
            {
                try
                {
                    CreateShortcut(executablePath, path);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }

        public void CancelAutoStartupAllUsers(string executablePath)
        {
            string filename = Path.GetFileName(executablePath);
            string path = Environment.GetFolderPath(Environment.SpecialFolder.System).Substring(0, 1) + @":\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup";
            string file = path + "\\" + filename + ".lnk";
            if (IsAutoStartupAllUsers(filename))
            {
                File.Delete(file);
            }
        }

        public void CancelAutoStartupCurrentUser(string executablePath)
        {
            string filename = Path.GetFileName(executablePath);
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string file = path + "\\" + filename + ".lnk";
            if (IsAutoStartupCurrentUser(filename))
            {
                File.Delete(file);
            }
        }

        private static void CreateShortcut(string executablePath, string path)
        {
            // 声明操作对象
            IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShellClass();
            string filename = Path.GetFileName(executablePath);
            string file = path + "\\" + filename + ".lnk";
            // 创建一个快捷方式
            IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(file);
            // 关联的程序
            shortcut.TargetPath = executablePath;
            // 参数
            //shortcut.Arguments = "";
            // 快捷方式描述，鼠标放到快捷方式上会显示出来哦
            shortcut.Description = filename + "应用程序";
            // 全局热键
            //shortcut.Hotkey = "CTRL+SHIFT+N";
            // 设置快捷方式的图标，这里是取程序图标，如果希望指定一个ico文件，那么请写路径。
            //shortcut.IconLocation = "notepad.exe, 0";
            // 保存，创建就成功了。
            shortcut.Save();
        }

        private bool IsAutoStartupAllUsers(string filename)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.System).Substring(0, 1) + @":\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup";
            if (Directory.Exists(path))
            {
                string lnk = path + "\\" + filename + ".lnk";
                string[] files = Directory.GetFiles(path, "*.lnk", SearchOption.TopDirectoryOnly);
                foreach (string file in files)
                {
                    if (file.Equals(lnk, StringComparison.InvariantCultureIgnoreCase))
                        return true;
                }
            }
            return false;
        }

        private bool IsAutoStartupCurrentUser(string filename)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string lnk = path + "\\" + filename + ".lnk";
            string[] files = Directory.GetFiles(path, "*.lnk", SearchOption.TopDirectoryOnly);
            foreach (string file in files)
            {
                if (file.Equals(lnk, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
            return false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
            openFileDialog1.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string file = textBox1.Text.Trim();
            if (File.Exists(file))
            {
                if (radioButton1.Checked)
                    SetSelfStarting(file);
                else if (radioButton2.Checked)
                    SetAutoStartupAllUsers(file);
                else if (radioButton3.Checked)
                    SetAutoStartupCurrentUser(file);
                MessageBox.Show("设置成功");
            }
            else
            {
                MessageBox.Show("指定的文件不存在！请确认。");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string file = textBox1.Text.Trim();
            if (radioButton1.Checked)
                CancelSelfStarting(file);
            else if (radioButton2.Checked)
                CancelAutoStartupAllUsers(file);
            else if (radioButton3.Checked)
                CancelAutoStartupCurrentUser(file);
            MessageBox.Show("取消成功");
        }
    }
}
