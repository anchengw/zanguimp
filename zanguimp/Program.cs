﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace zanguimp
{
    static class Program
    {
        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(System.IntPtr hWnd, int cmdShow);
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(System.IntPtr hWnd);
        private static void HandleRunningInstance(Process instance)
        {
            MessageBox.Show("已经在运行！", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ShowWindowAsync(instance.MainWindowHandle, 1);  //调用api函数，正常显示窗口
            SetForegroundWindow(instance.MainWindowHandle); //将窗口放置最前端
        }
        /// <summary>
        /// 获取进程
        /// </summary>
        /// <returns></returns>
        private static Process RunningInstance()
        {
            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(current.ProcessName);

            foreach (Process process in processes)
            {
                if (process.Id != current.Id)
                {
                    if (process.MainModule.FileName == current.MainModule.FileName)
                    {
                        return process;
                    }
                }
            }
            return null;
        }
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool runone;
            System.Threading.Mutex run = new System.Threading.Mutex(true, "single_zanguimp", out runone);
            
            if (runone)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            else
            {
                Process instance = RunningInstance();
                HandleRunningInstance(instance);
            }
        }
    }
}
