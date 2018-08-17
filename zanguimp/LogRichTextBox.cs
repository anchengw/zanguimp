using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RichTextLog
{
    public class LogRichTextBox
    {
        #region 日志记录、支持其他线程访问 
        public static RichTextBox richTextBoxRemote = null;
        private delegate void LogAppendDelegate(Color color, string text);
        /// <summary> 
        /// 追加显示文本 
        /// </summary> 
        /// <param name="color">文本颜色</param> 
        /// <param name="text">显示文本</param> 
        private static void LogAppend(Color color, string text)
        {
            if (richTextBoxRemote.Lines.Length > 0)
            {
                richTextBoxRemote.AppendText("\n");
                richTextBoxRemote.AppendText("\n");
                richTextBoxRemote.AppendText("❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆❆");
                richTextBoxRemote.AppendText("\n");
                richTextBoxRemote.AppendText("\n");
            }
            richTextBoxRemote.SelectionColor = color;
            richTextBoxRemote.AppendText(text);

            //richTextBoxRemote.SelectionStart = richTextBoxRemote.TextLength;
            richTextBoxRemote.ScrollToCaret();//滚动到控件光标处 
        }
        /// <summary> 
        /// 显示错误日志 
        /// </summary> 
        /// <param name="text"></param> 
        private static void LogError(string text)
        {
            LogAppendDelegate la = new LogAppendDelegate(LogAppend);
            richTextBoxRemote.Invoke(la, Color.Red, DateTime.Now.ToString("HH:mm:ss ") + text);
        }
        /// <summary> 
        /// 显示警告信息 
        /// </summary> 
        /// <param name="text"></param> 
        private static void LogWarning(string text)
        {
            LogAppendDelegate la = new LogAppendDelegate(LogAppend);
            richTextBoxRemote.Invoke(la, Color.Violet, DateTime.Now.ToString("HH:mm:ss ") + text);
        }
        /// <summary> 
        /// 显示信息 
        /// </summary> 
        /// <param name="text"></param> 
        private static void LogMessage(string text)
        {
            LogAppendDelegate la = new LogAppendDelegate(LogAppend);
            richTextBoxRemote.Invoke(la, Color.Black, DateTime.Now.ToString("HH:mm:ss ") + text);
        }
        /// <summary> 
        /// 清除日志 
        /// </summary> 
        /// <param name="text"></param> 
        public static void LogClear()
        {
            Action action = delegate () { richTextBoxRemote.Text = ""; };
            action();
        }
        public static void logMesg(string mesg,int type = 0)
        {
            if (string.IsNullOrEmpty(mesg) || richTextBoxRemote == null)
                return;
            richTextBoxRemote.ReadOnly = true;
            switch (type)
            {
                case 0:
                    LogMessage(mesg);
                    break;
                case 1:
                    LogWarning(mesg);
                    break;
                case 2:
                    LogError(mesg);
                    break;
                default:
                    LogMessage(mesg);
                    break;
            }
        }
        #endregion
    }
}
