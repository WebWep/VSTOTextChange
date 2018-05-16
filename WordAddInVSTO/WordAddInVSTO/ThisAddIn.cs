using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using System.Net;
using System.Configuration;
using Microsoft.Office.Tools;
using System.Text.RegularExpressions;

namespace WordAddInVSTO
{
    public partial class ThisAddIn
    {
        // 定义2个委托变量, 挂钩子回调方法
        private SafeNativeMethods.HookProc _mouseProc;
        private SafeNativeMethods.HookProc _keyboardProc;

        private IntPtr _hookIdMouse;
        private IntPtr _hookIdKeyboard;

        /// <summary>
        /// 缓存当前监听键盘按键得到句子，用来处理一些细节，如重复语句等。
        /// 目前处理的不是很完善，具体根据需求处理。
        /// </summary>
        private string cacheUpdateText = string.Empty;

        /// <summary>
        /// 当前插件的TaskPanel
        /// </summary>
        public CustomTaskPane CurCustomTaskPane = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _mouseProc = MouseHookCallback;
            _keyboardProc = KeyboardHookCallback;


            ////初始化 TaskPanel
            initTaskPanel();
        }

        private void initTaskPanel()
        {
            TaskPanel tp = new TaskPanel();
            this.CurCustomTaskPane = this.CustomTaskPanes.Add(tp, "TextChangeDemo");
            this.CurCustomTaskPane.Width = 400;
            this.CurCustomTaskPane.Visible = false;

            this.CurCustomTaskPane.VisibleChanged += _CustomTaskPane_VisibleChanged;
        }

        private void _CustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            //// 功能区开关和 task panel同步
            Globals.Ribbons.MenuRibbon.detectTBtn.Checked = this.CurCustomTaskPane.Visible;
            //// 清空当前监听文本
            TaskPanel.DetectTextLabel.Text = "";
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            UnhookWindowsHooks();
        }

        internal void SetWindowsHooks()
        {
            uint threadId = (uint)SafeNativeMethods.GetCurrentThreadId();

            //// 挂鼠标钩子
            _hookIdMouse =
                SafeNativeMethods.SetWindowsHookEx(
                    (int)SafeNativeMethods.HookType.WH_MOUSE,
                    _mouseProc,
                    IntPtr.Zero,
                    threadId);

            //// 挂键盘钩子
            _hookIdKeyboard =
                SafeNativeMethods.SetWindowsHookEx(
                   (int)SafeNativeMethods.HookType.WH_KEYBOARD,
                    _keyboardProc,
                    IntPtr.Zero,
                    threadId);
        }

        internal void UnhookWindowsHooks()
        {
            SafeNativeMethods.UnhookWindowsHookEx(_hookIdKeyboard);
            SafeNativeMethods.UnhookWindowsHookEx(_hookIdMouse);
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var mouseHookStruct =
                    (SafeNativeMethods.MouseHookStructEx)
                        Marshal.PtrToStructure(lParam, typeof(SafeNativeMethods.MouseHookStructEx));

                // 处理鼠标按键信息
                var message = (SafeNativeMethods.WindowMessages)wParam;
            }
            return SafeNativeMethods.CallNextHookEx(
                _hookIdKeyboard,
                nCode,
                wParam,
                lParam);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            //// nCode >= 0 处理按键信息
            //// nCode == 0 wParam 和 lParam 包含按键信息
            //// nCode > 0 wParam 和 lParam 包含按键信息, 并且处理消息队列
            if (nCode == 0)
            {
                // 处理按键信息
                try
                {
                    Int64 _lParam = lParam.ToInt64();
                    byte[] _byteArray = BitConverter.GetBytes(_lParam);

                    ////0 - 15 RepeatCount, KEYUP 和 SYSKEYUP RepeatCount 始终是 1
                    int _repeatCount = BitConverter.ToInt16(_byteArray, 0);

                    //// 30 , 按键状态发生改变前的状态
                    int _previousKeyStateFlag = ((_byteArray[3] << 1) & 128) >> 7;
                    //// 31 ,  按压 或者 释放 按键状态
                    //// 0：KEYDOWN 或 SYSKEYDOWN
                    //// 1: KEYUP 或者 SYSKEYUP
                    int _transitionStateFlag = (_byteArray[3] & 128) >> 7;


                    //// 3个值都为1，判定为 KEYUP 或者 SYSKEYUP
                    //// 标识当前按键结束
                    if (_repeatCount == 1 && _previousKeyStateFlag == 1 && _transitionStateFlag == 1)
                    {
                        //// 如果按键是打字常用按键，则处理语句。
                        bool _isTextInput = isTextInput(wParam.ToInt64());
                        if (_isTextInput)
                        {
                            GetUpdateSentenceText();
                        }                      
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("监听键盘按键, 处理按键消息出错：" + ex.Message);
                }
            }

            return SafeNativeMethods.CallNextHookEx(
                _hookIdKeyboard,
                nCode,
                wParam,
                lParam);
        }

        /// <summary>
        /// 判断当前键值是否正常输入文本键
        /// </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        private bool isTextInput(long v)
        {
            //// A-Z
            if (v >= (int)Keys.A && v <= (int)Keys.Z)
            {
                return true;
            }
            //// 0 ~ 9
            if (v >= (int)Keys.D0 && v <= (int)Keys.D9)
            {
                return true;
            }
            //// 数字键盘 0 ~ 9
            if (v >= (int)Keys.NumPad0 && v <= (int)Keys.NumPad9)
            {
                return true;
            }
            //// 句号、问号、分号等
            if (v >= (int)Keys.OemSemicolon && v <= (int)Keys.Oemtilde)
            {
                return true;
            }
            //// 引号，括号等
            if (v >= (int)Keys.OemOpenBrackets && v <= (int)Keys.OemQuotes)
            {
                return true;
            }
            //// Back 回退键 或 delete 键
            if (v == (int)Keys.Back || v == (int)Keys.Delete)
            {
                return true;
            }
            return false;
        }

        private async void GetUpdateSentenceText()
        {
            ////通过选取类型来获取光标位置
            Word.Selection currentSelection = Application.Selection;

            //// 暂存当前用户的 Overtype 设置
            bool userOvertype = Application.Options.Overtype;

            // 确保 Overtype false.
            if (Application.Options.Overtype)
            {
                Application.Options.Overtype = false;
            }

            // 判断当前选取是否是一个插入点(编辑区光标)
            if (currentSelection.Type == Word.WdSelectionType.wdSelectionIP)
            {
                // 当前选取的Range位置，即光标位置
                int _cur_start = currentSelection.Range.Start;
                int _cur_end = currentSelection.Range.End;

                //// 以当前光标位置作为Range起始区间
                Word.Range rng = this.Application.ActiveDocument.Range(_cur_start, _cur_end);
                ////通过扩展Range范围来获取更新后的句子
                rng.MoveStart(Word.WdUnits.wdSentence, -1);
                rng.MoveEnd(Word.WdUnits.wdSentence, 1);
                string _curRangeText = rng.Text;

               // Debug.WriteLine(_curRangeText);
                TaskPanel.DetectTextLabel.Text = _curRangeText;

                //// 正则匹配，如果是完整句子，发送到服务端。
                string _sentenceRegex = @"(\.|。|\?|？|\!|！|……)";
                if (Regex.IsMatch(_curRangeText, _sentenceRegex))
                {
                    //// 去掉重复句子
                    if (cacheUpdateText != _curRangeText)
                    {
                        string _jsonParam = "{\"sentence\":\"" + _curRangeText + "\"}";

                        string _requestUrl = ConfigurationManager.AppSettings["RequestSentenceUrl"];

                        HttpWebResponse hwr = await HttpTool.RequesPostAsync(_requestUrl, _jsonParam);
                        HttpTool.ResponseResult _curResult = HttpTool.GetResponseJson<HttpTool.ResponseResult>(hwr);

                        cacheUpdateText = _curRangeText;
                    }                    
                }
            }

            // 恢复用户 Overtype
            Application.Options.Overtype = userOvertype;
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
