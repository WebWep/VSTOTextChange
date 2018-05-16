using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddInVSTO
{
    public partial class MenuRibbon
    {
        private void MenuRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// 发送整个文档内容到指定接口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void sendBtn_Click(object sender, RibbonControlEventArgs e)
        {
            ////获取当前文档的内容
            Microsoft.Office.Interop.Word.Range _rng = Globals.ThisAddIn.Application.ActiveDocument.Content;
            string _curRangeText = _rng.Text;

            string _jsonParam = "{\"document\":\"" + _curRangeText + "\"}";

            string _requesUrl = ConfigurationManager.AppSettings["RequestDocumentUrl"];

            try
            {
                //// 异步等待请求WebAPI
                var hwr = await HttpTool.RequesPostAsync(_requesUrl, _jsonParam);
                HttpTool.ResponseResult _curResult = HttpTool.GetResponseJson<HttpTool.ResponseResult>(hwr);

                if (_curResult.Code == 2000)
                {
                    MessageBox.Show("远程保存成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(_curResult.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("发送整个文档内容出错：" + ex.Message);
            }           
        }

        /// <summary>
        /// 文本内容更新，实时监听开关
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void detectTBtn_Click(object sender, RibbonControlEventArgs e)
        {
            bool _curChecked = this.detectTBtn.Checked;

            //// 同步TaskPanel
            if (Globals.ThisAddIn.CurCustomTaskPane != null)
            {
                Globals.ThisAddIn.CurCustomTaskPane.Visible = _curChecked;
            }
            if (_curChecked)
            {
                ////打开开关, 挂上钩子
                Globals.ThisAddIn.SetWindowsHooks();              
            }
            else {
                ////关闭开关, 卸载钩子
                Globals.ThisAddIn.UnhookWindowsHooks();
            }
        }
    }
}
