using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddInVSTO
{
    public partial class TaskPanel : UserControl
    {
        public static Label DetectTextLabel;
        public TaskPanel()
        {
            InitializeComponent();

            ////初始化Label，用来实时显示更新的文本
            DetectTextLabel = this.detectTextLabel;
        }
    }
}
