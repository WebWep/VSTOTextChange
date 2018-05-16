namespace WordAddInVSTO
{
    partial class TaskPanel
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TaskPanel));
            this.detectTextLabel = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.loadingPictureBox = new System.Windows.Forms.PictureBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.loadingPictureBox)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // detectTextLabel
            // 
            this.detectTextLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.detectTextLabel.ImageIndex = 0;
            this.detectTextLabel.Location = new System.Drawing.Point(20, 50);
            this.detectTextLabel.Name = "detectTextLabel";
            this.detectTextLabel.Size = new System.Drawing.Size(223, 209);
            this.detectTextLabel.TabIndex = 0;
            this.detectTextLabel.Text = "     ";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.loadingPictureBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 309);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(263, 38);
            this.panel1.TabIndex = 1;
            // 
            // loadingPictureBox
            // 
            this.loadingPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("loadingPictureBox.Image")));
            this.loadingPictureBox.Location = new System.Drawing.Point(3, 3);
            this.loadingPictureBox.Name = "loadingPictureBox";
            this.loadingPictureBox.Size = new System.Drawing.Size(32, 32);
            this.loadingPictureBox.TabIndex = 2;
            this.loadingPictureBox.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.detectTextLabel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Padding = new System.Windows.Forms.Padding(20, 50, 20, 50);
            this.panel2.Size = new System.Drawing.Size(263, 309);
            this.panel2.TabIndex = 2;
            // 
            // TaskPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "TaskPanel";
            this.Size = new System.Drawing.Size(263, 347);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.loadingPictureBox)).EndInit();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox loadingPictureBox;
        private System.Windows.Forms.Label detectTextLabel;
        private System.Windows.Forms.Panel panel2;
    }
}
