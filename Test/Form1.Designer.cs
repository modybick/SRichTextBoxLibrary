namespace Test
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.sRichTextBox1 = new SRichTextBoxLibrary.SuperRichTextBox();
            this.SuspendLayout();
            // 
            // sRichTextBox1
            // 
            this.sRichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sRichTextBox1.Font = new System.Drawing.Font("メイリオ", 10F);
            this.sRichTextBox1.Location = new System.Drawing.Point(0, 0);
            this.sRichTextBox1.Name = "sRichTextBox1";
            this.sRichTextBox1.Size = new System.Drawing.Size(284, 262);
            this.sRichTextBox1.TabIndex = 0;
            this.sRichTextBox1.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.sRichTextBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private SRichTextBoxLibrary.SuperRichTextBox sRichTextBox1;
    }
}

