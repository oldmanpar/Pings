// AboutBox.Designer.cs

namespace Pings
{
    partial class AboutBox
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Label labelInfo;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        private void InitializeComponent()
        {
            this.okButton = new System.Windows.Forms.Button();
            this.labelInfo = new System.Windows.Forms.Label();
            this.SuspendLayout();

            // 
            // labelInfo
            // 
            // ★ 修正箇所 1: AutoSizeをfalse、TextAlignを中央に設定 ★
            this.labelInfo.AutoSize = false;
            this.labelInfo.Font = new System.Drawing.Font("MS UI Gothic", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            // WidthをフォームのClientSizeに合わせて設定 (例: ClientSize.Width - 40)
            this.labelInfo.Location = new System.Drawing.Point(10, 20);
            this.labelInfo.Name = "labelInfo";
            // フォームの幅(264)から余白(20*2)を引いて224を設定 (ClientSize.Width - 20)
            this.labelInfo.Size = new System.Drawing.Size(244, 100);
            this.labelInfo.TabIndex = 0;
            this.labelInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter; // ★ これで中央揃え ★

            // 
            // okButton
            // 
            // ★ 修正箇所 2: AnchorとLocationを中央に設定 ★
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.None))); // NoneまたはRightは好み
                                                                                                                                                               // 264 / 2 - 75 / 2 = 94.5 -> 95
            this.okButton.Location = new System.Drawing.Point(95, 145);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 25);
            this.okButton.TabIndex = 1;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // AboutBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(264, 182);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.labelInfo);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutBox";
            this.Padding = new System.Windows.Forms.Padding(10);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "バージョン情報";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}