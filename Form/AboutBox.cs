using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pings
{
    public partial class AboutBox : Form
    {
        public AboutBox()
        {
            InitializeComponent();

            // 表示文字列（元の labelInfo と同じ内容にする）
            var text = "Pings Ver2.00\noldmanpar\n";
            var author = "oldmanpar";
            var authorUrl = "https://github.com/oldmanpar/Pings/releases";

            // designer の labelInfo と同じ見た目・位置に LinkLabel を作成し、
            // "oldmanpar" 部分だけをリンクにする
            var linkLabelInfo = new LinkLabel
            {
                AutoSize = false,
                Location = this.labelInfo.Location,
                Size = this.labelInfo.Size,
                Font = this.labelInfo.Font,
                ForeColor = this.labelInfo.ForeColor,
                BackColor = this.labelInfo.BackColor,
                Text = text,
                TextAlign = this.labelInfo.TextAlign,
                LinkBehavior = LinkBehavior.HoverUnderline,
                TabIndex = this.labelInfo.TabIndex
            };

            // oldmanpar の開始位置と長さを計算してリンク領域を設定
            var start = text.IndexOf(author, StringComparison.Ordinal);
            if (start >= 0)
            {
                linkLabelInfo.Links.Add(start, author.Length, authorUrl);
            }

            linkLabelInfo.LinkClicked += (s, e) =>
            {
                var url = e.Link.LinkData as string ?? authorUrl;
                try
                {
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                catch
                {
                    try
                    {
                        Process.Start("explorer", url);
                    }
                    catch
                    {
                        // 無視
                    }
                }
            };

            // designer の labelInfo を非表示にして置き換える（位置・大きさ・フォントは維持）
            this.labelInfo.Visible = false;
            this.Controls.Add(linkLabelInfo);

            // フォームのリサイズで位置が変わる場合は同じ位置に保つ（labelInfo に合わせる）
            this.Resize += (s, e) =>
            {
                linkLabelInfo.Location = this.labelInfo.Location;
                linkLabelInfo.Size = this.labelInfo.Size;
            };
        }

        private void okButton_Click(object sender, System.EventArgs e)
        {
            // OKボタンが押されたらフォームを閉じる
            this.Close();
        }
    }
}