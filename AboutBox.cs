using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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

            // 指定された文字列を設定
            this.labelInfo.Text =
                "Pings Ver1.00\n" +
                "Susumu Tanaka";
        }

        private void okButton_Click(object sender, System.EventArgs e)
        {
            // OKボタンが押されたらフォームを閉じる
            this.Close();
        }
    }
}
