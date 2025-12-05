using System;
using System.Drawing;
using System.IO;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pings
{
    /// <summary>
    /// 特定のIPに対してPingを打ち続け、コマンドプロンプト風に表示・ログ保存する独立フォーム
    /// </summary>
    public partial class PingTerminalForm : Form
    {
        private readonly string _targetAddress;
        private readonly string _hostName;
        private TextBox _outputBox;
        private bool _keepRunning = true;
        private string _logFilePath;

        // コンストラクタ
        public PingTerminalForm(string targetAddress, string hostName)
        {
            _targetAddress = targetAddress;
            _hostName = hostName;

            // UIの初期化
            InitializeComponentCustom();

            // ログ保存の準備
            SetupLogFile();
        }

        // フォームが表示されたらPing開始
        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            await RunPingLoopAsync();
        }

        // フォームが閉じられたらループ停止
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _keepRunning = false;
            base.OnFormClosing(e);
        }

        // ---------------------------------------------------------
        // 1. メインロジック: Pingループ
        // ---------------------------------------------------------
        private async Task RunPingLoopAsync()
        {
            // 開始メッセージ
            string startMsg = $"Pinging {_targetAddress} [{_hostName}] with 32 bytes of data:";
            AddLine(startMsg);
            AddLine(new string('-', 60));

            using (Ping pingSender = new Ping())
            {
                // バッファ (32 bytes)
                byte[] buffer = Encoding.ASCII.GetBytes("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa");
                PingOptions options = new PingOptions(64, true); // TTL=64, Don't Fragment

                while (_keepRunning && !this.IsDisposed)
                {
                    try
                    {
                        // Ping送信 (タイムアウト 2000ms)
                        PingReply reply = await pingSender.SendPingAsync(_targetAddress, 2000, buffer, options);

                        // 結果の整形と表示
                        string message = FormatReply(reply);
                        AddLine(message);
                    }
                    catch (Exception ex)
                    {
                        AddLine($"[Error] {ex.InnerException?.Message ?? ex.Message}");
                    }

                    // 1秒待機 (ループ継続チェック)
                    if (_keepRunning) await Task.Delay(1000);
                }
            }

            AddLine(new string('-', 60));
            AddLine($"Ping session for {_targetAddress} ended.");
            AddLine($"Log saved to: {_logFilePath}");
        }

        // ---------------------------------------------------------
        // 2. 出力整形 (コマンドプロンプト風 + タイムスタンプ)
        // ---------------------------------------------------------
        private string FormatReply(PingReply reply)
        {
            string timestamp = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            if (reply.Status == IPStatus.Success)
            {
                // 成功時: [日時] Reply from IP: bytes=32 time=xxms TTL=xx
                return $"[{timestamp}] Reply from {reply.Address}: bytes={reply.Buffer.Length} time={reply.RoundtripTime}ms TTL={reply.Options?.Ttl}";
            }
            else
            {
                // 失敗時
                return $"[{timestamp}] Request timed out. ({reply.Status})";
            }
        }

        // ---------------------------------------------------------
        // 3. 表示とログ保存
        // ---------------------------------------------------------
        private void AddLine(string text)
        {
            if (this.IsDisposed || _outputBox.IsDisposed) return;

            // 画面への追記 (末尾に追加してスクロール)
            _outputBox.AppendText(text + Environment.NewLine);

            // ファイルへの追記
            SaveToLog(text);
        }

        private void SetupLogFile()
        {
            try
            {
                // 実行フォルダ\Ping_Result フォルダを作成
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string logDir = Path.Combine(baseDir, "Ping_Result");

                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                // ファイル名: Ping_IPアドレス_ホスト名_YYYYMMDD.log
                string safeIp = SanitizeFileName(_targetAddress.Replace(":", "-"));
                string safeHost = SanitizeFileName(_hostName);
                string fileName = $"Ping_{safeIp}_{safeHost}_{DateTime.Now:yyyyMMdd}.log";

                _logFilePath = Path.Combine(logDir, fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ログファイルの作成に失敗しました。\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveToLog(string line)
        {
            if (string.IsNullOrEmpty(_logFilePath)) return;

            try
            {
                // 毎回追記して閉じる (ファイルロック時間を最小化し、強制終了時もデータを残すため)
                using (StreamWriter sw = new StreamWriter(_logFilePath, true, Encoding.GetEncoding(932)))
                {
                    sw.WriteLine(line);
                }
            }
            catch
            {
                // ログ書き込みエラーでアプリを止めない
            }
        }

        // ファイル名用に不正文字を置換
        private string SanitizeFileName(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(input.Length);
            foreach (char c in input)
            {
                sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
            }
            return sb.ToString();
        }

        // ---------------------------------------------------------
        // 4. UI構築 (デザイナーレス)
        // ---------------------------------------------------------
        private void InitializeComponentCustom()
        {
            this.Text = $"Ping - {_targetAddress} ({_hostName})";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Icon = SystemIcons.Application; // デフォルトアイコン

            // 黒背景のテキストボックス
            _outputBox = new TextBox();
            _outputBox.Multiline = true;
            _outputBox.ReadOnly = true;
            _outputBox.Dock = DockStyle.Fill;
            _outputBox.BackColor = Color.Black;
            _outputBox.ForeColor = Color.LightGray;
            // 等幅フォント (Consolas がなければ Courier New)
            _outputBox.Font = new Font("Consolas", 10F, FontStyle.Regular);
            if (_outputBox.Font.Name != "Consolas")
            {
                _outputBox.Font = new Font("Courier New", 10F, FontStyle.Regular);
            }
            _outputBox.ScrollBars = ScrollBars.Vertical;

            this.Controls.Add(_outputBox);
        }
    }
}