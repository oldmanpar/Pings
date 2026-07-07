using System;
using System.Drawing;
using System.Windows.Forms;

namespace Pings
{
    /// <summary>
    /// 「操作方法」ヘルプウィンドウ。
    /// 1画面スクロール式・モードレス（開いたまま本体を操作可能）。
    /// </summary>
    public class HelpForm : Form
    {
        private readonly Font _bodyFont    = new Font("Meiryo UI", 9.5F);
        private readonly Font _sectionFont = new Font("Meiryo UI", 11F, FontStyle.Bold);
        private readonly Font _itemFont    = new Font("Meiryo UI", 9.5F, FontStyle.Bold);
        private static readonly Color SectionColor = Color.FromArgb(0, 70, 140);

        public HelpForm()
        {
            this.Text = "操作方法 - Pings";
            this.Size = new Size(720, 580);
            this.MinimumSize = new Size(520, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;

            var rtb = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                Font = _bodyFont,
                DetectUrls = false,
                WordWrap = true
            };

            // RichTextBox 自体は Padding を持たないため、白背景のパネルで包んで余白を作る
            var rtbHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(16, 14, 10, 4)
            };
            rtbHost.Controls.Add(rtb);

            // 下部の閉じるボタン
            var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            var btnClose = new Button
            {
                Text = "閉じる",
                Width = 90,
                Height = 28,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnClose.Location = new Point(bottomPanel.ClientSize.Width - btnClose.Width - 14, 8);
            btnClose.Click += (s, e) => this.Close();
            bottomPanel.Controls.Add(btnClose);
            this.CancelButton = btnClose;  // Escキーでも閉じられる

            this.Controls.Add(rtbHost);
            this.Controls.Add(bottomPanel);

            BuildHelpText(rtb);

            // 先頭を表示した状態で開く
            rtb.SelectionStart = 0;
            rtb.SelectionLength = 0;
            rtb.ScrollToCaret();
        }

        // ====================================================================
        // ヘルプ本文の構築
        // ====================================================================

        private void BuildHelpText(RichTextBox rtb)
        {
            Section(rtb, "1. 基本操作の流れ");
            Body(rtb,
                "  1) 監視対象の登録\n" +
                "     ・監視統計タブのグリッドに対象アドレス・Host名を直接入力\n" +
                "     ・またはメニュー「ファイル → 対象アドレス読込」でファイルから読み込み\n" +
                "  2) 送信間隔・タイムアウトを設定（監視中は変更不可）\n" +
                "  3) 「Ping開始」で監視スタート、「Ping停止」で終了\n" +
                "  4) 結果の保存\n" +
                "     ・オプション「Ping結果の自動保存」がONの場合は停止時に自動保存\n" +
                "     ・「Ping結果保存」ボタンで手動保存も可能\n");

            Section(rtb, "2. タブの説明");
            Item(rtb, "【監視統計】");
            Body(rtb,
                "  ・各拠点の応答状況をリアルタイム表示\n" +
                "  ・ステータス色： 緑=OK ／ 赤=Down ／ 青=復旧\n" +
                "  ・行をダブルクリック → その拠点の個別Pingターミナルを起動\n" +
                "  ・Trace列のヘッダーをダブルクリック → 全チェックを一括切替\n");
            Item(rtb, "【障害イベントログ】");
            Body(rtb,
                "  ・Down開始〜復旧の記録。Down前後の統計スナップショット付き\n" +
                "  ・カラムヘッダーのクリックでソート\n");
            Item(rtb, "【Traceroute】");
            Body(rtb,
                "  ・Trace列にチェックを入れた拠点へ実行（同時4拠点ずつ）\n" +
                "  ・停止・保存・クリアは画面下部のTracerouteグループのボタンで操作\n");
            Item(rtb, "【ステータスボード】");
            Body(rtb,
                "  ・全拠点の状態を1行チップで一覧表示（順番・アドレス・Host名・RTT/状態）\n" +
                "  ・チップは縦方向優先で配置され、ウィンドウの高さで右の列へ折り返し\n" +
                "  ・チップをダブルクリック → 監視統計タブの該当行へジャンプ\n");

            Section(rtb, "3. ファイル入出力");
            Body(rtb,
                "  ・対象アドレスの保存/読込： メニュー「ファイル」から。CSV形式で保存。\n" +
                "    読込はCSV形式に加え、スペース・タブ区切りのテキストにも対応\n" +
                "  ・結果の保存先： 実行ファイルと同じ場所の Ping_Result / Traceroute_Result フォルダ\n");

            Section(rtb, "4. オプションメニュー");
            Body(rtb,
                "  ・Ping結果の自動保存 / Trace結果の自動保存 … 停止時・完了時に自動でファイル保存\n" +
                "  ・ジッタ①②・標準偏差の表示切替 … 監視統計・障害イベントログの該当カラムの表示/非表示\n" +
                "  ・tracerouteで名前解決を行わない … ONにすると実行が高速化される\n");

            Section(rtb, "5. 統計指標の解説");
            Item(rtb, "【ジッタ①（最大値−最小値）】");
            Body(rtb,
                "  セッション中のRTT最大値と最小値の差。「最悪ケースの揺らぎ幅」を示す。\n" +
                "  瞬間的なスパイクも拾うため、ジッタ②より大きな値になりやすい。\n" +
                "   目安： 一般的なLAN環境では 5ms 以下、WAN回線では 20〜50ms 程度が安定の基準。\n");
            Item(rtb, "【ジッタ②（パケットペア平均）】");
            Body(rtb,
                "  連続する2つのPing応答時間の差（絶対値）の平均。「普段の揺らぎの平均値」。\n" +
                "  短い時間スケールの揺らぎを捉えやすく、RFC 3550（RTP）のジッタ概念に近い。\n" +
                "  VoIP・ビデオ会議など音声・映像品質の評価にはジッタ②の方が実用的。\n" +
                "   目安： VoIP品質確保には 30ms 以下が推奨（ITU-T G.114 参考値）。\n");
            Item(rtb, "【標準偏差】");
            Body(rtb,
                "  RTTのばらつきの統計指標。平均からの散らばり具合を表す。\n" +
                "  小さいほど応答時間が安定。ジッタ①②が「差」に着目するのに対し、\n" +
                "  標準偏差は全サンプルを使って「平均からの距離」を総合評価する。\n" +
                "   目安： LAN環境では 2〜3ms 以下、WAN・インターネット回線では 10〜20ms 以下が良好。\n");
            Item(rtb, "【使い分けの目安】");
            Body(rtb,
                "  ・ジッタ①　→ 瞬間的なスパイクの有無、最悪値の把握\n" +
                "  ・ジッタ②　→ VoIP・会議ツールの体感品質の評価\n" +
                "  ・標準偏差　→ 回線間の安定度の比較・トレンド監視\n" +
                "\n" +
                "  ※ 目安値は回線種別・距離・環境によって大きく異なります。あくまで参考値としてください。\n");
        }

        // 章見出し（青・太字・大きめ）
        private void Section(RichTextBox rtb, string text)
        {
            if (rtb.TextLength > 0) rtb.AppendText("\n");
            rtb.SelectionFont = _sectionFont;
            rtb.SelectionColor = SectionColor;
            rtb.AppendText("■ " + text + "\n");
        }

        // 項目見出し（太字）
        private void Item(RichTextBox rtb, string text)
        {
            rtb.SelectionFont = _itemFont;
            rtb.SelectionColor = Color.Black;
            rtb.AppendText(text + "\n");
        }

        // 本文
        private void Body(RichTextBox rtb, string text)
        {
            rtb.SelectionFont = _bodyFont;
            rtb.SelectionColor = Color.Black;
            rtb.AppendText(text);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _bodyFont.Dispose();
                _sectionFont.Dispose();
                _itemFont.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
