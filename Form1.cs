using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection; // Reflection を使用するために追加

// ---------------------------------------------------------
// 1. データ保持クラス (障害イベントログ)
// ---------------------------------------------------------

/// <summary>
/// 発生した障害イベント（Down → 復旧）を記録するクラス
/// </summary>
public class DisruptionLogItem
{
    public string 対象アドレス { get; set; }
    public string Host名 { get; set; }
    public DateTime Down開始日時 { get; set; }
    public DateTime 復旧日時 { get; set; }
    public string 失敗時間mmss { get; set; }

    // 統計値
    public double Down前平均ms { get; set; }
    public long Down前最小ms { get; set; }
    public long Down前最大ms { get; set; }

    public double 復旧後平均ms { get; set; }
    public long 復旧後最小ms { get; set; }
    public long 復旧後最大ms { get; set; }
}


// ---------------------------------------------------------
// 2. ソート可能な BindingList の実装 (新規追加)
// ---------------------------------------------------------

/// <summary>
/// DataGridViewでのソートを可能にするために、BindingListを拡張したクラス
/// </summary>
public class SortableBindingList<T> : BindingList<T>
{
    private List<T> originalList;
    private bool isSorted;
    private ListSortDirection sortDirection;
    private PropertyDescriptor sortProperty;

    public SortableBindingList(IList<T> list) : base(list)
    {
        originalList = (List<T>)list;
    }

    protected override bool SupportsSortingCore => true;
    protected override bool IsSortedCore => isSorted;
    protected override PropertyDescriptor SortPropertyCore => sortProperty;
    protected override ListSortDirection SortDirectionCore => sortDirection;

    protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
    {
        originalList = (List<T>)this.Items;

        // ソート対象のプロパティの Type を取得
        Type propType = prop.PropertyType;

        // IComparer を利用した汎用的な比較ロジック
        Comparison<T> comparer = (T x, T y) =>
        {
            object xValue = prop.GetValue(x);
            object yValue = prop.GetValue(y);

            // nullチェック
            if (xValue == null && yValue == null) return 0;
            if (xValue == null) return (direction == ListSortDirection.Ascending) ? -1 : 1;
            if (yValue == null) return (direction == ListSortDirection.Ascending) ? 1 : -1;

            // IComparable を実装している型（string, int, DateTime, doubleなど）で比較
            if (xValue is IComparable comparableX)
            {
                return comparableX.CompareTo(yValue) * (direction == ListSortDirection.Ascending ? 1 : -1);
            }

            // IComparable でない場合（通常は発生しないはず）
            return 0;
        };

        originalList.Sort(comparer);

        // ソート結果を反映
        isSorted = true;
        sortProperty = prop;
        sortDirection = direction;
        OnListChanged(new ListChangedEventArgs(ListChangedType.Reset, -1));
    }

    protected override void RemoveSortCore()
    {
        isSorted = false;
        sortDirection = ListSortDirection.Ascending;
        sortProperty = null;
        OnListChanged(new ListChangedEventArgs(ListChangedType.Reset, -1));
    }

    // DataGridView でのソートアイコン表示のために必要なメソッド
    public void Sort(PropertyDescriptor prop, ListSortDirection direction)
    {
        ApplySortCore(prop, direction);
    }
}


namespace Pings
{
    // ---------------------------------------------------------
    // 3. データ保持クラス (監視対象ノードの情報と統計情報)
    // ---------------------------------------------------------
    public class PingMonitorItem
    {
        // I. UI表示用プロパティ
        public int 順番 { get; set; }
        public string ステータス { get; set; }
        public string 対象アドレス { get; set; }
        public string Host名 { get; set; }

        // 統計情報
        public long 送信回数 { get; private set; }
        public long 失敗回数 { get; private set; }
        public int 連続失敗回数 { get; private set; }
        public string 連続失敗時間s { get; private set; }
        public string 最大失敗時間s { get; private set; }

        // メイングリッド表示用（現在のセッションの統計）
        public long 時間ms { get; set; }
        public double 平均値ms { get; private set; }
        public long 最小値ms { get; private set; }
        public long 最大値ms { get; private set; }

        // II. 内部状態管理用フィールド
        public bool IsUp { get; private set; } = false;
        public bool HasBeenDown { get; private set; } = false;

        private DateTime? _continuousDownStartTime = null;
        private TimeSpan _maxDisruptionDuration = TimeSpan.Zero;

        // 現在アクティブなログエントリを保持するフィールド
        private DisruptionLogItem _activeLogItem = null;

        // 計算用: 現在の「Up期間」の累積データ（復旧時にリセットされる）
        private long _currentSessionSum = 0;
        private int _currentSessionCount = 0;
        private long _currentSessionMin = 0;
        private long _currentSessionMax = 0;

        // Down直前の統計スナップショット
        private bool _isCurrentlyDown = false;
        private double _snapAvg = 0.0;
        private long _snapMin = 0;
        private long _snapMax = 0;

        // UIからの入力値
        public int 送信間隔ms { get; set; } = 500;
        public int タイムアウトms { get; set; } = 1000;

        // III. コンストラクタとリセット
        public PingMonitorItem() : this(0, "", "")
        {
        }

        public PingMonitorItem(int index, string address, string hostName)
        {
            順番 = index;
            対象アドレス = address;
            Host名 = hostName;
            ResetStatistics();
        }

        public void ResetStatistics()
        {
            送信回数 = 0;
            失敗回数 = 0;
            連続失敗回数 = 0;
            時間ms = 0;

            平均値ms = 0.0;
            最小値ms = 0;
            最大値ms = 0;

            _currentSessionSum = 0;
            _currentSessionCount = 0;
            _currentSessionMin = 0;
            _currentSessionMax = 0;

            ステータス = "";
            連続失敗時間s = "";
            最大失敗時間s = "";

            _continuousDownStartTime = null;
            _maxDisruptionDuration = TimeSpan.Zero;
            IsUp = false;
            HasBeenDown = false;
            _isCurrentlyDown = false;
            _activeLogItem = null;
        }

        // IV. 状態更新ロジック
        public void UpdateStatistics(long rtt, bool success, Action<DisruptionLogItem> logAction)
        {
            送信回数++;

            if (success)
            {
                // ========== 復旧時の処理 ==========
                if (_isCurrentlyDown)
                {
                    // 復旧イベント発生
                    if (_continuousDownStartTime.HasValue)
                    {
                        TimeSpan duration = DateTime.Now - _continuousDownStartTime.Value;

                        // ログ生成
                        DisruptionLogItem newLogItem = new DisruptionLogItem
                        {
                            対象アドレス = 対象アドレス,
                            Host名 = Host名,
                            Down開始日時 = _continuousDownStartTime.Value,
                            復旧日時 = DateTime.Now,
                            失敗時間mmss = duration.ToString(@"mm\:ss"),

                            // Down前の統計
                            Down前平均ms = _snapAvg,
                            Down前最小ms = _snapMin,
                            Down前最大ms = _snapMax,

                            // 復旧後統計の初期値 (復旧直後の値)
                            復旧後平均ms = rtt,
                            復旧後最小ms = rtt,
                            復旧後最大ms = rtt
                        };
                        logAction(newLogItem);

                        // 生成したログエントリを参照として保持
                        _activeLogItem = newLogItem;
                    }

                    // セッション統計のリセット
                    _currentSessionSum = 0;
                    _currentSessionCount = 0;
                    _currentSessionMin = 0;
                    _currentSessionMax = 0;

                    _isCurrentlyDown = false;
                }

                // ========== 共通: 成功時の統計更新 ==========
                時間ms = rtt;
                IsUp = true;
                連続失敗回数 = 0;
                連続失敗時間s = "";
                _continuousDownStartTime = null;

                ステータス = HasBeenDown ? "復旧" : "OK";

                // セッション統計の更新 (現在のUp期間の集計)
                _currentSessionCount++;
                _currentSessionSum += rtt;

                // 最小値/最大値の更新
                if (_currentSessionCount == 1)
                {
                    _currentSessionMin = rtt;
                    _currentSessionMax = rtt;
                }
                else
                {
                    if (rtt < _currentSessionMin) _currentSessionMin = rtt;
                    if (rtt > _currentSessionMax) _currentSessionMax = rtt;
                }

                // 公開プロパティへの反映
                平均値ms = (double)_currentSessionSum / _currentSessionCount;
                最小値ms = _currentSessionMin;
                最大値ms = _currentSessionMax;

                // アクティブログエントリの継続更新
                if (_activeLogItem != null)
                {
                    // 継続して成功している間は、ログエントリの「復旧後」の統計を現在のセッション統計で更新
                    _activeLogItem.復旧後平均ms = 平均値ms;
                    _activeLogItem.復旧後最小ms = 最小値ms;
                    _activeLogItem.復旧後最大ms = 最大値ms;
                }

            }
            else // failure
            {
                // ========== 失敗時の処理 ==========
                if (!_isCurrentlyDown)
                {
                    // Down開始イベント: Downになったらアクティブログをクリアし、更新を終了する
                    _activeLogItem = null;

                    // Down開始イベント: 現在の統計をスナップショットとして保存
                    _snapAvg = 平均値ms;
                    _snapMin = 最小値ms;
                    _snapMax = 最大値ms;

                    _continuousDownStartTime = DateTime.Now;
                    _isCurrentlyDown = true;
                }

                // Down継続中も毎回実行するべき処理
                失敗回数++;
                時間ms = 0;
                ステータス = "Down";
                IsUp = false;
                HasBeenDown = true;

                連続失敗回数++;

                TimeSpan currentDuration = DateTime.Now - _continuousDownStartTime.Value;
                連続失敗時間s = currentDuration.ToString(@"mm\:ss");

                if (currentDuration > _maxDisruptionDuration)
                {
                    _maxDisruptionDuration = currentDuration;
                    最大失敗時間s = _maxDisruptionDuration.ToString(@"mm\:ss");
                }
            }
        }
    }


    // ---------------------------------------------------------
    // 4. メインフォーム (UI と ロジック)
    // ---------------------------------------------------------
    public partial class Form1 : Form
    {
        private DataGridView dgvMonitor;
        private TextBox txtStartTime, txtEndTime;
        private ComboBox cmbInterval, cmbTimeout;
        private Button btnPingStart, btnStop, btnClear, btnSave, btnExit;

        private TabControl tabControl;
        private DataGridView dgvLog;

        // ★変更: BindingList から SortableBindingList へ★
        private SortableBindingList<DisruptionLogItem> disruptionLogList;

        private BindingList<PingMonitorItem> monitorList;
        private CancellationTokenSource cts;
        private int _nextIndex = 1;

        private System.Windows.Forms.Timer uiUpdateTimer;
        private const int UiUpdateInterval = 200;

        // ★追加: ソート状態保持用フィールド★
        private DataGridViewColumn currentSortColumn = null;
        private ListSortDirection currentSortDirection = ListSortDirection.Ascending;


        public Form1()
        {
            // ★修正箇所: デザイナー生成コードを呼び出す★
            InitializeComponent();

            // ★修正箇所: カスタムコンポーネントの初期化と配置★
            InitializeCustomComponents();

            // ★修正箇所: DataGridView のカラム設定を行う (InitializeCustomComponentsで dgvMonitor が生成された後)★
            SetupDataGridViewColumns();

            // ★修正箇所: メニューの設定を行う (InitializeCustomComponentsで MainMenuStrip が設定された後)★
            SetupMenuStrip();

            monitorList = new BindingList<PingMonitorItem>();
            dgvMonitor.DataSource = monitorList;

            dgvMonitor.DefaultValuesNeeded += DgvMonitor_DefaultValuesNeeded;
            this.FormClosing += Form1_FormClosing;

            // 初期データの追加
            monitorList.Add(new PingMonitorItem(_nextIndex++, "127.0.0.1", "loopback"));
            monitorList.Add(new PingMonitorItem(_nextIndex++, "8.8.8.8", "Google DNS1"));
            monitorList.Add(new PingMonitorItem(_nextIndex++, "8.8.4.4", "Google DNS2"));

            // ★修正箇所: UIを初期状態にする★
            UpdateUiState("Initial");
        }

        private void UpdateUiState(string state)
        {
            if (this.MainMenuStrip != null)
            {
                ToolStripMenuItem fileMenu = this.MainMenuStrip.Items[0] as ToolStripMenuItem;
                if (fileMenu != null)
                {
                    // [ファイル] -> [対象アドレス保存/読込] は Initial 状態でのみ有効
                    fileMenu.DropDownItems[0].Enabled = (state == "Initial" || state == "Stopped"); // 対象アドレス保存
                    fileMenu.DropDownItems[1].Enabled = (state == "Initial"); // 対象アドレス読込
                    fileMenu.DropDownItems[3].Enabled = true;                 // 終了 (常に有効)
                }
            }

            bool isEditable = (state == "Initial");
            dgvMonitor.ReadOnly = !isEditable;
            dgvMonitor.AllowUserToAddRows = isEditable;

            cmbInterval.Enabled = (state == "Initial");
            cmbTimeout.Enabled = (state == "Initial");

            switch (state)
            {
                case "Initial":
                    btnPingStart.Enabled = true;
                    btnStop.Enabled = false;
                    btnClear.Enabled = false;
                    btnSave.Enabled = false;
                    break;
                case "Running":
                    btnPingStart.Enabled = false;
                    btnStop.Enabled = true;
                    btnClear.Enabled = false;
                    btnSave.Enabled = false;
                    break;
                case "Stopped":
                    btnPingStart.Enabled = true;
                    btnStop.Enabled = false;
                    btnClear.Enabled = true;
                    btnSave.Enabled = true;
                    break;
            }
        }

        private void SetupMenuStrip()
        {
            MenuStrip menuStrip = new MenuStrip();
            menuStrip.Dock = DockStyle.Top;

            ToolStripMenuItem fileMenu = new ToolStripMenuItem("ファイル");

            ToolStripMenuItem saveAddressItem = new ToolStripMenuItem("対象アドレス保存");
            saveAddressItem.Click += BtnSaveAddress_Click;
            fileMenu.DropDownItems.Add(saveAddressItem);

            ToolStripMenuItem loadAddressItem = new ToolStripMenuItem("対象アドレス読込");
            loadAddressItem.Click += BtnLoadAddress_Click;
            fileMenu.DropDownItems.Add(loadAddressItem);

            fileMenu.DropDownItems.Add(new ToolStripSeparator());

            ToolStripMenuItem exitItem = new ToolStripMenuItem("終了");
            exitItem.Click += BtnExit_Click;
            fileMenu.DropDownItems.Add(exitItem);

            menuStrip.Items.Add(fileMenu);
            menuStrip.Items.Add(new ToolStripMenuItem("オプション")); // オプションメニュー (空でも可)

            ToolStripMenuItem helpMenu = new ToolStripMenuItem("ヘルプ");
            ToolStripMenuItem versionItem = new ToolStripMenuItem("バージョン情報");
            versionItem.Click += VersionItem_Click;
            helpMenu.DropDownItems.Add(versionItem);

            menuStrip.Items.Add(helpMenu);

            this.Controls.Add(menuStrip);
            this.MainMenuStrip = menuStrip; // Form の MainMenuStrip プロパティに設定

            menuStrip.SendToBack();
        }

        private void BtnPingStart_Click(object sender, EventArgs e)
        {
            // 1. 既に監視が開始されていたら、一度停止させる（再開時の安全策）
            StopMonitoring();
            ClearVeiw();

            // 2. DataGridViewの編集内容を確定
            // ユーザーがアドレスを入力中だった場合、この操作でリストに反映される
            dgvMonitor.EndEdit();

            // 3. 次の新規行の「順番」を決定
            if (monitorList.Any())
            {
                // 既存の最大の順番+1を設定
                _nextIndex = monitorList.Max(i => i.順番) + 1;
            }
            else
            {
                _nextIndex = 1;
            }


            // 4. Ping監視の開始
            StartMonitoring();
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            StopMonitoring();
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            StopMonitoring();
            ClearVeiw();

            if (btnPingStart != null)
            {
                btnPingStart.Enabled = true; // 有効化
            }

            if (dgvMonitor != null)
            {
                dgvMonitor.AllowUserToDeleteRows = true;
            }
        }

        private void ResetLogSortIndicators()
        {
            if (currentSortColumn != null)
            {
                currentSortColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                currentSortColumn = null;
                currentSortDirection = ListSortDirection.Ascending;
            }
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            // 監視実行中は終了をブロック
            if (cts != null)
            {
                MessageBox.Show("Ping監視中はアプリケーションを終了できません。\n[停止]ボタンを押して監視を止めてください。", "終了不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            StopMonitoring();
            this.Close();
        }

        private void VersionItem_Click(object sender, EventArgs e)
        {
            // AboutBox クラスが存在することを前提とします。
            using (var aboutBox = new AboutBox())
            {
                aboutBox.ShowDialog(this);
            }
        }





        private void DgvMonitor_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            if (e.Row.IsNewRow)
            {
                e.Row.Cells[1].Value = _nextIndex;
            }
        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中はアプリケーションを終了できません。\n[停止]ボタンを押して監視を止めてください。", "終了不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
            else
            {
                StopMonitoring();
            }
        }


        private void UiUpdateTimer_Tick(object sender, EventArgs e)
        {
            if (dgvMonitor.IsHandleCreated)
            {
                dgvMonitor.Refresh();
                // dgvLog.Refresh(); はソート機能が ListChangedType.Reset を発火するため不要。
                // ただし、復旧後平均などの継続更新のために残しておいても害はない。
                dgvLog.Refresh();
            }
        }

        private void BtnSaveAddress_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSVファイル (*.csv)|*.csv|テキストファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*";
                sfd.Title = "監視対象アドレスを保存";
                sfd.FileName = "PingTargets.csv";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Shift-JIS (Encoding 932) で書き込み
                        using (StreamWriter sw = new StreamWriter(sfd.FileName, false, System.Text.Encoding.GetEncoding(932)))
                        {
                            // ヘッダー行
                            sw.WriteLine("対象アドレス,Host名");

                            // monitorList からデータを取得
                            // ユーザーが追加中の空の最終行を除外するため、Host名または対象アドレスがある行のみを保存します。
                            foreach (var item in monitorList.Where(i => !string.IsNullOrEmpty(i.対象アドレス) || !string.IsNullOrEmpty(i.Host名)))
                            {
                                // CSV形式でアドレスとホスト名を書き出し
                                sw.WriteLine($"{item.対象アドレス},{item.Host名}");
                            }
                        }
                        MessageBox.Show("監視対象アドレスをファイルに保存しました。", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ファイルの保存中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnLoadAddress_Click(object objectSender, EventArgs eventArgs)
        {
            // 監視中はロードをブロック
            if (cts != null)
            {
                MessageBox.Show("Ping監視中はアドレスを読み込めません。\n[停止]ボタンを押して監視を止めてください。", "読込不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                // CSV、TXT、LOGなどExPing形式で使われる拡張子に対応
                ofd.Filter = "アドレスファイル (*.csv;*.txt;*.log)|*.csv;*.txt;*.log|すべてのファイル (*.*)|*.*";
                ofd.Title = "監視対象アドレスを読み込み";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var newMonitorList = new List<PingMonitorItem>();
                        int index = 1;

                        // Shift-JIS (Encoding 932) でファイルを読み込みます。
                        System.Text.Encoding encoding = System.Text.Encoding.GetEncoding(932);

                        var allLines = new List<string>();
                        using (StreamReader sr = new StreamReader(ofd.FileName, encoding))
                        {
                            string line;
                            // 1. ファイル全体を読み込む
                            while ((line = sr.ReadLine()) != null)
                            {
                                allLines.Add(line);
                            }
                        }

                        // 2. ヘッダー検出とスキップ
                        bool isSavedCsvFormat = false;
                        if (allLines.Any())
                        {
                            // ヘッダー判定用に空白を除去
                            string headerCandidate = allLines[0].Trim().Replace(" ", "");

                            // プログラムで保存したCSVヘッダーと一致するかチェック（大文字小文字無視）
                            if (headerCandidate.Equals("対象アドレス,Host名", StringComparison.OrdinalIgnoreCase))
                            {
                                allLines.RemoveAt(0); // ヘッダー行をスキップ
                                isSavedCsvFormat = true;
                            }
                        }

                        // 3. 各行を解析
                        foreach (string currentLine in allLines)
                        {
                            string trimmedLine = currentLine.Trim();

                            if (string.IsNullOrWhiteSpace(trimmedLine)) continue;

                            // 注釈行のチェック (行頭が [ # ; ')
                            if (trimmedLine.StartsWith("[") || trimmedLine.StartsWith("#") || trimmedLine.StartsWith(";") || trimmedLine.StartsWith("'"))
                            {
                                continue;
                            }

                            string address;
                            string hostName;

                            if (isSavedCsvFormat)
                            {
                                // ★ プログラム保存CSV専用パス ★
                                // カンマ区切りを強制し、ExPingのスペース/タブ解析を完全にバイパス
                                string[] parts = currentLine.Split(new[] { ',' }, 2).Select(p => p.Trim()).ToArray();
                                address = parts[0];
                                hostName = parts.Length > 1 ? parts[1] : "";
                            }
                            else
                            {
                                // ★ ExPing形式/プレーンテキスト パス ★

                                // 2-1. ExPing形式の解析 (半角スペースまたはタブが区切り)
                                int separatorIndex = -1;
                                for (int i = 0; i < currentLine.Length; i++)
                                {
                                    // 行頭ではない、かつ、半角スペースまたはタブを見つける
                                    if ((currentLine[i] == ' ' || currentLine[i] == '\t') && currentLine.Substring(0, i).Trim().Length > 0)
                                    {
                                        separatorIndex = i;
                                        break;
                                    }
                                }

                                if (separatorIndex != -1)
                                {
                                    // ExPing形式: アドレスは区切り文字まで、備考はそれ以降
                                    address = currentLine.Substring(0, separatorIndex).Trim();
                                    hostName = currentLine.Substring(separatorIndex).Trim();

                                    // Host名として残った文字列を「綺麗に整理」
                                    hostName = System.Text.RegularExpressions.Regex.Replace(hostName, @"\s+", " ").Trim();
                                }
                                else if (currentLine.Contains(','))
                                {
                                    // 2-2. CSV形式 (フォールバック): ExPing形式でなかった場合、カンマで分割
                                    string[] parts = currentLine.Split(new[] { ',' }, 2).Select(p => p.Trim()).ToArray();
                                    address = parts[0];
                                    hostName = parts.Length > 1 ? parts[1] : "";
                                }
                                else
                                {
                                    // 2-3. 単純なアドレスのみ
                                    address = trimmedLine;
                                    hostName = "";
                                }
                            }

                            // アドレスが空の場合はスキップ
                            if (string.IsNullOrEmpty(address)) continue;

                            // リストに追加
                            newMonitorList.Add(new PingMonitorItem(index++, address, hostName));
                        }

                        // 既存のリストをクリアして新しいリストに置き換え、UIを更新
                        monitorList.Clear();
                        foreach (var item in newMonitorList)
                        {
                            monitorList.Add(item);
                        }

                        _nextIndex = index; // 次に追加される行のためのインデックスを更新

                        monitorList.ResetBindings();
                        UpdateUiState("Initial");

                        MessageBox.Show("監視対象アドレスをファイルから読み込みました。", "読込完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ファイルの読み込み中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void BtnSaveResult_Click(object sender, EventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中は結果を保存できません。\n[停止]ボタンを押して監視を止めてください。", "保存不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.FileName = $"Pings_Result_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                sfd.Filter = "CSVファイル (*.csv)|*.csv|すべてのファイル (*.*)|*.*";
                sfd.Title = "監視結果を保存";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Shift-JIS (Encoding.GetEncoding(932)) でファイルに書き込みます。
                        using (StreamWriter sw = new StreamWriter(sfd.FileName, false, System.Text.Encoding.GetEncoding(932)))
                        {
                            // --- 1. ヘッダー情報 ---
                            sw.WriteLine($"開始日時　：　{txtStartTime.Text}");
                            sw.WriteLine($"終了日時　：　{txtEndTime.Text}");
                            sw.WriteLine($"送信間隔　：　{cmbInterval.Text} [ms]     タイムアウト　：　{cmbTimeout.Text} [ms]");
                            sw.WriteLine(); // 空行

                            // --- 2. 監視統計データ (dgvMonitorの内容) ---
                            sw.WriteLine("--- 監視統計 ---");

                            // ヘッダー行
                            string header = "ｽﾃｰﾀｽ,順番,対象アドレス,Host名,送信回数,失敗回数,連続失敗回数,連続失敗時間[mm:ss],最大失敗時間[mm:ss],時間[ms],平均値[ms],最小値[ms],最大値[ms]";
                            sw.WriteLine(header);

                            // データ行
                            foreach (var item in monitorList.OrderBy(i => i.順番))
                            {
                                string line = string.Join(",",
                                    item.ステータス,
                                    item.順番,
                                    item.対象アドレス,
                                    item.Host名,
                                    item.送信回数,
                                    item.失敗回数,
                                    item.連続失敗回数,
                                    item.連続失敗時間s,
                                    item.最大失敗時間s,
                                    item.時間ms,
                                    // 平均値はF1形式で出力
                                    $"{item.平均値ms:F1}",
                                    item.最小値ms,
                                    item.最大値ms
                                );
                                sw.WriteLine(line);
                            }

                            sw.WriteLine(); // 空行
                            sw.WriteLine(); // 空行

                            // --- 3. 障害イベントログデータ (dgvLogの内容) ---
                            sw.WriteLine("--- 障害イベントログ ---");

                            if (disruptionLogList.Any())
                            {
                                // ヘッダー行
                                string logHeader = "対象アドレス,Host名,Down開始日時,復旧日時,失敗時間[mm:ss]," +
                                                   "Down前平均[ms],Down前最小[ms],Down前最大[ms]," +
                                                   "復旧後平均[ms],復旧後最小[ms],復旧後最大[ms]";
                                sw.WriteLine(logHeader);

                                // データ行 (常に復旧日時でソートして書き出す)
                                foreach (var logItem in disruptionLogList.OrderBy(i => i.復旧日時))
                                {
                                    string logLine = string.Join(",",
                                        logItem.対象アドレス,
                                        logItem.Host名,
                                        logItem.Down開始日時.ToString("yyyy/MM/dd HH:mm:ss"),
                                        logItem.復旧日時.ToString("yyyy/MM/dd HH:mm:ss"),
                                        logItem.失敗時間mmss,
                                        // 平均値はF1形式で出力
                                        $"{logItem.Down前平均ms:F1}",
                                        logItem.Down前最小ms,
                                        logItem.Down前最大ms,
                                        $"{logItem.復旧後平均ms:F1}",
                                        logItem.復旧後最小ms,
                                        logItem.復旧後最大ms
                                    );
                                    sw.WriteLine(logLine);
                                }
                            }
                            else
                            {
                                sw.WriteLine("障害イベントは記録されていません。");
                            }
                        }

                        MessageBox.Show("監視結果をCSVファイルに保存しました。", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ファイルの保存中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }


            if (btnPingStart != null)
            {
                btnPingStart.Enabled = true; // 有効化
            }
        }


        private void AddDisruptionLogItem(DisruptionLogItem item)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<DisruptionLogItem>(AddDisruptionLogItem), item);
            }
            else
            {
                disruptionLogList.Add(item);
                // ソート状態を維持して追加された行が正しい位置に来るようにリスト全体をリセット
                // ただし、ソートされている場合にのみ ApplySortCore を呼び出す必要があるため、
                // dgvLog_ColumnHeaderMouseClick で最後にソートした状態を保持し、
                // ApplySortCore を手動で呼び出すか、ここでは行末に追加するだけに留める。
                // データの整合性を優先し、ここでは単に追加のみとする。

                // 復旧日時でソートされていない限り、行末に追加し続ける
                if (currentSortColumn == null || currentSortColumn.DataPropertyName == "復旧日時" && currentSortDirection == ListSortDirection.Ascending)
                {
                    if (dgvLog.RowCount > 0)
                    {
                        dgvLog.FirstDisplayedScrollingRowIndex = dgvLog.RowCount - 1;
                    }
                }
            }
        }

        private void StartMonitoring()
        {
            cts = new CancellationTokenSource();
            txtStartTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            txtEndTime.Text = "";

            int interval = int.Parse(cmbInterval.Text);
            int timeout = int.Parse(cmbTimeout.Text);

            UpdateUiState("Running");
            uiUpdateTimer.Start();

            // ログのソート状態をリセット
            ResetLogSortIndicators();

            Action<DisruptionLogItem> logAction = AddDisruptionLogItem;

            foreach (var item in monitorList.Where(i => !string.IsNullOrEmpty(i.対象アドレス) && i.順番 > 0))
            {
                item.送信間隔ms = interval;
                item.タイムアウトms = timeout;
                item.ResetStatistics();
                Task.Run(() => RunPingLoopAsync(item, cts.Token, logAction));
            }

            if (dgvMonitor != null)
            {
                dgvMonitor.AllowUserToDeleteRows = false;
            }
        }

        private void StopMonitoring()
        {
            if (cts != null)
            {
                uiUpdateTimer.Stop();
                cts.Cancel();
                cts.Dispose();
                cts = null;
                txtEndTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                UpdateUiState("Stopped");
                
                if (btnPingStart != null)
                {
                    btnPingStart.Enabled = false; // 無効化 (グレーアウト)
                }
            }
        }

        private void ClearVeiw()
        {
            foreach (var item in monitorList)
            {
                item.ResetStatistics();
            }

            disruptionLogList.Clear();

            txtStartTime.Text = "";
            txtEndTime.Text = "";

            monitorList.ResetBindings();
            UpdateUiState("Initial");

            // ログのソートインジケータもリセット
            ResetLogSortIndicators();
        }

        private async Task RunPingLoopAsync(PingMonitorItem item, CancellationToken token, Action<DisruptionLogItem> logAction)
        {
            using (var ping = new Ping())
            {
                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        PingReply reply = await ping.SendPingAsync(item.対象アドレス, item.タイムアウトms);
                        item.UpdateStatistics(reply.RoundtripTime, reply.Status == IPStatus.Success, logAction);
                        await Task.Delay(item.送信間隔ms, token);
                    }
                    catch (TaskCanceledException)
                    {
                        break;
                    }
                    catch (Exception)
                    {
                        item.UpdateStatistics(0, false, logAction);
                        await Task.Delay(item.送信間隔ms, token);
                    }
                }
            }
            this.Invoke((MethodInvoker)delegate { dgvMonitor.Refresh(); });
        }

        private void InitializeCustomComponents()
        {
            this.Text = "Pings B版";
            this.Size = new Size(1300, 600); // 画面サイズを広めに設定
            // this.MinimumSize = new Size(900, 400);
            this.StartPosition = FormStartPosition.CenterScreen;

            // UI更新タイマーの初期化
            uiUpdateTimer = new System.Windows.Forms.Timer();
            uiUpdateTimer.Interval = UiUpdateInterval;
            uiUpdateTimer.Tick += UiUpdateTimer_Tick;

            // 上部パネル (日時、設定)
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(10) };
            this.Controls.Add(topPanel);

            // 開始日時/終了日時
            Label lblStart = new Label { Text = "開始日時", Location = new Point(10, 5), AutoSize = true };
            txtStartTime = new TextBox { Location = new Point(10, 23), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };
            Label lblEnd = new Label { Text = "終了日時", Location = new Point(170, 5), AutoSize = true };
            txtEndTime = new TextBox { Location = new Point(170, 23), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };

            topPanel.Controls.Add(lblStart);
            topPanel.Controls.Add(txtStartTime);
            topPanel.Controls.Add(lblEnd);
            topPanel.Controls.Add(txtEndTime);

            // 送信間隔
            Label lblInterval = new Label { Text = "送信間隔 [ms]", Location = new Point(350, 5), AutoSize = true };
            cmbInterval = new ComboBox { Location = new Point(350, 23), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbInterval.Items.AddRange(new object[] { "100", "500", "1000", "2000" });
            cmbInterval.SelectedIndex = 1;

            // タイムアウト
            Label lblTimeout = new Label { Text = "タイムアウト [ms]", Location = new Point(480, 5), AutoSize = true };
            cmbTimeout = new ComboBox { Location = new Point(480, 23), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTimeout.Items.AddRange(new object[] { "500", "1000", "2000", "5000" });
            cmbTimeout.SelectedIndex = 1;

            topPanel.Controls.Add(lblInterval);
            topPanel.Controls.Add(cmbInterval);
            topPanel.Controls.Add(lblTimeout);
            topPanel.Controls.Add(cmbTimeout);

            // タブコントロール
            tabControl = new TabControl { Dock = DockStyle.Fill };
            this.Controls.Add(tabControl);

            // 監視統計タブ
            TabPage statsPage = new TabPage("監視統計");
            dgvMonitor = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                BackgroundColor = SystemColors.ControlLightLight,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            statsPage.Controls.Add(dgvMonitor);
            tabControl.Controls.Add(statsPage);

            // 障害イベントログタブ
            TabPage logPage = new TabPage("障害イベントログ");
            dgvLog = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                ReadOnly = true,
                BackgroundColor = SystemColors.ControlLightLight,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToDeleteRows = false,
                AllowUserToAddRows = false
            };
            logPage.Controls.Add(dgvLog);
            tabControl.Controls.Add(logPage);

            // 下部パネル (ボタン)
            Panel bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            this.Controls.Add(bottomPanel);

            btnPingStart = new Button { Text = "Ping開始", Location = new Point(10, 5), Width = 80 };
            btnStop = new Button { Text = "停止", Location = new Point(100, 5), Width = 80 };
            btnClear = new Button { Text = "クリア", Location = new Point(190, 5), Width = 80 };
            btnSave = new Button { Text = "Ping結果保存", Location = new Point(280, 5), Width = 110 };
            btnExit = new Button { Text = "終了", Location = new Point(this.ClientSize.Width - 90, 5), Width = 80, Anchor = AnchorStyles.Right };

            bottomPanel.Controls.Add(btnPingStart);
            bottomPanel.Controls.Add(btnStop);
            bottomPanel.Controls.Add(btnClear);
            bottomPanel.Controls.Add(btnSave);
            bottomPanel.Controls.Add(btnExit);

            // イベントハンドラの設定
            btnPingStart.Click += BtnPingStart_Click;
            btnStop.Click += BtnStop_Click;
            btnClear.Click += BtnClear_Click;
            btnExit.Click += BtnExit_Click;
            btnSave.Click += BtnSaveResult_Click;

            // コントロールのZオーダー調整 (一番上にタブ、一番下にメニュー)
            tabControl.BringToFront();
            bottomPanel.SendToBack();
            topPanel.SendToBack();
            if (this.MainMenuStrip != null)
            {
                this.MainMenuStrip.SendToBack();
            }
        }


        private void SetupDataGridViewColumns()
        {
            dgvMonitor.AllowUserToAddRows = true;
            dgvMonitor.ReadOnly = false;
            dgvMonitor.Columns.Clear();

            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "ステータス", HeaderText = "ｽﾃｰﾀｽ", Width = 60, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "順番", HeaderText = "順番", Width = 50, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "対象アドレス", HeaderText = "対象アドレス", Width = 120, ReadOnly = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Host名", HeaderText = "Host名", Width = 120, ReadOnly = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "送信回数", HeaderText = "送信回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗回数", HeaderText = "失敗回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "連続失敗回数", HeaderText = "連続失敗回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "連続失敗時間s", HeaderText = "連続失敗時間[mm:ss]", Width = 130, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最大失敗時間s", HeaderText = "最大失敗時間[mm:ss]", Width = 130, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "時間ms", HeaderText = "時間[ms]", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "平均値ms", HeaderText = "平均値[ms]", Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" }, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最小値ms", HeaderText = "最小値[ms]", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最大値ms", HeaderText = "最大値[ms]", Width = 80, ReadOnly = true });

            dgvMonitor.CellFormatting += (s, e) =>
            {
                if (e.ColumnIndex == 0)
                {
                    if (dgvMonitor.Rows[e.RowIndex].DataBoundItem is PingMonitorItem item)
                    {
                        Color backColor;
                        Color selectionColor;
                        if (item.ステータス == "Down") { backColor = Color.MistyRose; selectionColor = Color.Red; }
                        else if (item.ステータス == "OK") { backColor = Color.LightGreen; selectionColor = Color.Green; }
                        else if (item.ステータス == "復旧") { backColor = Color.LightSkyBlue; selectionColor = Color.Blue; }
                        else { backColor = SystemColors.Window; selectionColor = SystemColors.Highlight; }
                        e.CellStyle.BackColor = backColor;
                        e.CellStyle.SelectionBackColor = selectionColor;
                    }
                }
                else
                {
                    e.CellStyle.BackColor = SystemColors.Window;
                    e.CellStyle.SelectionBackColor = SystemColors.Highlight;
                }
            };

            // ★障害イベントログ用カラム設定 (ソート対応のための変更)★
            // List<DisruptionLogItem> を基に SortableBindingList を初期化
            disruptionLogList = new SortableBindingList<DisruptionLogItem>(new List<DisruptionLogItem>());
            dgvLog.DataSource = disruptionLogList;

            // ★追加: カラムヘッダーのクリックイベントをフック★
            dgvLog.ColumnHeaderMouseClick += dgvLog_ColumnHeaderMouseClick;

            dgvLog.Columns.Clear();
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "対象アドレス", HeaderText = "対象アドレス", Width = 120 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Host名", HeaderText = "Host名", Width = 120 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down開始日時", HeaderText = "Down開始日時", Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "yyyy/MM/dd HH:mm:ss" } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧日時", HeaderText = "復旧日時", Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "yyyy/MM/dd HH:mm:ss" } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗時間mmss", HeaderText = "失敗時間[mm:ss]", Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });

            // Down前の統計
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前平均ms", HeaderText = "Down前平均[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最小ms", HeaderText = "Down前最小[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最大ms", HeaderText = "Down前最大[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });

            // 復旧後の統計
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後平均ms", HeaderText = "復旧後平均[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最小ms", HeaderText = "復旧後最小[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最大ms", HeaderText = "復旧後最大[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });

            // 初期ソートは復旧日時（昇順）
            // DataGridView の初期表示ではソートインジケータは設定しない
        }

        // ★新規追加: dgvLog のソートロジック★
        private void dgvLog_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewColumn newColumn = dgvLog.Columns[e.ColumnIndex];

            // 現在のソート方向を決定
            ListSortDirection direction;
            if (currentSortColumn != null && newColumn.DataPropertyName == currentSortColumn.DataPropertyName)
            {
                // 同じカラムを再度クリックした場合、ソート方向を反転させる
                direction = (currentSortDirection == ListSortDirection.Ascending) ?
                            ListSortDirection.Descending :
                            ListSortDirection.Ascending;
            }
            else
            {
                // 新しいカラムをクリックした場合、昇順でソートを開始する
                direction = ListSortDirection.Ascending;
            }

            // PropertyDescriptor を取得
            PropertyDescriptor prop = TypeDescriptor.GetProperties(typeof(DisruptionLogItem))
                                    .Find(newColumn.DataPropertyName, true);

            if (prop != null)
            {
                // 既存のソートインジケータをリセット
                if (currentSortColumn != null)
                {
                    currentSortColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                }

                // ソートを実行
                disruptionLogList.Sort(prop, direction);

                // ソートインジケータを設定
                newColumn.HeaderCell.SortGlyphDirection = (direction == ListSortDirection.Ascending) ?
                                                            SortOrder.Ascending :
                                                            SortOrder.Descending;

                // ソート状態を保存
                currentSortColumn = newColumn;
                currentSortDirection = direction;
            }
        }


    }
}