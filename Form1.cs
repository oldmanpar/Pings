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
using System.Diagnostics;
using System.Collections.Concurrent;

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

    // ★追加: 障害中の失敗回数を保持します。★
    public int 失敗回数 { get; set; }

    // ★既にあるプロパティ★
    // ★追加: 連続失敗回数をint型で保持します。★
    public int 連続失敗回数 { get; set; }

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

        // 計算用: CURRENT_SESSION の Up 時間
        private long _currentSessionUpTimeMs = 0;
        private int _currentSessionSuccessCount = 0;
        private long _currentSessionMin = 0;
        private long _currentSessionMax = 0;

        // Down時の統計スナップショット
        private bool _isCurrentlyDown = false;
        private double _snapAvg = 0.0;
        private long _snapMin = 0;
        private long _snapMax = 0;

        // ★追加: 現在進行中の障害で発生した失敗回数（ログ化用）★
        private int _currentDisruptionFailureCount = 0;

        // UIからの入力値
        public int 送信間隔ms { get; set; } = 500;
        public int タイムアウトms { get; set; } = 1000;

        // 追加: Traceroute 対象フラグ
        public bool Trace { get; set; } = false;

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

            _currentSessionUpTimeMs = 0;
            _currentSessionSuccessCount = 0;
            _currentSessionMin = 0;
            _currentSessionMax = 0;

            _continuousDownStartTime = null;
            _maxDisruptionDuration = TimeSpan.Zero;
            IsUp = false;
            HasBeenDown = false;
            _isCurrentlyDown = false;
            _activeLogItem = null;

            // 表示用項目も初期化（起動直後と同じ状態にする）
            ステータス = "";                 // ステータスを空にする（起動時と同じ見た目）
            連続失敗時間s = "";             // 表示用時間文字列をクリア
            最大失敗時間s = "";             // 表示用最大時間文字列をクリア

            // ★リセット: 現在の障害中失敗回数もリセット★
            _currentDisruptionFailureCount = 0;
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

                            // ★障害中の失敗回数をログにセット★
                            失敗回数 = _currentDisruptionFailureCount,

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
                    _currentSessionUpTimeMs = 0;
                    _currentSessionSuccessCount = 0;
                    _currentSessionMin = 0;
                    _currentSessionMax = 0;

                    // ★リセット: 障害中カウンタをクリア★
                    _currentDisruptionFailureCount = 0;

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
                _currentSessionSuccessCount++;
                _currentSessionUpTimeMs += rtt;

                // 最小値/最大値の更新
                if (_currentSessionSuccessCount == 1)
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
                平均値ms = (double)_currentSessionUpTimeMs / _currentSessionSuccessCount;
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
                    // Down開始イベント: DownになったらアクティBLOGをクリアし、更新を終了する
                    _activeLogItem = null;

                    // Down開始イベント: 現在の統計をスナップショットとして保存
                    _snapAvg = 平均値ms;
                    _snapMin = 最小値ms;
                    _snapMax = 最大値ms;

                    _continuousDownStartTime = DateTime.Now;
                    _isCurrentlyDown = true;

                    // 障害カウントを開始（最初は0だったのでここで初期化は不要だが明示的に）
                    _currentDisruptionFailureCount = 0;
                }

                // Down継続中も毎回実行するべき処理
                失敗回数++;
                // ★障害中カウントをインクリメント★
                _currentDisruptionFailureCount++;

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

        private Button btnTraceroute; // 追加: Traceroute実行ボタン
        private TabPage traceroutePage; // 追加: Traceroute タブ

        // 変更: 単一 TextBox -> 複数列表示用パネル + TextBox 持ち分
        private TableLayoutPanel traceroutePanel;
        private Dictionary<string, TextBox> tracerouteTextBoxes;

        private SemaphoreSlim tracerouteSemaphore = new SemaphoreSlim(4); // 追加: 同時実行制限（任意で調整）

        private Button btnSaveTraceroute; // 追加: Traceroute 出力保存ボタン
        private Button btnClearTraceroute; // 追加: Traceroute 出力クリアボタン
        private Button btnStopTraceroute; // 追加: Traceroute 停止ボタン

        // Traceroute専用キャンセルソース（Ping監視とは独立）
        private CancellationTokenSource tracerouteCts;

        // 追加: フィールド（他の private フィールド群の近くに追加）
        private volatile bool _isTracerouteRunning = false;

        // 追加: メニューでの自動保存チェックアイテム参照
        private ToolStripMenuItem mnuAutoSaveAllPing;

        // 追加: メニューでの自動保存チェックアイテム参照
        private ToolStripMenuItem mnuAutoSaveTraceroute;

        // 保存時の1ターゲットあたりの表示固定幅（px）。要件3に対応。
        private const int TracerouteColumnWidth = 480;

        // 追加フィールド（クラス内、他の private フィールド群の近くに追加）
        private readonly object _allPingLogLock = new object();

        // 高速ロードのための再利用 Regex（空白連続を1個に置換）
        private static readonly System.Text.RegularExpressions.Regex _wsRegex =
            new System.Text.RegularExpressions.Regex(@"\s+", System.Text.RegularExpressions.RegexOptions.Compiled);

        // -------------------------
        // Traceroute の進捗管理フィールド
        // -------------------------
        // 各アドレスの完了状態を保持（Run 中は null でない）
        private ConcurrentDictionary<string, bool> _tracerouteCompletion;
        // ユーザーによって停止メッセージを既に挿入したかを保持
        private ConcurrentDictionary<string, bool> _tracerouteStoppedByUser;

        // 追加: Ping停止後の編集許可制御フラグ
        private bool _allowEditAfterStop = true;

        public Form1()
        {
            // ★修正箇所: デザイナー生成コードを呼び出す★
            InitializeComponent();

            // ★修正箇所: メニューを先に初期化しておく（Dock順の問題を回避）★
            SetupMenuStrip();

            // ★修正箇所: カスタムコンポーネントの初期化と配置★
            InitializeCustomComponents();

            // ★修正箇所: DataGridView のカラム設定を行う (InitializeCustomComponentsで dgvMonitor が生成された後)★
            SetupDataGridViewColumns();

            // (SetupMenuStrip は上で先に呼んでいるためここでは呼ばない)

            monitorList = new BindingList<PingMonitorItem>();
            // 追加: 変更検知と UI イベントをフック
            monitorList.ListChanged += MonitorList_ListChanged;
            dgvMonitor.DataSource = monitorList;

            // DataGridView のユーザーによる削除や行確定を補足
            dgvMonitor.UserDeletedRow += DgvMonitor_UserDeletedRow;
            dgvMonitor.RowValidated += DgvMonitor_RowValidated;

            dgvMonitor.DefaultValuesNeeded += DgvMonitor_DefaultValuesNeeded;
            this.FormClosing += Form1_FormClosing;

            // 初期データの追加
            monitorList.Add(new PingMonitorItem(_nextIndex++, "127.0.0.1", "loopback"));
            monitorList.Add(new PingMonitorItem(_nextIndex++, "8.8.8.8", "Google DNS1"));
            monitorList.Add(new PingMonitorItem(_nextIndex++, "8.8.4.4", "Google DNS2"));

            // ★修正箇所: UIを初期状態にする★
            UpdateUiState("Initial");
        }

        // 5) UpdateUiState を少し修正し、Running 中は Trace 列のみ編集可能にする
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

            // デフォルト: DataGridView を編集可能にして、列ごとに ReadOnly を設定する
            dgvMonitor.ReadOnly = false;
            dgvMonitor.AllowUserToAddRows = isEditable;

            // 各列の ReadOnly を状況に応じて調整
            foreach (DataGridViewColumn col in dgvMonitor.Columns)
            {
                if (state == "Initial")
                {
                    // Initial: 対象アドレス/Host名 は編集可、順番は readonly、Trace は編集可
                    if (col.DataPropertyName == "対象アドレス" || col.DataPropertyName == "Host名" || col.DataPropertyName == "Trace")
                        col.ReadOnly = false;
                    else
                        col.ReadOnly = true;
                }
                else if (state == "Running")
                {
                    // Running: 変更できるのは Trace のみ（チェックの切替）
                    if (col.DataPropertyName == "Trace")
                        col.ReadOnly = false;
                    else
                        col.ReadOnly = true;
                }
                else // Stopped
                {
                    // 変更: Ping停止後は対象アドレス/Host名を編集不可にする（ただし
                    //       _allowEditAfterStop が true の場合のみ編集可とする）
                    if (col.DataPropertyName == "Trace")
                    {
                        // Trace 列は Stopped 時も切替可能（従来仕様維持）
                        col.ReadOnly = false;
                    }
                    else if (col.DataPropertyName == "対象アドレス" || col.DataPropertyName == "Host名")
                    {
                        // Stop直後は編集不可。Ping結果保存またはPing結果クリアで許可される。
                        col.ReadOnly = !_allowEditAfterStop;
                    }
                    else
                    {
                        col.ReadOnly = true;
                    }
                }
            }

            // Traceroute ボタン制御（既存） + 停止ボタン
            if (btnTraceroute != null)
            {
                if (state != "Running")
                {
                    btnTraceroute.Enabled = false;
                }
                UpdateTracerouteButtons();
            }

            // 既存のスイッチ文でボタン有効/無効を設定（省略せずそのまま保持してください）
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
                    // Ping停止直後はクリア/保存を有効にし、編集許可は _allowEditAfterStop で制御
                    btnClear.Enabled = true;
                    btnSave.Enabled = true;
                    break;
            }
        }

        // 変更済: Traceroute ボタン状態更新メソッド
        // UpdateTracerouteButtons を修正：自動保存オン時に手動保存ボタンを無効にする
        private void UpdateTracerouteButtons()
        {
            if (btnTraceroute == null || btnSaveTraceroute == null || btnClearTraceroute == null || btnStopTraceroute == null) return;

            if (this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)UpdateTracerouteButtons);
                return;
            }

            // Traceroute 実行は Trace チェックが一つでもあれば有効（Ping 実行有無に依存しない）
            bool hasCheckedTrace = (monitorList != null) && monitorList.Any(i => i.Trace && !string.IsNullOrWhiteSpace(i.対象アドレス));
            btnTraceroute.Enabled = hasCheckedTrace && !_isTracerouteRunning;

            // Traceroute 出力の有無で保存/クリアを制御（実行中は無効）
            bool hasTracerouteOutput = tracerouteTextBoxes != null && tracerouteTextBoxes.Values.Any(t => !string.IsNullOrEmpty(t.Text));

            // 自動保存が有効なときは手動保存ボタンを無効にする
            bool autoSave = (mnuAutoSaveTraceroute != null && mnuAutoSaveTraceroute.Checked);

            btnSaveTraceroute.Enabled = !_isTracerouteRunning && hasTracerouteOutput && !autoSave;
            btnClearTraceroute.Enabled = !_isTracerouteRunning && hasTracerouteOutput; // クリアは自動保存時でも有効にしておく

            // 停止ボタンは実行中のみ有効
            btnStopTraceroute.Enabled = _isTracerouteRunning;
        }
        private void SetupMenuStrip()
        {
            // 既に MainMenuStrip が存在する場合はそれを再利用する
            if (this.MainMenuStrip != null)
            {
                this.MainMenuStrip.Dock = DockStyle.Top;
                this.MainMenuStrip.Padding = new Padding(0);
                this.MainMenuStrip.Margin = new Padding(0);
                // フォームの Controls に未登録なら登録する
                if (!this.Controls.Contains(this.MainMenuStrip))
                {
                    this.Controls.Add(this.MainMenuStrip);
                }
                // メニューは先頭に確保
                try { this.Controls.SetChildIndex(this.MainMenuStrip, 0); } catch { }
                return;
            }

            var menuStrip = new MenuStrip
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(0),
                Margin = new Padding(0),
                RenderMode = ToolStripRenderMode.System
            };

            var fileMenu = new ToolStripMenuItem("ファイル");
            var saveAddressItem = new ToolStripMenuItem("対象アドレス保存");
            saveAddressItem.Click += BtnSaveAddress_Click;
            fileMenu.DropDownItems.Add(saveAddressItem);

            var loadAddressItem = new ToolStripMenuItem("対象アドレス読込");
            loadAddressItem.Click += BtnLoadAddress_Click;
            fileMenu.DropDownItems.Add(loadAddressItem);

            fileMenu.DropDownItems.Add(new ToolStripSeparator());

            var exitItem = new ToolStripMenuItem("終了");
            exitItem.Click += BtnExit_Click;
            fileMenu.DropDownItems.Add(exitItem);

            menuStrip.Items.Add(fileMenu);

            var optionsMenu = new ToolStripMenuItem("オプション");
            // 5) 全Ping結果を自動保存するチェックボックス
            mnuAutoSaveAllPing = new ToolStripMenuItem("全Ping結果を保存する") { CheckOnClick = true };
            optionsMenu.DropDownItems.Add(mnuAutoSaveAllPing);

            // 追加: Traceroute 自動保存チェックボックス（初期値オフ）
            mnuAutoSaveTraceroute = new ToolStripMenuItem("Trace結果の自動保存") { CheckOnClick = true, Checked = true };
            mnuAutoSaveTraceroute.CheckedChanged += MnuAutoSaveTraceroute_CheckedChanged;
            optionsMenu.DropDownItems.Add(mnuAutoSaveTraceroute);

            menuStrip.Items.Add(optionsMenu);

            var helpMenu = new ToolStripMenuItem("ヘルプ");
            var versionItem = new ToolStripMenuItem("バージョン情報");
            versionItem.Click += VersionItem_Click;
            helpMenu.DropDownItems.Add(versionItem);
            menuStrip.Items.Add(helpMenu);

            this.Controls.Add(menuStrip);
            this.MainMenuStrip = menuStrip;

            // ここでは一度だけインデックスを調整（不要なら削除可）
            try { this.Controls.SetChildIndex(menuStrip, 0); } catch { }
        }

        private void BtnPingStart_Click(object sender, EventArgs e)
        {
            // 1. 既に監視が開始されていたら、一度停止させる（再開時の安全策）
            StopMonitoring();
            ClearVeiw();

            // 2. DataGridViewの編集内容を確定
            dgvMonitor.EndEdit();

            // 3. 空欄行を除外
            var validItems = monitorList
                .Where(i => !string.IsNullOrWhiteSpace(i.対象アドレス) || !string.IsNullOrWhiteSpace(i.Host名))
                .ToList();

            monitorList.ListChanged -= MonitorList_ListChanged;
            try
            {
                monitorList.Clear();
                foreach (var item in validItems)
                {
                    monitorList.Add(item);
                }
            }
            finally
            {
                monitorList.ListChanged += MonitorList_ListChanged;
            }

            // 4. 次の新規行の「順番」を決定
            if (monitorList.Any())
            {
                _nextIndex = monitorList.Max(i => i.順番) + 1;
            }
            else
            {
                _nextIndex = 1;
            }

            // 5. Ping監視の開始
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

            // クリア操作後は編集を許可する
            _allowEditAfterStop = true;
            UpdateUiState("Initial");

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
                // 順番列は現在インデックス 2 (0:ステータス,1:Trace,2:順番)
                e.Row.Cells[2].Value = _nextIndex;
                _nextIndex++;
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
            // Handle が作成されているかを正しいプロパティ名でチェック
            if (dgvMonitor != null && dgvMonitor.IsHandleCreated)
            {
                try
                {
                    dgvMonitor.Refresh();
                }
                catch
                {
                    // UI スレッドが破棄されているなどの例外を無視
                }
            }

            // dgvLog が null でないかを確認して Refresh を呼ぶ
            if (dgvLog != null && dgvLog.IsHandleCreated)
            {
                try
                {
                    dgvLog.Refresh();
                }
                catch
                {
                    // 無視
                }
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

        // 変更: ファイルを一括読み込みせずに StreamReader で逐次処理することでメモリ/速度改善
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
                        var newMonitorItems = new List<PingMonitorItem>();
                        int index = 1;

                        // Shift-JIS (Encoding 932) でファイルを読み込みます。
                        System.Text.Encoding encoding = System.Text.Encoding.GetEncoding(932);

                        var allLines = new List<string>();
                        using (StreamReader sr = new StreamReader(ofd.FileName, encoding))
                        {
                            string line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                allLines.Add(line);
                            }
                        }

                        // ヘッダー検出とスキップ
                        bool isSavedCsvFormat = false;
                        if (allLines.Any())
                        {
                            string headerCandidate = allLines[0].Trim().Replace(" ", "");
                            if (headerCandidate.Equals("対象アドレス,Host名", StringComparison.OrdinalIgnoreCase))
                            {
                                allLines.RemoveAt(0);
                                isSavedCsvFormat = true;
                            }
                        }

                        // 各行を解析して監視対象リストを構築
                        foreach (string currentLine in allLines)
                        {
                            string trimmedLine = currentLine.Trim();
                            if (string.IsNullOrWhiteSpace(trimmedLine)) continue;

                            if (trimmedLine.StartsWith("[") || trimmedLine.StartsWith("#") || trimmedLine.StartsWith(";") || trimmedLine.StartsWith("'"))
                            {
                                continue;
                            }

                            string address;
                            string hostName;

                            if (isSavedCsvFormat)
                            {
                                string[] parts = currentLine.Split(new[] { ',' }, 2).Select(p => p.Trim()).ToArray();
                                address = parts[0];
                                hostName = parts.Length > 1 ? parts[1] : "";
                            }
                            else
                            {
                                int separatorIndex = -1;
                                for (int i = 0; i < currentLine.Length; i++)
                                {
                                    if ((currentLine[i] == ' ' || currentLine[i] == '\t') && currentLine.Substring(0, i).Trim().Length > 0)
                                    {
                                        separatorIndex = i;
                                        break;
                                    }
                                }

                                if (separatorIndex != -1)
                                {
                                    address = currentLine.Substring(0, separatorIndex).Trim();
                                    hostName = currentLine.Substring(separatorIndex).Trim();
                                    hostName = System.Text.RegularExpressions.Regex.Replace(hostName, @"\s+", " ").Trim();
                                }
                                else if (currentLine.Contains(','))
                                {
                                    string[] parts = currentLine.Split(new[] { ',' }, 2).Select(p => p.Trim()).ToArray();
                                    address = parts[0];
                                    hostName = parts.Length > 1 ? parts[1] : "";
                                }
                                else
                                {
                                    address = trimmedLine;
                                    hostName = "";
                                }
                            }

                            if (string.IsNullOrEmpty(address)) continue;

                            newMonitorItems.Add(new PingMonitorItem(index++, address, hostName));
                        }

                        // リストを差し替え（UI更新を最小化）
                        var newBinding = new BindingList<PingMonitorItem>(newMonitorItems);

                        try
                        {
                            if (monitorList != null)
                            {
                                monitorList.ListChanged -= MonitorList_ListChanged;
                            }
                        }
                        catch { /* ignore */ }

                        if (dgvMonitor != null) dgvMonitor.SuspendLayout();

                        monitorList = newBinding;
                        monitorList.ListChanged += MonitorList_ListChanged;
                        dgvMonitor.DataSource = monitorList;

                        if (dgvMonitor != null)
                        {
                            dgvMonitor.ResumeLayout();
                            dgvMonitor.Refresh();
                        }

                        _nextIndex = index;
                        UpdateUiState("Initial");

                        // 変更: 読み込んだ件数を表示するメッセージへ
                        MessageBox.Show($"監視対象アドレスを　{newMonitorItems.Count}件　読み込みました。", "読込完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            // 実行ファイルと同じフォルダへ保存（ダイアログは表示しない）
            string folder = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(folder, $"Ping_Result_{DateTime.Now:yyyyMMdd}.csv");

            try
            {
                // 追記モードで保存（再実行時は追記される）
                // Shift-JIS (Encoding.GetEncoding(932)) でファイルに書き込みます。
                using (StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding(932)))
                {
                    // 区切りヘッダ（追記したときに分かりやすくする）
                    sw.WriteLine("================================================================");
                    sw.WriteLine($"Saved: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                    sw.WriteLine();

                    // --- 1. ヘッダー情報 ---
                    sw.WriteLine($"開始日時　：　{txtStartTime.Text}");
                    sw.WriteLine($"終了日時　：　{txtEndTime.Text}");
                    sw.WriteLine($"送信間隔　：　{cmbInterval.Text} [ms]     タイムアウト　：　{cmbTimeout.Text} [ms]");
                    sw.WriteLine(); // 空行

                    // --- 2. 監視統計データ (dgvMonitorの内容) ---
                    sw.WriteLine("--- 監視統計 ---");

                    // ヘッダー行
                    string header = "ｽﾃｰﾀｽ,順番,対象アドレス,Host名,送信回数,失敗回数,連続失敗回数,連続失敗時間[hh:mm:ss],最大失敗時間[hh:mm:ss],時間[ms],平均値[ms],最小値[ms],最大値[ms]";
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
                        // ヘッダー行（失敗回数を追加）
                        string logHeader = "対象アドレス,Host名,Down開始日時,復旧日時,失敗回数,失敗時間[hh:mm:ss]," +
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
                                logItem.失敗回数,
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
                    sw.WriteLine();
                }

                MessageBox.Show($"監視結果をファイルに保存しました（追記モード）：\n{filePath}", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 保存後は対象アドレス/Host名の編集を許可する
                _allowEditAfterStop = true;
                UpdateUiState("Stopped");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ファイルの保存中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (btnPingStart != null)
            {
                btnPingStart.Enabled = true; // 有効化
            }
        }

        // monitorList の変更（追加・削除・リセット）を受けて順番を再計算する
        private void MonitorList_ListChanged(object sender, ListChangedEventArgs e)
        {
            if (e.ListChangedType == ListChangedType.ItemAdded ||
                e.ListChangedType == ListChangedType.ItemDeleted ||
                e.ListChangedType == ListChangedType.Reset)
            {
                RecalculateOrderNumbers();
            }
        }

        // DataGridView で行を削除したとき（DELキー等）
        private void DgvMonitor_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            RecalculateOrderNumbers();
        }

        // メニューのチェック変更ハンドラ（シンプルにUIを更新）
        private void MnuAutoSaveTraceroute_CheckedChanged(object sender, EventArgs e)
        {
            // チェック切替時に手動保存ボタンの有効/無効を即時反映
            UpdateTracerouteButtons();
        }

        // 自動保存の共通処理（追記）。BtnSaveTraceroute_Click と同じファイル名規則で保存するが UI はクリアしない。
        private void SaveTracerouteOutputsAutoAppend()
        {
            if (tracerouteTextBoxes == null || tracerouteTextBoxes.Count == 0) return;

            string folder = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                foreach (var kv in tracerouteTextBoxes)
                {
                    string address = kv.Key;
                    string content = kv.Value.Text;
                    if (string.IsNullOrEmpty(content)) continue;

                    // Host名を monitorList から取得（同一アドレスが複数ある場合は最初のものを使用）
                    string hostName = monitorList?
                        .FirstOrDefault(i => string.Equals(i.対象アドレス, address, StringComparison.OrdinalIgnoreCase))
                        ?.Host名 ?? "";

                    string safeAddress = SanitizeFileName(address);
                    string safeHost = SanitizeFileName(hostName);
                    string fileName = Path.Combine(folder, $"Traceroute_result_{DateTime.Now:yyyyMMdd}_{safeAddress}_{safeHost}.log");

                    using (var sw = new StreamWriter(fileName, true, System.Text.Encoding.GetEncoding(932)))
                    {
                        sw.WriteLine("----------------------------------------------------------------");
                        sw.WriteLine($"Saved: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                        sw.WriteLine(content);
                        sw.WriteLine();
                    }
                }
            }
            catch (Exception ex)
            {
                // 自動保存の失敗はユーザーに通知する（必要ならログに落とす）
                try
                {
                    MessageBox.Show($"Traceroute 自動保存中にエラーが発生しました:\n{ex.Message}", "保存エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { /* UIが無い場面では無視 */ }
            }
        }

        // 行の編集確定後に呼んでおく（新規行コミットの補助）
        private void DgvMonitor_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            // 小コストなので常に再計算して差し支えありません
            RecalculateOrderNumbers();
        }

        // 実際に順番を詰めて UI を更新するメソッド
        private bool _isRecalculating = false;

        private void RecalculateOrderNumbers()
        {
            if (_isRecalculating) return; // 再入時は即座に抜ける
            _isRecalculating = true;

            // 再入/再帰防止：イベントを一時解除
            monitorList.ListChanged -= MonitorList_ListChanged;
            try
            {
                for (int i = 0; i < monitorList.Count; i++)
                {
                    monitorList[i].順番 = i + 1;
                }
                _nextIndex = monitorList.Count + 1;

                try
                {
                    if (dgvMonitor.IsHandleCreated)
                    {
                        dgvMonitor.Invoke((MethodInvoker)delegate
                        {
                            monitorList.ResetBindings();
                            dgvMonitor.Refresh();
                        });
                    }
                    else
                    {
                        monitorList.ResetBindings();
                    }
                }
                catch
                {
                    // UI スレッドが破棄されている等は無視
                }
            }
            finally
            {
                // 忘れずに再登録
                monitorList.ListChanged += MonitorList_ListChanged;
                _isRecalculating = false;
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
            // 監視開始時は停止直後の編集を許可しない
            _allowEditAfterStop = false;

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

            // 5) メニューの「全Ping結果を保存する」がオンのとき、実行ファイルの場所へ自動保存（追記）
            if (mnuAutoSaveAllPing != null && mnuAutoSaveAllPing.Checked)
            {
                try
                {
                    SaveAllPingResultsAutoAppend();
                }
                catch
                {
                    // 保存失敗はログや例外にせず無視（必要なら通知を追加）
                }
            }

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

        // All_Ping_Result_yyyymmdd.log に開始情報と簡易ヘッダーを追記する
        private void SaveAllPingResultsAutoAppend()
        {
            string folder = AppDomain.CurrentDomain.BaseDirectory;
            string fileName = Path.Combine(folder, $"All_Ping_Result_{DateTime.Now:yyyyMMdd}.log");
            using (var sw = new StreamWriter(fileName, true, System.Text.Encoding.GetEncoding(932)))
            {
                sw.WriteLine("========================================");
                sw.WriteLine($"開始日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                sw.WriteLine($"送信間隔: {cmbInterval?.Text} ms  タイムアウト: {cmbTimeout?.Text} ms");
                sw.WriteLine("--- 監視対象 ---");
                foreach (var item in monitorList.OrderBy(i => i.順番))
                {
                    sw.WriteLine($"{item.順番}: {item.対象アドレス}  ({item.Host名})");
                }
                sw.WriteLine();
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

                // 停止直後は対象アドレス/Host名の編集を禁止する
                _allowEditAfterStop = false;

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

            // Traceroute 関連もクリア（ユーザーが明示的にクリア/保存したときのみ呼ばれる想定）
            tracerouteTextBoxes?.Clear();
            if (traceroutePanel != null)
            {
                if (traceroutePanel.InvokeRequired)
                {
                    traceroutePanel.Invoke(new Action(() => traceroutePanel.Controls.Clear()));
                }
                else
                {
                    traceroutePanel.Controls.Clear();
                }
            }
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
            this.Size = new Size(1300, 700);
            this.StartPosition = FormStartPosition.CenterScreen;

            uiUpdateTimer = new System.Windows.Forms.Timer();
            uiUpdateTimer.Interval = UiUpdateInterval;
            uiUpdateTimer.Tick += UiUpdateTimer_Tick;

            // Top panel (日時、設定)
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(10) };
            Label lblStart = new Label { Text = "開始日時", Location = new Point(10, 5), AutoSize = true };
            txtStartTime = new TextBox { Location = new Point(10, 23), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };
            Label lblEnd = new Label { Text = "終了日時", Location = new Point(170, 5), AutoSize = true };
            txtEndTime = new TextBox { Location = new Point(170, 23), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };
            topPanel.Controls.Add(lblStart);
            topPanel.Controls.Add(txtStartTime);
            topPanel.Controls.Add(lblEnd);
            topPanel.Controls.Add(txtEndTime);

            Label lblInterval = new Label { Text = "送信間隔 [ms]", Location = new Point(350, 5), AutoSize = true };
            cmbInterval = new ComboBox { Location = new Point(350, 23), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbInterval.Items.AddRange(new object[] { "100", "500", "1000", "2000" });
            cmbInterval.SelectedIndex = 1;

            Label lblTimeout = new Label { Text = "タイムアウト [ms]", Location = new Point(480, 5), AutoSize = true };
            cmbTimeout = new ComboBox { Location = new Point(480, 23), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTimeout.Items.AddRange(new object[] { "500", "1000", "2000", "5000" });
            cmbTimeout.SelectedIndex = 1;

            topPanel.Controls.Add(lblInterval);
            topPanel.Controls.Add(cmbInterval);
            topPanel.Controls.Add(lblTimeout);
            topPanel.Controls.Add(cmbTimeout);

            // content panel
            Panel contentPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0) };

            // tabControl inside content
            tabControl = new TabControl { Dock = DockStyle.Fill };
            TabPage statsPage = new TabPage("監視統計");
            dgvMonitor = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                BackgroundColor = SystemColors.ControlLight,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            statsPage.Controls.Add(dgvMonitor);
            tabControl.Controls.Add(statsPage);

            TabPage logPage = new TabPage("障害イベントログ");
            dgvLog = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                ReadOnly = true,
                BackgroundColor = SystemColors.ControlLight,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToDeleteRows = false,
                AllowUserToAddRows = false
            };
            logPage.Controls.Add(dgvLog);
            tabControl.Controls.Add(logPage);

            traceroutePage = new TabPage("Traceroute");

            // Traceroute 用の TableLayoutPanel を用意し、実行時に列を追加していく設計
            traceroutePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                ColumnCount = 1,
                RowCount = 1
            };
            traceroutePage.Controls.Add(traceroutePanel);
            tabControl.Controls.Add(traceroutePage);

            contentPanel.Controls.Add(tabControl);
            contentPanel.Controls.Add(topPanel);

            // bottom panel - グループで分けて配置
            Panel bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 80, Margin = new Padding(0) };

            // Ping 関連グループ
            GroupBox gbPing = new GroupBox { Text = "Ping", Dock = DockStyle.None, Width = 460, Padding = new Padding(8), AutoSize = false };
            btnPingStart = new Button { Text = "Ping開始", Location = new Point(8, 22), Width = 100 };
            btnStop = new Button { Text = "Ping停止", Location = new Point(116, 22), Width = 100 };
            btnClear = new Button { Text = "Ping結果クリア", Location = new Point(224, 22), Width = 110 }; // 名前変更
            btnSave = new Button { Text = "Ping結果保存", Location = new Point(340, 22), Width = 110 };

            gbPing.Controls.Add(btnPingStart);
            gbPing.Controls.Add(btnStop);
            gbPing.Controls.Add(btnClear);
            gbPing.Controls.Add(btnSave);

            // Traceroute 関連グループ
            GroupBox gbTrace = new GroupBox { Text = "Traceroute", Dock = DockStyle.None, Width = 760, Padding = new Padding(8), AutoSize = false };
            btnTraceroute = new Button { Text = "Traceroute実行", Location = new Point(8, 22), Width = 120 };
            btnStopTraceroute = new Button { Text = "Traceroute停止", Location = new Point(136, 22), Width = 120 };
            btnClearTraceroute = new Button { Text = "Trace結果クリア", Location = new Point(264, 22), Width = 140 }; // 名前変更
            btnSaveTraceroute = new Button { Text = "Trace結果保存", Location = new Point(412, 22), Width = 140 }; // 名前変更

            gbTrace.Controls.Add(btnTraceroute);
            gbTrace.Controls.Add(btnStopTraceroute);
            gbTrace.Controls.Add(btnClearTraceroute);
            gbTrace.Controls.Add(btnSaveTraceroute);

            // 左側にグループボックスを流す FlowLayoutPanel を作成して、それぞれの枠が"囲み切る"ように配置
            var leftFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                AutoSize = false,
                Width = gbPing.Width + gbTrace.Width + 24,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8),
                Margin = new Padding(0)
            };
            leftFlow.Controls.Add(gbPing);
            leftFlow.Controls.Add(gbTrace);

            // 右側に終了ボタンを置くためのパネル（枠に被らないように右端に固定）
            var rightPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 120,
                Padding = new Padding(8)
            };

            // add exit button to the right (inside rightPanel)
            btnExit = new Button { Text = "終了", Anchor = AnchorStyles.Top | AnchorStyles.Right, Width = 80, Height = 30 };
            btnExit.Location = new Point(rightPanel.ClientSize.Width - btnExit.Width - 10, 20);
            // adjust location when resized
            rightPanel.Resize += (s, e) =>
            {
                btnExit.Location = new Point(Math.Max(8, rightPanel.ClientSize.Width - btnExit.Width - 10), 20);
            };
            rightPanel.Controls.Add(btnExit);

            // add flow and right panel to bottom panel
            bottomPanel.Controls.Add(rightPanel);
            bottomPanel.Controls.Add(leftFlow);

            // add to form
            if (this.MainMenuStrip != null && !this.Controls.Contains(this.MainMenuStrip))
            {
                this.Controls.Add(this.MainMenuStrip);
                this.MainMenuStrip.Dock = DockStyle.Top;
            }
            this.Controls.Add(contentPanel);
            this.Controls.Add(bottomPanel);

            // handlers
            btnPingStart.Click += BtnPingStart_Click;
            btnStop.Click += BtnStop_Click;
            btnClear.Click += BtnClear_Click;
            btnExit.Click += BtnExit_Click;
            btnSave.Click += BtnSaveResult_Click;
            btnTraceroute.Click += BtnTraceroute_Click;
            btnSaveTraceroute.Click += BtnSaveTraceroute_Click;
            btnClearTraceroute.Click += BtnClearTraceroute_Click;
            btnStopTraceroute.Click += BtnStopTraceroute_Click;

            // ensure menu spacing
            int menuHeight = 0;
            if (this.MainMenuStrip != null)
            {
                menuHeight = this.MainMenuStrip.PreferredSize.Height;
            }
            contentPanel.Padding = new Padding(0, menuHeight, 0, 0);

            // initial traceroute button state
            tracerouteTextBoxes = new Dictionary<string, TextBox>(StringComparer.OrdinalIgnoreCase);
            UpdateTracerouteButtons();
        }

        private void SetupDataGridViewColumns()
        {
            dgvMonitor.AllowUserToAddRows = true;
            dgvMonitor.ReadOnly = false;
            dgvMonitor.Columns.Clear();

            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "ステータス", HeaderText = "ｽﾃｰﾀｽ", Width = 60, ReadOnly = true });

            var traceCol = new DataGridViewCheckBoxColumn
            {
                DataPropertyName = "Trace",
                HeaderText = "Trace",
                Width = 60,
                TrueValue = true,
                FalseValue = false,
                ThreeState = false
            };
            dgvMonitor.Columns.Add(traceCol);

            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "順番", HeaderText = "順番", Width = 50, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "対象アドレス", HeaderText = "対象アドレス", Width = 120, ReadOnly = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Host名", HeaderText = "Host名", Width = 120, ReadOnly = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "送信回数", HeaderText = "送信回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗回数", HeaderText = "失敗回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "連続失敗回数", HeaderText = "連続失敗回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "連続失敗時間s", HeaderText = "連続失敗時間[hh:mm:ss]", Width = 130, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = @"hh\:mm\:ss" }, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最大失敗時間s", HeaderText = "最大失敗時間[hh:mm:ss]", Width = 130, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = @"hh\:mm\:ss" }, ReadOnly = true });
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

            disruptionLogList = new SortableBindingList<DisruptionLogItem>(new List<DisruptionLogItem>());
            dgvLog.DataSource = disruptionLogList;
            dgvLog.ColumnHeaderMouseClick += dgvLog_ColumnHeaderMouseClick;
            dgvLog.Columns.Clear();
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "対象アドレス", HeaderText = "対象アドレス", Width = 120 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Host名", HeaderText = "Host名", Width = 120 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down開始日時", HeaderText = "Down開始日時", Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "yyyy/MM/dd HH:mm:ss" } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧日時", HeaderText = "復旧日時", Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "yyyy/MM/dd HH:mm:ss" } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗回数", HeaderText = "失敗回数", Width = 80, ReadOnly = true });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗時間mmss", HeaderText = "失敗時間[hh:mm:ss]", Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = @"hh\:mm\:ss" } });

            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前平均ms", HeaderText = "Down前平均[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最小ms", HeaderText = "Down前最小[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最大ms", HeaderText = "Down前最大[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });

            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後平均ms", HeaderText = "復旧後平均[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最小ms", HeaderText = "復旧後最小[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最大ms", HeaderText = "復旧後最大[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });

            dgvMonitor.ColumnHeaderMouseDoubleClick -= DgvMonitor_ColumnHeaderMouseDoubleClick;
            dgvMonitor.ColumnHeaderMouseDoubleClick += DgvMonitor_ColumnHeaderMouseDoubleClick;

            // 追加: チェックボックスの値変更を検出してボタン状態を更新する
            dgvMonitor.CellValueChanged -= DgvMonitor_CellValueChanged;
            dgvMonitor.CellValueChanged += DgvMonitor_CellValueChanged;

            dgvMonitor.CurrentCellDirtyStateChanged -= DgvMonitor_CurrentCellDirtyStateChanged;
            dgvMonitor.CurrentCellDirtyStateChanged += DgvMonitor_CurrentCellDirtyStateChanged;

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

        // 4) ヘッダのダブルクリックイベントハンドラを追加
        private void DgvMonitor_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.ColumnIndex >= dgvMonitor.Columns.Count) return;

            var col = dgvMonitor.Columns[e.ColumnIndex];
            if (col.DataPropertyName == "Trace")
            {
                // ひとつでも未チェックがあれば全チェック、すべてチェックなら全解除
                bool anyUnchecked = monitorList.Any(i => !i.Trace);
                foreach (var item in monitorList)
                {
                    item.Trace = anyUnchecked;
                }

                // バインディング更新
                try {
                    monitorList.ResetBindings();
                }
                catch
                {
                    dgvMonitor.Refresh();
                }

                // 追加: Trace 列の一括切替後に Traceroute ボタンを更新
                UpdateTracerouteButtons();
            }
        }

        // ------------------- Traceroute ロジック -------------------
        private async void BtnTraceroute_Click(object sender, EventArgs e)
        {
            // チェックされた行のみを対象
            var targets = monitorList
                .Where(i => i.Trace && !string.IsNullOrWhiteSpace(i.対象アドレス))
                .Select(i => i.対象アドレス)
                .Distinct()
                .ToList();

            if (!targets.Any())
            {
                MessageBox.Show("Traceroute対象が選択されていません。Trace 列にチェックを入れてください。", "実行不可", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Traceroute 実行フラグを立て、ボタン状態を更新
            _isTracerouteRunning = true;
            UpdateTracerouteButtons();

            // Tracerouteタブを選択
            tabControl.SelectedTab = traceroutePage;

            // Determine whether to reuse existing textboxes (append) or recreate (clear)
            bool reuseExisting = tracerouteTextBoxes != null
                && tracerouteTextBoxes.Count > 0
                && tracerouteTextBoxes.Count == targets.Count
                && targets.All(t => tracerouteTextBoxes.ContainsKey(t));

            if (!reuseExisting)
            {
                // Clear and create new layout if targets differ
                tracerouteTextBoxes.Clear();
                traceroutePanel.Controls.Clear();

                // 列数を targets.Count に設定（固定幅の列を作る）
                int colCount = Math.Max(1, targets.Count);
                traceroutePanel.ColumnCount = colCount;
                traceroutePanel.RowCount = 1;
                traceroutePanel.ColumnStyles.Clear();
                traceroutePanel.AutoScroll = true;

                for (int i = 0; i < colCount; i++)
                {
                    // 要件3: 固定幅（折り返ししない）にするため Absolute を使用
                    traceroutePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, TracerouteColumnWidth));
                }

                // BtnTraceroute_Click 内の「各ターゲットに対して縦にラベル + TextBox を作る」ループを以下に置換してください。
                // 変更点: ラベルの Text を "対象アドレス_Host名" 形式にしています。

                for (int i = 0; i < targets.Count; i++)
                {
                    string address = targets[i];

                    var container = new Panel { Dock = DockStyle.Fill, Padding = new Padding(2), Width = TracerouteColumnWidth };

                    // MonitorList から Host名 を取得（同一アドレスが複数あれば先頭を使用）
                    string hostName = monitorList?
                        .FirstOrDefault(m => string.Equals(m.対象アドレス, address, StringComparison.OrdinalIgnoreCase))
                        ?.Host名 ?? "";

                    var lbl = new Label
                    {
                        // 要求どおり "対象アドレス_ホスト名" 形式で表示
                        Text = $"{address}_{hostName}",
                        Dock = DockStyle.Top,
                        Height = 18,
                        TextAlign = ContentAlignment.MiddleLeft,
                        AutoEllipsis = true
                    };

                    var tb = new TextBox
                    {
                        Multiline = true,
                        ReadOnly = true,
                        ScrollBars = ScrollBars.Both,
                        WordWrap = false, // 折り返しなし
                        Font = new Font(FontFamily.GenericMonospace, 9f),
                        Dock = DockStyle.Fill,
                        BackColor = SystemColors.Window
                    };

                    // Label を上に、TextBox を下に配置
                    container.Controls.Add(tb);
                    container.Controls.Add(lbl);

                    traceroutePanel.Controls.Add(container, i, 0);
                    tracerouteTextBoxes[address] = tb;
                }
            }
            else
            {
                // reuse existing: just add header line to each targeted box
            }

            // 共通ヘッダ（全列に表示）を出力 — 既存の結果が残っている場合は追記される
            AppendTracerouteOutputToAll("================================================================\r\n", false);
            AppendTracerouteOutputToAll($"== Traceroute: {DateTime.Now:yyyy/MM/dd HH:mm:ss} ==\r\n", false);
            AppendTracerouteOutputToAll("----------------------------------------------------------------\r\n", false);

            // Traceroute専用 CTS（Ping の cts と連動させるが独立してキャンセル可能）
            tracerouteCts?.Dispose();
            tracerouteCts = CancellationTokenSource.CreateLinkedTokenSource(cts?.Token ?? CancellationToken.None);

            var token = tracerouteCts.Token;

            bool wasCanceled = false;

            // 完了状態をアドレス単位で保持する辞書 (並列用)
            var completion = new ConcurrentDictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            var stoppedByUser = new ConcurrentDictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            foreach (var a in targets)
            {
                completion[a] = false;
                stoppedByUser[a] = false;
            }

            // フィールドにセットして Stop ボタンからアクセス可能にする
            _tracerouteCompletion = completion;
            _tracerouteStoppedByUser = stoppedByUser;

            try
            {
                await RunTracerouteForAddressesAsync(targets, token, completion);
            }
            catch (OperationCanceledException)
            {
                wasCanceled = true;
            }
            catch (Exception ex)
            {
                AppendTracerouteOutputToAll($"--- エラー: {ex.Message} ---\r\n\r\n", false);
            }
            finally
            {
                // token の状態もチェックして確実にユーザー停止を検出
                if (tracerouteCts != null && tracerouteCts.IsCancellationRequested) wasCanceled = true;

                // 変更: 停止（キャンセル）された場合のみ、未完了ターゲットへ停止メッセージを挿入する。
                //       既に完了しているターゲットには停止メッセージを挿入しない。
                foreach (var address in targets)
                {
                    bool done = _tracerouteCompletion != null && _tracerouteCompletion.TryGetValue(address, out var d) && d;
                    bool stoppedAlready = _tracerouteStoppedByUser != null && _tracerouteStoppedByUser.TryGetValue(address, out var s) && s;

                    if (done)
                    {
                        // 完了済み: 完了メッセージを付加（ユーザー停止メッセージが既にある場合は付けない）
                        if (!stoppedAlready)
                        {
                            AppendTracerouteOutput(address, "=== Traceroute 完了 ===\r\n\r\n", false);
                        }
                    }
                    else
                    {
                        // 未完了:
                        // - ユーザーが停止を要求した（wasCanceled == true）場合は停止メッセージを挿入
                        // - 既に BtnStopTraceroute により挿入済み（stoppedAlready）は二重挿入を避ける
                        if (wasCanceled && !stoppedAlready)
                        {
                            AppendTracerouteOutput(address, "=== Tracerouteは途中で停止されました。 ===\r\n\r\n", false);
                            if (_tracerouteStoppedByUser != null) _tracerouteStoppedByUser[address] = true;
                        }
                        else if (!wasCanceled && !stoppedAlready)
                        {
                            // キャンセルでない（通常は例外等の異常終了）場合は状態不明メッセージ
                            AppendTracerouteOutput(address, "=== Traceroute 終了 ===\r\n\r\n", false);
                            if (_tracerouteStoppedByUser != null) _tracerouteStoppedByUser[address] = true;
                        }
                    }
                }

                AppendTracerouteOutputToAll("================================================================\r\n\r\n", false);

                // 実行終了後にフラグを戻してボタンを更新
                _isTracerouteRunning = false;
                UpdateTracerouteButtons();

                // 自動保存追加
                if (mnuAutoSaveTraceroute != null && mnuAutoSaveTraceroute.Checked)
                {
                    SaveTracerouteOutputsAutoAppend();
                }

                tracerouteCts?.Dispose();
                tracerouteCts = null;

                // フィールドをクリア
                _tracerouteCompletion = null;
                _tracerouteStoppedByUser = null;
            }
        }

        private async Task RunTracerouteForAddressesAsync(List<string> addresses, CancellationToken token, ConcurrentDictionary<string, bool> completion)
        {
            // 並列実行（SemaphoreSlim により制御）
            var tasks = addresses.Select(async address =>
            {
                await tracerouteSemaphore.WaitAsync(token).ConfigureAwait(false);
                try
                {
                    await RunTracerouteAsync(address, token, completion).ConfigureAwait(false);
                }
                finally
                {
                    tracerouteSemaphore.Release();
                }
            }).ToArray();

            await Task.WhenAll(tasks).ConfigureAwait(false);
        }

        private async Task RunTracerouteAsync(string address, CancellationToken token, ConcurrentDictionary<string, bool> completion)
        {
            // Windows の tracert を使用。-d (名前解決なし) を付けることで高速化。
            // タイムアウト等は tracert の -w で調整可能（ms）。ここでは cmbTimeout の値を利用。
            int tryTimeoutMs = 4000;
            int parsed;
            if (int.TryParse(cmbTimeout?.Text, out parsed))
            {
                // tracert の -w は各ホップのタイムアウト(ms)。ここで使用。
                tryTimeoutMs = Math.Max(100, parsed);
            }

            string arguments = $"-d -w {tryTimeoutMs} {address}";

            // 各アドレスごとに見出し行をセット（追記）
            AppendTracerouteOutput(address, $"--- tracert {address} (timeout={tryTimeoutMs}ms) ---\r\n", false);

            var psi = new ProcessStartInfo
            {
                FileName = "tracert",
                Arguments = arguments,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.GetEncoding(932) // Windows の tracert 出力エンコーディング (日本環境想定)
            };

            using (var proc = new Process { StartInfo = psi })
            {
                try
                {
                    proc.Start();

                    // StandardOutput を逐次読み取り、1行到着ごとに UI へ反映する
                    using (var reader = proc.StandardOutput)
                    {
                        // キャンセル時にプロセスを強制終了する登録
                        using (token.Register(() =>
                        {
                            try
                            {
                                if (!proc.HasExited)
                                {
                                    proc.Kill();
                                }
                            }
                            catch { /* ignore */ }
                        }))
                        {
                            while (!reader.EndOfStream && !token.IsCancellationRequested)
                            {
                                string line;
                                try
                                {
                                    line = await reader.ReadLineAsync().ConfigureAwait(false);
                                }
                                catch (Exception exRead)
                                {
                                    line = $"(出力取得エラー: {exRead.Message})";
                                }

                                if (line == null) break;

                                // 1行到着ごとにアドレス別 TextBox へ追加（タイムスタンプは行頭に付けない仕様）
                                AppendTracerouteOutput(address, line + Environment.NewLine, false);
                            }
                        }
                    }

                    // プロセス終了を待つ（既に終了していることが多い）
                    try
                    {
                        if (!proc.HasExited)
                        {
                            proc.WaitForExit(1000);
                        }
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    AppendTracerouteOutput(address, $"(tracert 実行エラー: {ex.Message})\r\n", false);
                }
                finally
                {
                    if (!proc.HasExited)
                    {
                        try { proc.Kill(); } catch { }
                    }

                    // このアドレスは処理が終了した（成功・失敗に関わらず）としてマーク
                    if (completion != null) completion[address] = true;
                }
            }
        }

        /// <summary>
        /// 指定したアドレスの出力領域へ1行ずつ追加する。アドレスごとに分割表示するためのメソッド。
        /// ※ 仕様変更: 行の最初にタイムスタンプを記載しない（keepTimestamp を無視）。
        /// </summary>
        private void AppendTracerouteOutput(string address, string text, bool keepTimestamp)
        {
            if (string.IsNullOrEmpty(address))
            {
                // フォールバック: 全列へ出力
                AppendTracerouteOutputToAll(text, keepTimestamp);
                return;
            }

            if (tracerouteTextBoxes == null || !tracerouteTextBoxes.ContainsKey(address))
            {
                // 未作成ならフォールバックで全体に出力
                AppendTracerouteOutputToAll(text, keepTimestamp);
                return;
            }

            var tb = tracerouteTextBoxes[address];
            if (tb == null) return;

            if (tb.InvokeRequired)
            {
                tb.Invoke(new Action<string, string, bool>(AppendTracerouteOutput), address, text, keepTimestamp);
                return;
            }

            // 仕様: 行頭タイムスタンプを付けない → ここではタイムスタンプ挿入を行わない
            tb.AppendText(text);
            tb.SelectionStart = tb.Text.Length;
            tb.ScrollToCaret();

            // 保存/クリア/実行ボタン状態更新
            UpdateTracerouteButtons();
        }

        /// <summary>
        /// 全てのトレーサウト領域に同報で出力する（ヘッダーなど）。
        /// ※ 仕様変更: 行の最初にタイムスタンプを記載しない（keepTimestamp を無視）。
        /// </summary>
        private void AppendTracerouteOutputToAll(string text, bool keepTimestamp)
        {
            if (tracerouteTextBoxes == null || tracerouteTextBoxes.Count == 0) return;

            foreach (var kv in tracerouteTextBoxes)
            {
                var tb = kv.Value;
                if (tb == null) continue;

                if (tb.InvokeRequired)
                {
                    tb.Invoke(new Action<string, bool>((t, k) =>
                    {
                        // タイムスタンプを付けない仕様
                        tb.AppendText(t);
                        tb.SelectionStart = tb.Text.Length;
                        tb.ScrollToCaret();
                    }), text, keepTimestamp);
                }
                else
                {
                    // タイムスタンプを付けない仕様
                    tb.AppendText(text);
                    tb.SelectionStart = tb.Text.Length;
                    tb.ScrollToCaret();
                }
            }

            UpdateTracerouteButtons();
        }

        private void BtnSaveTraceroute_Click(object sender, EventArgs e)
        {
            if (tracerouteTextBoxes == null || tracerouteTextBoxes.Count == 0)
            {
                MessageBox.Show("保存するTraceroute結果がありません。", "保存不可", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 実行ファイルと同じフォルダへ保存する仕様（要件1）
            string folder = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                // 各対象アドレスごとにファイルを作成（追記モード）。ファイル名に使えない文字は置換。
                foreach (var kv in tracerouteTextBoxes)
                {
                    string address = kv.Key;
                    string content = kv.Value.Text;
                    if (string.IsNullOrEmpty(content)) continue;

                    // Host名を monitorList から取得（同一アドレスが複数ある場合は最初のものを使用）
                    string hostName = monitorList?
                        .FirstOrDefault(i => string.Equals(i.対象アドレス, address, StringComparison.OrdinalIgnoreCase))
                        ?.Host名 ?? "";

                    // ファイル名: Traceroute_result_yyyymmdd_対象アドレス_ホスト名.log (拡張子を .log に統一)
                    string safeAddress = SanitizeFileName(address);
                    string safeHost = SanitizeFileName(hostName);
                    string fileName = Path.Combine(folder, $"Traceroute_result_{DateTime.Now:yyyyMMdd}_{safeAddress}_{safeHost}.log");

                    // 追記で保存（既存があれば追記）
                    using (var sw = new StreamWriter(fileName, true, System.Text.Encoding.GetEncoding(932)))
                    {
                        sw.WriteLine("----------------------------------------------------------------");
                        sw.WriteLine($"Saved: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                        sw.WriteLine(content);
                        sw.WriteLine();
                    }

                    // 保存後は画面上の該当TextBoxをクリア（要件3）
                    if (kv.Value.InvokeRequired)
                    {
                        kv.Value.Invoke(new Action(() => kv.Value.Clear()));
                    }
                    else
                    {
                        kv.Value.Clear();
                    }
                }

                MessageBox.Show("Traceroute 出力を実行ファイルと同じフォルダへ保存しました（対象ごとに追記保存）。", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ファイルの保存中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // 保存後はボタン状態を更新（表示をクリアしているため Save は無効化される）
            UpdateTracerouteButtons();
        }

        // helper: ファイル名に使えない文字を安全に置換
        private string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "unknown";
            var invalid = Path.GetInvalidFileNameChars();
            var s = new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
            // さらに長すぎる場合は短縮
            if (s.Length > 120) s = s.Substring(0, 120);
            return s;
        }

        // 追加: チェックボックスの即時反映用イベント（編集状態をコミットして CellValueChanged を発火させる）
        private void DgvMonitor_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvMonitor.IsCurrentCellDirty)
            {
                var cell = dgvMonitor.CurrentCell;
                if (cell is DataGridViewCheckBoxCell)
                {
                    // チェックボックスは編集直後にコミットすることで CellValueChanged を発火させる
                    dgvMonitor.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        // 追加: チェックボックス変更を検知したら Traceroute ボタン状態を更新
        private void DgvMonitor_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var col = dgvMonitor.Columns[e.ColumnIndex];
            if (col.DataPropertyName == "Trace")
            {
                UpdateTracerouteButtons();
            }
        }

        // 追加: Trace結果クリアの実装
        private void BtnClearTraceroute_Click(object sender, EventArgs e)
        {
            if (tracerouteTextBoxes == null) return;

            // UI スレッド上で安全にクリア（各 TextBox）
            foreach (var tb in tracerouteTextBoxes.Values)
            {
                if (tb == null) continue;
                if (tb.InvokeRequired)
                {
                    tb.Invoke(new Action(() => tb.Clear()));
                }
                else
                {
                    tb.Clear();
                }
            }

            // ボタン状態更新
            UpdateTracerouteButtons();
        }

        // 追加: Traceroute 停止ボタンの処理（実行中にキャンセル）
        private void BtnStopTraceroute_Click(object sender, EventArgs e)
        {
            // 仕様: Traceroute停止ボタンが押された時点で、まだ完了していないターゲットには
            //       即座に「途中で停止されました」メッセージを挿入する（完了済みはスキップ）。
            try
            {
                if (_tracerouteCompletion != null)
                {
                    foreach (var kv in _tracerouteCompletion)
                    {
                        var address = kv.Key;
                        var done = kv.Value;
                        bool stoppedAlready = _tracerouteStoppedByUser != null && _tracerouteStoppedByUser.TryGetValue(address, out var s) && s;

                        if (!done && !stoppedAlready)
                        {
                            // 未完了のターゲットに停止メッセージを挿入（完了済みは挿入しない）
                            AppendTracerouteOutput(address, "=== Tracerouteは途中で停止されました。 ===\r\n\r\n", false);
                            if (_tracerouteStoppedByUser != null) _tracerouteStoppedByUser[address] = true;
                        }
                    }
                }
            }
            catch { /* UI更新中などの例外は無視 */ }

            if (tracerouteCts != null)
            {
                try
                {
                    tracerouteCts.Cancel();
                }
                catch { }
            }

            // UIはすぐに更新するが、実際の終了判定は BtnTraceroute の finally で行う
            _isTracerouteRunning = false;
            UpdateTracerouteButtons();
        }
        // ---------------------------------------------------------------
        // （既存の他メソッド、フィールドはそのまま）
    }
}