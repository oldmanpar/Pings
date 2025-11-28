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
using System.Reflection;

// ---------------------------------------------------------
// 1. データ保持クラス (障害イベントログ)
// ---------------------------------------------------------
public class DisruptionLogItem
{
    public string 対象アドレス { get; set; }
    public string Host名 { get; set; }
    public DateTime Down開始日時 { get; set; }
    public DateTime 復旧日時 { get; set; }
    public string 失敗時間mmss { get; set; }
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
// 2. ソート可能な BindingList の実装
// ---------------------------------------------------------
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
        Type propType = prop.PropertyType;

        Comparison<T> comparer = (T x, T y) =>
        {
            object xValue = prop.GetValue(x);
            object yValue = prop.GetValue(y);

            if (xValue == null && yValue == null) return 0;
            if (xValue == null) return (direction == ListSortDirection.Ascending) ? -1 : 1;
            if (yValue == null) return (direction == ListSortDirection.Ascending) ? 1 : -1;

            if (xValue is IComparable comparableX)
            {
                return comparableX.CompareTo(yValue) * (direction == ListSortDirection.Ascending ? 1 : -1);
            }
            return 0;
        };

        originalList.Sort(comparer);
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
        public int 項番 { get; set; }
        public string ステータス { get; set; }
        public string 対象アドレス { get; set; }
        public string Host名 { get; set; }

        public long 送信回数 { get; private set; }
        public long 失敗回数 { get; private set; }
        public int 連続失敗回数 { get; private set; }
        public string 連続失敗時間s { get; private set; }
        public string 最大失敗時間s { get; private set; }

        public long 時間ms { get; set; }
        public double 平均値ms { get; private set; }
        public long 最小値ms { get; private set; }
        public long 最大値ms { get; private set; }

        public bool IsUp { get; private set; } = false;
        public bool HasBeenDown { get; private set; } = false;

        private DateTime? _continuousDownStartTime = null;
        private TimeSpan _maxDisruptionDuration = TimeSpan.Zero;
        private DisruptionLogItem _activeLogItem = null;

        private long _currentSessionSum = 0;
        private int _currentSessionCount = 0;
        private long _currentSessionMin = 0;
        private long _currentSessionMax = 0;

        private bool _isCurrentlyDown = false;
        private double _snapAvg = 0.0;
        private long _snapMin = 0;
        private long _snapMax = 0;

        public int 送信間隔ms { get; set; } = 500;
        public int タイムアウトms { get; set; } = 1000;

        public PingMonitorItem() : this(0, "", "") { }

        public PingMonitorItem(int index, string address, string hostName)
        {
            項番 = index;
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

        // ... (Form1.cs 内)

        public void UpdateStatistics(long rtt, bool success, Action<DisruptionLogItem> logAction)
        {
            送信回数++;

            if (success)
            {
                if (_isCurrentlyDown)
                {
                    // ★修正箇所: 復旧時: ログアイテムを新規作成する代わりに、既存の_activeLogItemを更新する
                    if (_activeLogItem != null)
                    {
                        TimeSpan duration = DateTime.Now - _continuousDownStartTime.Value;

                        // 既存のログアイテムを復旧情報で補完
                        _activeLogItem.復旧日時 = DateTime.Now; // 復旧日時を確定
                        _activeLogItem.失敗時間mmss = duration.ToString(@"hh\:mm\:ss");
                        _activeLogItem.連続失敗回数 = 連続失敗回数;
                        // 復旧時の初回RTTを復旧後平均/最小/最大に設定（すぐ下でセッション平均に更新される）
                        _activeLogItem.復旧後平均ms = rtt;
                        _activeLogItem.復旧後最小ms = rtt;
                        _activeLogItem.復旧後最大ms = rtt;
                    }

                    _currentSessionSum = 0;
                    _currentSessionCount = 0;
                    _currentSessionMin = 0;
                    _currentSessionMax = 0;
                    _isCurrentlyDown = false;
                }

                時間ms = rtt;
                IsUp = true;
                連続失敗回数 = 0;
                連続失敗時間s = "";
                _continuousDownStartTime = null;
                ステータス = HasBeenDown ? "復旧" : "OK";

                _currentSessionCount++;
                _currentSessionSum += rtt;

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

                平均値ms = (double)_currentSessionSum / _currentSessionCount;
                最小値ms = _currentSessionMin;
                最大値ms = _currentSessionMax;

                if (_activeLogItem != null)
                {
                    // 復旧後も成功が続いている場合、セッション統計で更新
                    _activeLogItem.復旧後平均ms = 平均値ms;
                    _activeLogItem.復旧後最小ms = 最小値ms;
                    _activeLogItem.復旧後最大ms = 最大値ms;
                }
            }
            else
            {
                if (!_isCurrentlyDown)
                {
                    // ★修正箇所: Down時: ログアイテムを新規作成し、リストに即座に追加する
                    _snapAvg = 平均値ms;
                    _snapMin = 最小値ms;
                    _snapMax = 最大値ms;
                    _continuousDownStartTime = DateTime.Now;
                    _isCurrentlyDown = true;

                    // 新しいログアイテムを作成し、リストに追加 (復旧日時をDateTime.MinValueで仮設定)
                    DisruptionLogItem newLogItem = new DisruptionLogItem
                    {
                        対象アドレス = 対象アドレス,
                        Host名 = Host名,
                        Down開始日時 = _continuousDownStartTime.Value,
                        復旧日時 = DateTime.MinValue, // 未復旧のマーク
                        失敗時間mmss = "",
                        連続失敗回数 = 1, // 最初の失敗なので1
                        Down前平均ms = _snapAvg,
                        Down前最小ms = _snapMin,
                        Down前最大ms = _snapMax,
                        // 復旧後 stats は 0/0.0
                    };
                    logAction(newLogItem);
                    _activeLogItem = newLogItem;
                }

                失敗回数++;
                時間ms = 0;
                ステータス = "Down";
                IsUp = false;
                HasBeenDown = true;
                連続失敗回数++;

                TimeSpan currentDuration = DateTime.Now - _continuousDownStartTime.Value;
                連続失敗時間s = currentDuration.ToString(@"hh\:mm\:ss");

                if (currentDuration > _maxDisruptionDuration)
                {
                    _maxDisruptionDuration = currentDuration;
                    最大失敗時間s = _maxDisruptionDuration.ToString(@"hh\:mm\:ss");
                }

                if (_activeLogItem != null)
                {
                    // 連続失敗回数をアクティブなログアイテムに反映させる
                    _activeLogItem.連続失敗回数 = 連続失敗回数;
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

        private SortableBindingList<DisruptionLogItem> disruptionLogList;
        private BindingList<PingMonitorItem> monitorList;

        private CancellationTokenSource cts;
        private System.Windows.Forms.Timer uiUpdateTimer;
        private const int UiUpdateInterval = 200;

        private DataGridViewColumn currentSortColumn = null;
        private ListSortDirection currentSortDirection = ListSortDirection.Ascending;

        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
            SetupDataGridViewColumns();
            SetupMenuStrip();

            monitorList = new BindingList<PingMonitorItem>();
            dgvMonitor.DataSource = monitorList;

            dgvMonitor.DefaultValuesNeeded += DgvMonitor_DefaultValuesNeeded;
            this.FormClosing += Form1_FormClosing;

            // 初期データの追加
            int idx = 1;
            monitorList.Add(new PingMonitorItem(idx++, "127.0.0.1", "loopback"));
            monitorList.Add(new PingMonitorItem(idx++, "8.8.8.8", "Google DNS1"));
            monitorList.Add(new PingMonitorItem(idx++, "8.8.4.4", "Google DNS2"));

            UpdateUiState("Initial");
        }

        // ----------------------------------------------------------------------
        // UI制御ロジック (仕様変更の要)
        // ----------------------------------------------------------------------
        private void UpdateUiState(string state)
        {
            // state: "Initial", "Running", "Stopped_Unsaved", "Stopped_Saved"

            if (this.MainMenuStrip != null)
            {
                ToolStripMenuItem fileMenu = this.MainMenuStrip.Items[0] as ToolStripMenuItem;
                if (fileMenu != null)
                {
                    // ファイル読込・保存は「初期状態」または「停止・保存済」のときのみ許可
                    // (停止・未保存時は編集を禁止するため、読込も整合性維持のためブロック推奨だが、
                    //  要望2「停止させた後...編集が出来てしまう」への対応として、
                    //  未保存状態での編集系アクションを制限する)
                    bool canEdit = (state == "Initial");
                    fileMenu.DropDownItems[0].Enabled = canEdit; // 対象アドレス保存
                    fileMenu.DropDownItems[1].Enabled = canEdit; // 対象アドレス読込
                    fileMenu.DropDownItems[3].Enabled = true;    // 終了
                }
            }

            cmbInterval.Enabled = (state == "Initial");
            cmbTimeout.Enabled = (state == "Initial");

            switch (state)
            {
                case "Initial": // 起動直後、またはクリア後
                    btnPingStart.Enabled = true;
                    btnStop.Enabled = false;
                    btnClear.Enabled = false;
                    btnSave.Enabled = false;

                    dgvMonitor.ReadOnly = false;
                    dgvMonitor.AllowUserToAddRows = true;
                    dgvMonitor.AllowUserToDeleteRows = true;
                    break;

                case "Running": // 監視中
                    btnPingStart.Enabled = false;
                    btnStop.Enabled = true;
                    btnClear.Enabled = false;
                    btnSave.Enabled = false;

                    dgvMonitor.ReadOnly = true;
                    dgvMonitor.AllowUserToAddRows = false;
                    dgvMonitor.AllowUserToDeleteRows = false;
                    break;

                case "Stopped_Unsaved": // 停止直後 (未保存、未クリア) -> 要望2,4対応
                    btnPingStart.Enabled = false; // 要望4: クリアか保存まではグレーアウト
                    btnStop.Enabled = false;
                    btnClear.Enabled = true;
                    btnSave.Enabled = true;

                    // 要望2,3: 停止後も編集・削除を禁止して結果を保護
                    dgvMonitor.ReadOnly = true;
                    dgvMonitor.AllowUserToAddRows = false;
                    dgvMonitor.AllowUserToDeleteRows = false;
                    break;

                case "Stopped_Saved": // 保存後 -> 要望4対応
                    btnPingStart.Enabled = true;  // 保存後は開始可能
                    btnStop.Enabled = false;
                    btnClear.Enabled = true;
                    btnSave.Enabled = true; // 何度でも保存可

                    // 要望3: 保存後も編集・削除は禁止 (開始ボタンを押すとクリアされるフローのため)
                    dgvMonitor.ReadOnly = true;
                    dgvMonitor.AllowUserToAddRows = false;
                    dgvMonitor.AllowUserToDeleteRows = false;
                    break;
            }
        }

        // ----------------------------------------------------------------------
        // イベントハンドラ
        // ----------------------------------------------------------------------

        private void BtnPingStart_Click(object sender, EventArgs e)
        {
            // 安全のため一度停止処理を呼ぶ
            StopMonitoring(false); // UI更新はしない

            dgvMonitor.EndEdit();

            // 要望5: Ping開始を押下したときは、結果を一度クリアする
            foreach (var item in monitorList)
            {
                item.ResetStatistics();
            }
            disruptionLogList.Clear();
            ResetLogSortIndicators();
            txtStartTime.Text = "";
            txtEndTime.Text = "";

            // 項番のリフレッシュ (念のため)
            RenumberMonitorList();

            StartMonitoring();
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            StopMonitoring(true);
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            // 完全にリセット
            StopMonitoring(false);

            foreach (var item in monitorList)
            {
                item.ResetStatistics();
            }
            disruptionLogList.Clear();
            txtStartTime.Text = "";
            txtEndTime.Text = "";

            monitorList.ResetBindings();
            ResetLogSortIndicators();

            UpdateUiState("Initial");
        }

        private void BtnSaveResult_Click(object sender, EventArgs e)
        {
            if (cts != null) return; // 監視中は押せないはずだが念のため

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.FileName = $"Pings_Result_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                sfd.Filter = "CSVファイル (*.csv)|*.csv|すべてのファイル (*.*)|*.*";
                sfd.Title = "監視結果を保存";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (StreamWriter sw = new StreamWriter(sfd.FileName, false, System.Text.Encoding.GetEncoding(932)))
                        {
                            sw.WriteLine($"開始日時　：　{txtStartTime.Text}");
                            sw.WriteLine($"終了日時　：　{txtEndTime.Text}");
                            sw.WriteLine($"送信間隔　：　{cmbInterval.Text} [ms]     タイムアウト　：　{cmbTimeout.Text} [ms]");
                            sw.WriteLine();

                            sw.WriteLine("--- 監視統計 ---");
                            string header = "ｽﾃｰﾀｽ,項番,対象アドレス,Host名,送信回数,失敗回数,連続失敗回数,連続失敗時間[hh:mm:ss],最大失敗時間[hh:mm:ss],時間[ms],平均値[ms],最小値[ms],最大値[ms]";
                            sw.WriteLine(header);

                            foreach (var item in monitorList.OrderBy(i => i.項番))
                            {
                                string line = string.Join(",",
                                    item.ステータス,
                                    item.項番,
                                    item.対象アドレス,
                                    item.Host名,
                                    item.送信回数,
                                    item.失敗回数,
                                    item.連続失敗回数,
                                    item.連続失敗時間s,
                                    item.最大失敗時間s,
                                    item.時間ms,
                                    $"{item.平均値ms:F1}",
                                    item.最小値ms,
                                    item.最大値ms
                                );
                                sw.WriteLine(line);
                            }

                            sw.WriteLine();
                            sw.WriteLine();

                            sw.WriteLine("--- 障害イベントログ ---");
                            if (disruptionLogList.Any())
                            {
                                string logHeader = "対象アドレス,Host名,Down開始日時,復旧日時,連続失敗回数,失敗時間[hh:mm:ss]," +
                                                   "Down前平均[ms],Down前最小[ms],Down前最大[ms]," +
                                                   "復旧後平均[ms],復旧後最小[ms],復旧後最大[ms]";
                                sw.WriteLine(logHeader);

                                foreach (var logItem in disruptionLogList.OrderBy(i => i.復旧日時))
                                {
                                    string logLine = string.Join(",",
                                        logItem.対象アドレス,
                                        logItem.Host名,
                                        logItem.Down開始日時.ToString("yyyy/MM/dd HH:mm:ss"),
                                        logItem.復旧日時.ToString("yyyy/MM/dd HH:mm:ss"),
                                        logItem.連続失敗回数,
                                        logItem.失敗時間mmss,
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

                        // 要望4: 保存後は開始ボタンを有効化する
                        UpdateUiState("Stopped_Saved");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ファイルの保存中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // ----------------------------------------------------------------------
        // 監視ロジック
        // ----------------------------------------------------------------------

        private void StartMonitoring()
        {
            cts = new CancellationTokenSource();
            txtStartTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            txtEndTime.Text = "";

            int interval = int.Parse(cmbInterval.Text);
            int timeout = int.Parse(cmbTimeout.Text);

            UpdateUiState("Running");
            uiUpdateTimer.Start();

            Action<DisruptionLogItem> logAction = AddDisruptionLogItem;

            foreach (var item in monitorList.Where(i => !string.IsNullOrEmpty(i.対象アドレス) && i.項番 > 0))
            {
                item.送信間隔ms = interval;
                item.タイムアウトms = timeout;
                // Note: リセットはBtnStart側で一括で行っているが、念のためここでも呼んでも良い
                Task.Run(() => RunPingLoopAsync(item, cts.Token, logAction));
            }
        }

        private void StopMonitoring(bool updateUi)
        {
            if (cts != null)
            {
                cts.Cancel();
                cts.Dispose();
                cts = null;
            }

            uiUpdateTimer.Stop();

            if (updateUi)
            {
                if (txtEndTime != null)
                {
                    txtEndTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    txtEndTime.BackColor = SystemColors.Window;
                }
                // 要望4: 停止直後は開始不可、編集不可
                UpdateUiState("Stopped_Unsaved");
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
                    catch (TaskCanceledException) { break; }
                    catch (Exception)
                    {
                        item.UpdateStatistics(0, false, logAction);
                        await Task.Delay(item.送信間隔ms, token);
                    }
                }
            }
            this.Invoke((MethodInvoker)delegate { dgvMonitor.Refresh(); });
        }

        private void AddDisruptionLogItem(DisruptionLogItem item)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<DisruptionLogItem>(AddDisruptionLogItem), item);
            }
            else
            {
                if (item.復旧日時 == DateTime.MinValue)
                {
                    // Down時: 新しいログを追加
                    disruptionLogList.Add(item);
                }
                // 復旧時: アイテムは既にあるため、UI更新はTimerに任せる。

                // ★修正箇所: ログが追加されたら、未復旧の項目に注目させるため末尾を表示
                if (dgvLog.RowCount > 0) dgvLog.FirstDisplayedScrollingRowIndex = dgvLog.RowCount - 1;
            }
        }
        private void DgvMonitor_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            // 要望1: 新規レコードの項番が1増えない問題の修正
            // 既存のリストの最大項番を取得して +1 するロジックに変更
            int nextId = 1;
            if (monitorList != null && monitorList.Count > 0)
            {
                // monitorList内の最大の項番を探す
                nextId = monitorList.Max(x => x.項番) + 1;
            }
            e.Row.Cells[1].Value = nextId; // Column Index 1 is 項番
        }

        private void dgvMonitor_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            RenumberMonitorList();
        }

        private void RenumberMonitorList()
        {
            if (cts != null) return;
            int newIndex = 1;
            // 項番順に並べ替えて1から振り直す
            foreach (var item in monitorList.OrderBy(i => i.項番))
            {
                item.項番 = newIndex++;
            }
            // _nextIndexはここでは不要(DefaultValuesNeededで動的に計算するため)
            monitorList.ResetBindings();
        }

        // ----------------------------------------------------------------------
        // その他 UI/Utility
        // ----------------------------------------------------------------------

        private void InitializeCustomComponents()
        {
            this.Text = "Pings B版";
            this.Size = new Size(1300, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            uiUpdateTimer = new System.Windows.Forms.Timer();
            uiUpdateTimer.Interval = UiUpdateInterval;
            uiUpdateTimer.Tick += UiUpdateTimer_Tick;

            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(10) };
            this.Controls.Add(topPanel);

            Label lblStart = new Label { Text = "開始日時", Location = new Point(10, 1), AutoSize = true };
            txtStartTime = new TextBox { Location = new Point(10, 19), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };
            Label lblEnd = new Label { Text = "終了日時", Location = new Point(170, 1), AutoSize = true };
            txtEndTime = new TextBox { Location = new Point(170, 19), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };

            topPanel.Controls.Add(lblStart);
            topPanel.Controls.Add(txtStartTime);
            topPanel.Controls.Add(lblEnd);
            topPanel.Controls.Add(txtEndTime);

            Label lblInterval = new Label { Text = "送信間隔 [ms]", Location = new Point(350, 1), AutoSize = true };
            cmbInterval = new ComboBox { Location = new Point(350, 19), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbInterval.Items.AddRange(new object[] { "100", "500", "1000", "2000" });
            cmbInterval.SelectedIndex = 1;

            Label lblTimeout = new Label { Text = "タイムアウト [ms]", Location = new Point(480, 1), AutoSize = true };
            cmbTimeout = new ComboBox { Location = new Point(480, 19), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTimeout.Items.AddRange(new object[] { "500", "1000", "2000", "5000" });
            cmbTimeout.SelectedIndex = 1;

            topPanel.Controls.Add(lblInterval);
            topPanel.Controls.Add(cmbInterval);
            topPanel.Controls.Add(lblTimeout);
            topPanel.Controls.Add(cmbTimeout);

            tabControl = new TabControl { Dock = DockStyle.Fill };
            this.Controls.Add(tabControl);

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

            TabPage logPage = new TabPage("障害イベントログ");
            dgvLog = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                ReadOnly = true,
                BackgroundColor = SystemColors.ControlLightLight,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                // ★修正箇所: 行削除を禁止する設定を追加
                AllowUserToDeleteRows = false
            };
            logPage.Controls.Add(dgvLog);
            tabControl.Controls.Add(logPage);

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

            btnPingStart.Click += BtnPingStart_Click;
            btnStop.Click += BtnStop_Click;
            btnClear.Click += BtnClear_Click;
            btnExit.Click += BtnExit_Click;
            btnSave.Click += BtnSaveResult_Click;

            tabControl.BringToFront();
            bottomPanel.SendToBack();
            topPanel.SendToBack();
            if (this.MainMenuStrip != null) this.MainMenuStrip.SendToBack();
        }

        private void SetupDataGridViewColumns()
        {
            dgvMonitor.AllowUserToAddRows = true;
            dgvMonitor.ReadOnly = false;
            dgvMonitor.Columns.Clear();

            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "ステータス", HeaderText = "ｽﾃｰﾀｽ", Width = 60, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "項番", HeaderText = "項番", Width = 50, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "対象アドレス", HeaderText = "対象アドレス", Width = 120, ReadOnly = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Host名", HeaderText = "Host名", Width = 120, ReadOnly = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "送信回数", HeaderText = "送信回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗回数", HeaderText = "失敗回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "連続失敗回数", HeaderText = "連続失敗回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "連続失敗時間s", HeaderText = "連続失敗時間[hh:mm:ss]", Width = 130, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最大失敗時間s", HeaderText = "最大失敗時間[hh:mm:ss]", Width = 130, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "時間ms", HeaderText = "時間[ms]", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "平均値ms", HeaderText = "平均値[ms]", Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" }, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最小値ms", HeaderText = "最小値[ms]", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最大値ms", HeaderText = "最大値[ms]", Width = 80, ReadOnly = true });

            dgvMonitor.AllowUserToDeleteRows = true;
            dgvMonitor.UserDeletedRow += dgvMonitor_UserDeletedRow;

            // ★修正箇所: CellFormatting イベントハンドラを追加し、未復旧の項目を空欄にする
            dgvLog.CellFormatting += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    if (dgvLog.Rows[e.RowIndex].DataBoundItem is DisruptionLogItem item)
                    {
                        // 復旧日時が MinValue の場合、未復旧として空欄で表示
                        if (e.ColumnIndex == dgvLog.Columns["復旧日時"].Index && item.復旧日時 == DateTime.MinValue)
                        {
                            e.Value = string.Empty;
                            e.FormattingApplied = true;
                        }
                        // 失敗時間mmss が空の場合、未復旧として空欄で表示
                        else if (e.ColumnIndex == dgvLog.Columns["失敗時間mmss"].Index && item.復旧日時 == DateTime.MinValue)
                        {
                            e.Value = string.Empty;
                            e.FormattingApplied = true;
                        }
                        // 復旧後平均ms が 0.0 の場合（未復旧の場合）、空欄で表示
                        else if ((e.ColumnIndex == dgvLog.Columns["復旧後平均ms"].Index || e.ColumnIndex == dgvLog.Columns["復旧後最小ms"].Index || e.ColumnIndex == dgvLog.Columns["復旧後最大ms"].Index) && item.復旧日時 == DateTime.MinValue)
                        {
                            e.Value = string.Empty;
                            e.FormattingApplied = true;
                        }
                    }
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
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "連続失敗回数", HeaderText = "連続失敗回数", Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗時間mmss", HeaderText = "失敗時間[hh:mm:ss]", Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前平均ms", HeaderText = "Down前平均[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最小ms", HeaderText = "Down前最小[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最大ms", HeaderText = "Down前最大[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後平均ms", HeaderText = "復旧後平均[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最小ms", HeaderText = "復旧後最小[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最大ms", HeaderText = "復旧後最大[ms]", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight } });
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
            menuStrip.Items.Add(new ToolStripMenuItem("オプション"));

            ToolStripMenuItem helpMenu = new ToolStripMenuItem("ヘルプ");
            ToolStripMenuItem versionItem = new ToolStripMenuItem("バージョン情報");
            versionItem.Click += VersionItem_Click;
            helpMenu.DropDownItems.Add(versionItem);

            menuStrip.Items.Add(helpMenu);
            this.Controls.Add(menuStrip);
            this.MainMenuStrip = menuStrip;
            menuStrip.SendToBack();
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中はアプリケーションを終了できません。\n[停止]ボタンを押して監視を止めてください。", "終了不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            StopMonitoring(false);
            this.Close();
        }

        private void VersionItem_Click(object sender, EventArgs e)
        {
            using (var aboutBox = new AboutBox())
            {
                aboutBox.ShowDialog(this);
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
                        using (StreamWriter sw = new StreamWriter(sfd.FileName, false, System.Text.Encoding.GetEncoding(932)))
                        {
                            sw.WriteLine("対象アドレス,Host名");
                            foreach (var item in monitorList.Where(i => !string.IsNullOrEmpty(i.対象アドレス) || !string.IsNullOrEmpty(i.Host名)))
                            {
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
            if (cts != null)
            {
                MessageBox.Show("Ping監視中はアドレスを読み込めません。\n[停止]ボタンを押して監視を止めてください。", "読込不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "アドレスファイル (*.csv;*.txt;*.log)|*.csv;*.txt;*.log|すべてのファイル (*.*)|*.*";
                ofd.Title = "監視対象アドレスを読み込み";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var newMonitorList = new List<PingMonitorItem>();
                        int index = 1;
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

                        foreach (string currentLine in allLines)
                        {
                            string trimmedLine = currentLine.Trim();
                            if (string.IsNullOrWhiteSpace(trimmedLine)) continue;
                            if (trimmedLine.StartsWith("[") || trimmedLine.StartsWith("#") || trimmedLine.StartsWith(";") || trimmedLine.StartsWith("'")) continue;

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
                            newMonitorList.Add(new PingMonitorItem(index++, address, hostName));
                        }

                        monitorList.Clear();
                        foreach (var item in newMonitorList)
                        {
                            monitorList.Add(item);
                        }

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

        private void UiUpdateTimer_Tick(object sender, EventArgs e)
        {
            if (dgvMonitor.IsHandleCreated)
            {
                dgvMonitor.Refresh();
                dgvLog.Refresh();
            }
        }

        private void dgvLog_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewColumn newColumn = dgvLog.Columns[e.ColumnIndex];
            ListSortDirection direction;

            if (currentSortColumn != null && newColumn.DataPropertyName == currentSortColumn.DataPropertyName)
            {
                direction = (currentSortDirection == ListSortDirection.Ascending) ? ListSortDirection.Descending : ListSortDirection.Ascending;
            }
            else
            {
                direction = ListSortDirection.Ascending;
            }

            PropertyDescriptor prop = TypeDescriptor.GetProperties(typeof(DisruptionLogItem)).Find(newColumn.DataPropertyName, true);

            if (prop != null)
            {
                if (currentSortColumn != null)
                {
                    currentSortColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                }

                disruptionLogList.Sort(prop, direction);
                newColumn.HeaderCell.SortGlyphDirection = (direction == ListSortDirection.Ascending) ? SortOrder.Ascending : SortOrder.Descending;
                currentSortColumn = newColumn;
                currentSortDirection = direction;
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中はアプリケーションを終了できません。\n[停止]ボタンを押して監視を止めてください。", "終了不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
            else
            {
                StopMonitoring(false);
            }
        }
    }
}