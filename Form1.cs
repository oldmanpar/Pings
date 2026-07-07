using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;
using Pings.Models;
using Pings.Repositories;
using Pings.Services;
using Pings.Utils;

namespace Pings
{
    public partial class Form1 : Form
    {
        private enum UiState { Initial, Running, Stopped }

        // 現在のUI状態を導出するプロパティ
        private UiState CurrentUiState =>
            cts != null ? UiState.Running :
            string.IsNullOrEmpty(txtStartTime.Text) ? UiState.Initial : UiState.Stopped;

        // -------------------------
        // 依存オブジェクト (Service / Repository)
        // -------------------------
        private readonly PingFileRepository _repository;
        private readonly PingService _pingService;
        private readonly TracerouteService _tracerouteService;

        // -------------------------
        // データバインディング
        // -------------------------
        private SortableBindingList<DisruptionLogItem> disruptionLogList;
        private BindingList<PingMonitorItem> monitorList;

        // -------------------------
        // UIコントロール (デザイナーコードの代わり)
        // -------------------------
        private DataGridView dgvMonitor;
        private DataGridView dgvLog;
        private TextBox txtStartTime, txtEndTime;
        private ComboBox cmbInterval, cmbTimeout;
        private Button btnPingStart, btnStop, btnClear, btnSave, btnExit;
        private TabControl tabControl;
        private TabPage statsPage;
        private TabPage traceroutePage;

        // ステータスボード関連
        private TabPage statusBoardPage;
        private Panel statusBoardHost;
        private FlowLayoutPanel statusBoardFlow;
        private Label lblBoardTotal, lblBoardOk, lblBoardDown, lblBoardRecovery;
        private TableLayoutPanel traceroutePanel;
        private Button btnTraceroute, btnSaveTraceroute, btnClearTraceroute, btnStopTraceroute;
        private ToolStripMenuItem mnuAutoSaveAllPing, mnuAutoSaveTraceroute;
        private ToolStripMenuItem mnuShowJitter1, mnuShowJitter2, mnuShowStdDev;
        private ToolStripMenuItem mnuTracerouteNoResolve;

        // インターフェイス選択用 (非表示機能)
        private ComboBox cmbInterface;
        private Button btnRefreshInterfaces;

        // 操作方法ウィンドウ（モードレス・単一インスタンス）
        private HelpForm helpForm;

        // -------------------------
        // 状態管理フィールド
        // -------------------------
        private CancellationTokenSource cts;
        private CancellationTokenSource tracerouteCts;
        private System.Windows.Forms.Timer uiUpdateTimer;
        private int _nextIndex = 1;
        private bool _allowEditAfterStop = true;
        private bool _allowIntervalTimeoutEdit = true;
        private bool _autoSaveFailed = false;
        private volatile bool _isTracerouteRunning = false;
        private bool _recalculating = false;

        // Traceroute関連
        private Dictionary<string, TextBox> tracerouteTextBoxes;
        private SemaphoreSlim tracerouteSemaphore = new SemaphoreSlim(4);
        private const int TracerouteColumnWidth = 480;
        private ConcurrentDictionary<string, bool> _tracerouteCompletion;
        private ConcurrentDictionary<string, bool> _tracerouteStoppedByUser;
        // 出力有無のフラグ（全テキストボックス走査を避けるため）
        private volatile bool _tracerouteHasOutput = false;

        // ソート状態
        private DataGridViewColumn currentSortColumn = null;
        private ListSortDirection currentSortDirection = ListSortDirection.Ascending;

        private const int UiUpdateInterval = 1000;

        // DataGridView カラム名定数
        private const string ColJitter1       = "colJitter1";
        private const string ColJitter2       = "colJitter2";
        private const string ColStdDev        = "colStdDev";
        private const string ColLogJitter1Pre  = "colLogJitter1_Pre";
        private const string ColLogJitter2Pre  = "colLogJitter2_Pre";
        private const string ColLogStdDevPre   = "colLogStdDev_Pre";
        private const string ColLogJitter1Post = "colLogJitter1_Post";
        private const string ColLogJitter2Post = "colLogJitter2_Post";
        private const string ColLogStdDevPost  = "colLogStdDev_Post";

        public Form1()
        {
            InitializeComponent();
            SetupMenuStrip();
            InitializeCustomComponents();
            SetupDataGridViewColumns();

            UpdateColumnVisibility();

            _repository = new PingFileRepository();
            _tracerouteService = new TracerouteService();
            _pingService = new PingService((logItem) =>
            {
                if (this.InvokeRequired)
                    this.Invoke(new Action(() => AddDisruptionLogItem(logItem)));
                else
                    AddDisruptionLogItem(logItem);
            });

            monitorList = new BindingList<PingMonitorItem>();
            monitorList.ListChanged += MonitorList_ListChanged;
            dgvMonitor.DataSource = monitorList;

            dgvMonitor.UserDeletedRow += DgvMonitor_UserDeletedRow;
            dgvMonitor.RowValidated += DgvMonitor_RowValidated;
            dgvMonitor.DefaultValuesNeeded += DgvMonitor_DefaultValuesNeeded;
            this.FormClosing += Form1_FormClosing;

            monitorList.Add(new PingMonitorItem(_nextIndex++, "127.0.0.1", "loopback"));
            monitorList.Add(new PingMonitorItem(_nextIndex++, "8.8.8.8", "Google DNS1"));
            monitorList.Add(new PingMonitorItem(_nextIndex++, "8.8.4.4", "Google DNS2"));

            UpdateUiState(UiState.Initial);
        }
    }
}
