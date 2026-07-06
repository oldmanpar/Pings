using System;

namespace Pings.Models
{
    /// <summary>
    /// 監視対象ノードの情報・統計・状態を保持するクラス。
    /// 統計の更新は RecordSuccess / RecordFailure を通じて行う。
    /// </summary>
    public class PingMonitorItem
    {
        // ============================================================
        // I. UI バインド用プロパティ（DataGridView が参照する）
        // ============================================================
        public int    順番           { get; set; }
        public string ステータス     { get; set; }
        public string 対象アドレス   { get; set; }
        public string Host名         { get; set; }

        public long   送信回数       { get; private set; }
        public long   失敗回数       { get; private set; }
        public int    連続失敗回数   { get; private set; }
        public string 連続失敗時間s  { get; private set; }
        public string 最大失敗時間s  { get; private set; }

        public long   時間ms         { get; private set; }
        public double 平均値ms       { get; private set; }
        public long   最小値ms       { get; private set; }
        public long   最大値ms       { get; private set; }
        public long   Jitter1_MaxMin { get; private set; }
        public double Jitter2_PktPair { get; private set; }
        public double StdDev         { get; private set; }

        // ============================================================
        // II. 設定値（Form1 から書き込む）
        // ============================================================
        public int  送信間隔ms { get; set; } = 1000;
        public int  タイムアウトms { get; set; } = 2000;
        public bool Trace      { get; set; } = false;

        // ============================================================
        // III. 内部状態（private）
        // ============================================================
        private bool      _isCurrentlyDown             = false;
        private DateTime? _continuousDownStartTime     = null;
        private TimeSpan  _maxDisruptionDuration       = TimeSpan.Zero;

        private long   _currentSessionUpTimeMs         = 0;
        private int    _currentSessionSuccessCount      = 0;
        private long   _currentSessionMin              = 0;
        private long   _currentSessionMax              = 0;
        private double _currentSessionSumSquares       = 0;
        private long   _previousRtt                    = -1;
        private long   _jitterDiffSum                  = 0;
        private int    _jitterDiffCount                = 0;

        // Down 開始時点のスナップショット
        private double _snapAvg    = 0.0;
        private long   _snapMin    = 0;
        private long   _snapMax    = 0;
        private long   _snapJitter1 = 0;
        private double _snapJitter2 = 0.0;
        private double _snapStdDev  = 0.0;

        private int  _currentDisruptionFailureCount = 0;
        private long _totalDownCount               = 0;
        private long _totalRecoveryCount           = 0;

        // 復旧後の統計をリアルタイム更新するためのログ参照
        private DisruptionLogItem _activeLogItem = null;

        // ============================================================
        // コンストラクタ
        // ============================================================
        public PingMonitorItem() : this(0, "", "") { }

        public PingMonitorItem(int index, string address, string hostName)
        {
            順番 = index;
            対象アドレス = address;
            Host名 = hostName;
            ResetData();
        }

        // ============================================================
        // IV. 統計更新メソッド（PingService から呼ぶ）
        // ============================================================

        /// <summary>
        /// Ping 成功時に呼ぶ。Down→復旧の遷移が発生した場合はログアイテムを返す。
        /// </summary>
        public DisruptionLogItem RecordSuccess(long rtt)
        {
            送信回数++;
            DisruptionLogItem newLogItem = null;

            // ---- 復旧遷移 ----
            if (_isCurrentlyDown && _continuousDownStartTime.HasValue)
            {
                _totalRecoveryCount++;
                ステータス = $"復旧_{_totalRecoveryCount}";

                TimeSpan duration = DateTime.Now - _continuousDownStartTime.Value;

                newLogItem = new DisruptionLogItem
                {
                    対象アドレス = this.対象アドレス,
                    Host名       = this.Host名,
                    Down開始日時 = _continuousDownStartTime.Value,
                    復旧日時     = DateTime.Now,
                    失敗時間mmss = duration.ToString(@"mm\:ss"),
                    失敗回数     = _currentDisruptionFailureCount,
                    Down前平均ms = _snapAvg,
                    Down前最小ms = _snapMin,
                    Down前最大ms = _snapMax,
                    Down前Jitter1 = _snapJitter1,
                    Down前Jitter2 = _snapJitter2,
                    Down前StdDev  = _snapStdDev,
                    復旧後平均ms = rtt,
                    復旧後最小ms = rtt,
                    復旧後最大ms = rtt,
                    復旧後Jitter1 = 0,
                    復旧後Jitter2 = 0,
                    復旧後StdDev  = 0
                };
                _activeLogItem = newLogItem;

                // セッション統計リセット
                _currentSessionUpTimeMs    = 0;
                _currentSessionSuccessCount = 0;
                _currentSessionMin         = 0;
                _currentSessionMax         = 0;
                _currentSessionSumSquares  = 0;
                _previousRtt               = -1;
                _jitterDiffSum             = 0;
                _jitterDiffCount           = 0;
                Jitter1_MaxMin             = 0;
                Jitter2_PktPair            = 0.0;
                StdDev                     = 0.0;
                _currentDisruptionFailureCount = 0;
                _isCurrentlyDown           = false;
            }

            // ---- 共通: 成功時 ----
            時間ms        = rtt;
            連続失敗回数  = 0;
            連続失敗時間s         = "";
            _continuousDownStartTime = null;

            _currentSessionSuccessCount++;
            _currentSessionUpTimeMs   += rtt;
            _currentSessionSumSquares += rtt * rtt;

            if (_currentSessionSuccessCount == 1)
            {
                _currentSessionMin = rtt;
                _currentSessionMax = rtt;
            }
            else
            {
                if (rtt < _currentSessionMin) _currentSessionMin = rtt;
                if (rtt > _currentSessionMax) _currentSessionMax = rtt;

                if (_previousRtt >= 0)
                {
                    long diff = Math.Abs(rtt - _previousRtt);
                    _jitterDiffSum += diff;
                    _jitterDiffCount++;
                    Jitter2_PktPair = (double)_jitterDiffSum / _jitterDiffCount;
                }
            }

            _previousRtt  = rtt;
            平均値ms      = (double)_currentSessionUpTimeMs / _currentSessionSuccessCount;
            Jitter1_MaxMin = _currentSessionMax - _currentSessionMin;
            最小値ms      = _currentSessionMin;
            最大値ms      = _currentSessionMax;

            double mean     = 平均値ms;
            double variance = (_currentSessionSumSquares / _currentSessionSuccessCount) - (mean * mean);
            StdDev          = Math.Sqrt(Math.Max(0, variance));

            // 復旧後の統計をリアルタイム更新
            if (_activeLogItem != null)
            {
                _activeLogItem.復旧後平均ms  = 平均値ms;
                _activeLogItem.復旧後最小ms  = 最小値ms;
                _activeLogItem.復旧後最大ms  = 最大値ms;
                _activeLogItem.復旧後Jitter1 = Jitter1_MaxMin;
                _activeLogItem.復旧後Jitter2 = Jitter2_PktPair;
                _activeLogItem.復旧後StdDev  = StdDev;
            }

            if (string.IsNullOrEmpty(ステータス)) ステータス = "OK";

            return newLogItem;
        }

        /// <summary>
        /// Ping 失敗時に呼ぶ。
        /// </summary>
        public void RecordFailure()
        {
            送信回数++;

            if (!_isCurrentlyDown)
            {
                _totalDownCount++;
                ステータス = $"Down_{_totalDownCount}";

                _activeLogItem = null;
                _snapAvg    = 平均値ms;
                _snapMin    = 最小値ms;
                _snapMax    = 最大値ms;
                _snapJitter1 = Jitter1_MaxMin;
                _snapJitter2 = Jitter2_PktPair;
                _snapStdDev  = StdDev;

                _continuousDownStartTime       = DateTime.Now;
                _isCurrentlyDown               = true;
                _currentDisruptionFailureCount = 0;
            }

            失敗回数++;
            _currentDisruptionFailureCount++;
            時間ms = 0;
            連続失敗回数++;
            _previousRtt = -1;

            TimeSpan currentDuration = DateTime.Now - _continuousDownStartTime.Value;
            連続失敗時間s = currentDuration.ToString(@"mm\:ss");

            if (currentDuration > _maxDisruptionDuration)
            {
                _maxDisruptionDuration = currentDuration;
                最大失敗時間s = _maxDisruptionDuration.ToString(@"mm\:ss");
            }
        }

        // ============================================================
        // V. データリセット
        // ============================================================

        /// <summary>統計データをリセットする（監視開始時に呼ぶ）</summary>
        public void ResetData()
        {
            送信回数     = 0;
            失敗回数     = 0;
            連続失敗回数 = 0;
            時間ms       = 0;
            平均値ms     = 0.0;
            最小値ms     = 0;
            最大値ms     = 0;
            Jitter1_MaxMin = 0;
            Jitter2_PktPair = 0.0;
            StdDev       = 0.0;

            _totalDownCount              = 0;
            _totalRecoveryCount          = 0;
            _currentSessionUpTimeMs      = 0;
            _currentSessionSuccessCount  = 0;
            _currentSessionMin           = 0;
            _currentSessionMax           = 0;
            _currentSessionSumSquares    = 0;
            _previousRtt                 = -1;
            _jitterDiffSum               = 0;
            _jitterDiffCount             = 0;
            _continuousDownStartTime     = null;
            _maxDisruptionDuration       = TimeSpan.Zero;
            _isCurrentlyDown             = false;
            _activeLogItem               = null;
            _currentDisruptionFailureCount = 0;

            ステータス    = "";
            連続失敗時間s = "";
            最大失敗時間s = "";
        }
    }
}
