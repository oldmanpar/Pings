using System;

namespace Pings.Models
{
    /// <summary>
    /// 監視対象ノードの情報と統計情報を保持するデータクラス
    /// </summary>
    public class PingMonitorItem
    {
        // I. UI表示用プロパティ
        public int 順番 { get; set; }
        public string ステータス { get; set; }
        public string 対象アドレス { get; set; }
        public string Host名 { get; set; }

        // 統計情報
        public long 送信回数 { get; set; }
        public long 失敗回数 { get; set; }
        public int 連続失敗回数 { get; set; }
        public string 連続失敗時間s { get; set; }
        public string 最大失敗時間s { get; set; }

        // メイングリッド表示用
        public long 時間ms { get; set; }
        public double 平均値ms { get; set; }
        public long 最小値ms { get; set; }
        public long 最大値ms { get; set; }

        // ★追加: 新規統計項目 (UIバインド用)
        public long Jitter1_MaxMin { get; set; } // 最大値 - 最小値
        public double Jitter2_PktPair { get; set; } // パケットペア平均
        public double StdDev { get; set; }       // 標準偏差

        // 設定値
        public int 送信間隔ms { get; set; } = 1000;
        public int タイムアウトms { get; set; } = 2000;
        public bool Trace { get; set; } = false;

        // II. 内部状態管理用フィールド (Serviceから操作)
        public bool IsUp { get; set; } = false;
        public bool HasBeenDown { get; set; } = false;

        // 統計計算用の状態変数
        public DateTime? ContinuousDownStartTime { get; set; } = null;
        public TimeSpan MaxDisruptionDuration { get; set; } = TimeSpan.Zero;

        // セッション統計計算用
        public long CurrentSessionUpTimeMs { get; set; } = 0;
        public int CurrentSessionSuccessCount { get; set; } = 0;
        public long CurrentSessionMin { get; set; } = 0;
        public long CurrentSessionMax { get; set; } = 0;

        // ★追加: 統計計算用内部変数
        public double CurrentSessionSumSquares { get; set; } = 0; // 二乗和(標準偏差用)
        public long PreviousRtt { get; set; } = -1;               // 前回のRTT(Jitter2用)
        public long JitterDiffSum { get; set; } = 0;              // 差分和(Jitter2用)
        public int JitterDiffCount { get; set; } = 0;             // 差分回数(Jitter2用)

        // Down時の統計スナップショット
        public bool IsCurrentlyDown { get; set; } = false;
        public double SnapAvg { get; set; } = 0.0;
        public long SnapMin { get; set; } = 0;
        public long SnapMax { get; set; } = 0;

        // ★追加: Down時スナップショット用
        public long SnapJitter1 { get; set; } = 0;
        public double SnapJitter2 { get; set; } = 0.0;
        public double SnapStdDev { get; set; } = 0.0;

        // 障害中の失敗回数カウンター
        public int CurrentDisruptionFailureCount { get; set; } = 0;

        // 現在アクティブなログエントリへの参照
        public DisruptionLogItem ActiveLogItem { get; set; } = null;

        public PingMonitorItem() : this(0, "", "") { }

        public PingMonitorItem(int index, string address, string hostName)
        {
            順番 = index;
            対象アドレス = address;
            Host名 = hostName;
            ResetData();
        }

        /// <summary>
        /// 統計データをリセットします（ロジックではなくデータ初期化のみ）
        /// </summary>
        public void ResetData()
        {
            送信回数 = 0;
            失敗回数 = 0;
            連続失敗回数 = 0;
            時間ms = 0;

            平均値ms = 0.0;
            最小値ms = 0;
            最大値ms = 0;

            // ★追加: リセット
            Jitter1_MaxMin = 0;
            Jitter2_PktPair = 0.0;
            StdDev = 0.0;

            CurrentSessionUpTimeMs = 0;
            CurrentSessionSuccessCount = 0;
            CurrentSessionMin = 0;
            CurrentSessionMax = 0;

            // ★追加: 内部変数リセット
            CurrentSessionSumSquares = 0;
            PreviousRtt = -1;
            JitterDiffSum = 0;
            JitterDiffCount = 0;

            ContinuousDownStartTime = null;
            MaxDisruptionDuration = TimeSpan.Zero;
            IsUp = false;
            HasBeenDown = false;
            IsCurrentlyDown = false;
            ActiveLogItem = null;

            ステータス = "";
            連続失敗時間s = "";
            最大失敗時間s = "";
            CurrentDisruptionFailureCount = 0;
        }
    }
}