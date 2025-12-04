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

        // Down時の統計スナップショット
        public bool IsCurrentlyDown { get; set; } = false;
        public double SnapAvg { get; set; } = 0.0;
        public long SnapMin { get; set; } = 0;
        public long SnapMax { get; set; } = 0;

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

            CurrentSessionUpTimeMs = 0;
            CurrentSessionSuccessCount = 0;
            CurrentSessionMin = 0;
            CurrentSessionMax = 0;

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