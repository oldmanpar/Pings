using System;

namespace Pings.Models
{
    /// <summary>
    /// 発生した障害イベント（Down → 復旧）を記録するデータクラス
    /// </summary>
    public class DisruptionLogItem
    {
        public string 対象アドレス { get; set; }
        public string Host名 { get; set; }
        public DateTime Down開始日時 { get; set; }
        public DateTime 復旧日時 { get; set; }
        public string 失敗時間mmss { get; set; }
        public int 失敗回数 { get; set; }

        // 統計値
        public double Down前平均ms { get; set; }
        public long Down前最小ms { get; set; }
        public long Down前最大ms { get; set; }

        public double 復旧後平均ms { get; set; }
        public long 復旧後最小ms { get; set; }
        public long 復旧後最大ms { get; set; }

        // ↓↓↓ 以下を追加
        public long Down前Jitter1 { get; set; }
        public double Down前Jitter2 { get; set; }
        public double Down前StdDev { get; set; }
        public long 復旧後Jitter1 { get; set; }
        public double 復旧後Jitter2 { get; set; }
        public double 復旧後StdDev { get; set; }
    }
}