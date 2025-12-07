using System;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using Pings.Models;

namespace Pings.Services
{
    /// <summary>
    /// Ping実行と統計計算、ログ生成を行うサービスクラス
    /// </summary>
    public class PingService
    {
        // ログ発生時に呼び出されるアクション
        private readonly Action<DisruptionLogItem> _onLogCreated;

        public PingService(Action<DisruptionLogItem> onLogCreated)
        {
            _onLogCreated = onLogCreated;
        }

        /// <summary>
        /// 単一アイテムのPingループを実行
        /// </summary>
        public async Task RunPingLoopAsync(PingMonitorItem item, CancellationToken token)
        {
            using (var ping = new Ping())
            {
                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        PingReply reply = await ping.SendPingAsync(item.対象アドレス, item.タイムアウトms);
                        UpdateStatistics(item, reply.RoundtripTime, reply.Status == IPStatus.Success);
                        await Task.Delay(item.送信間隔ms, token);
                    }
                    catch (TaskCanceledException)
                    {
                        break;
                    }
                    catch
                    {
                        UpdateStatistics(item, 0, false);
                        await Task.Delay(item.送信間隔ms, token);
                    }
                }
            }
        }

        /// <summary>
        /// 統計情報の更新ロジック
        /// </summary>
        private void UpdateStatistics(PingMonitorItem item, long rtt, bool success)
        {
            item.送信回数++;

            if (success)
            {
                // ========== 復旧時の処理 ==========
                if (item.IsCurrentlyDown)
                {
                    if (item.ContinuousDownStartTime.HasValue)
                    {
                        TimeSpan duration = DateTime.Now - item.ContinuousDownStartTime.Value;

                        DisruptionLogItem newLogItem = new DisruptionLogItem
                        {
                            対象アドレス = item.対象アドレス,
                            Host名 = item.Host名,
                            Down開始日時 = item.ContinuousDownStartTime.Value,
                            復旧日時 = DateTime.Now,
                            失敗時間mmss = duration.ToString(@"mm\:ss"),
                            失敗回数 = item.CurrentDisruptionFailureCount,

                            // Down前統計
                            Down前平均ms = item.SnapAvg,
                            Down前最小ms = item.SnapMin,
                            Down前最大ms = item.SnapMax,
                            Down前Jitter1 = item.SnapJitter1,
                            Down前Jitter2 = item.SnapJitter2,
                            Down前StdDev = item.SnapStdDev,

                            // 復旧後統計 (初期値として現在のRTTを入れる)
                            復旧後平均ms = rtt,
                            復旧後最小ms = rtt,
                            復旧後最大ms = rtt,
                            復旧後Jitter1 = 0,
                            復旧後Jitter2 = 0,
                            復旧後StdDev = 0
                        };

                        // コールバック経由でログ登録
                        _onLogCreated?.Invoke(newLogItem);
                        item.ActiveLogItem = newLogItem;
                    }

                    // 統計リセット
                    item.CurrentSessionUpTimeMs = 0;
                    item.CurrentSessionSuccessCount = 0;
                    item.CurrentSessionMin = 0;
                    item.CurrentSessionMax = 0;

                    // ★追加: 拡張統計リセット
                    item.CurrentSessionSumSquares = 0;
                    item.PreviousRtt = -1; // 復旧直後はパケットペアがない状態にする
                    item.JitterDiffSum = 0;
                    item.JitterDiffCount = 0;
                    item.Jitter1_MaxMin = 0;
                    item.Jitter2_PktPair = 0.0;
                    item.StdDev = 0.0;

                    item.CurrentDisruptionFailureCount = 0;
                    item.IsCurrentlyDown = false;
                }

                // ========== 共通: 成功時 ==========
                item.時間ms = rtt;
                item.IsUp = true;
                item.連続失敗回数 = 0;
                item.連続失敗時間s = "";
                item.ContinuousDownStartTime = null;
                item.ステータス = item.HasBeenDown ? "復旧" : "OK";

                item.CurrentSessionSuccessCount++;
                item.CurrentSessionUpTimeMs += rtt;
                item.CurrentSessionSumSquares += (rtt * rtt); // 二乗和加算

                // 最小・最大更新
                if (item.CurrentSessionSuccessCount == 1)
                {
                    item.CurrentSessionMin = rtt;
                    item.CurrentSessionMax = rtt;
                    // 1回目なのでJitter2(差分)は計算できない
                }
                else
                {
                    if (rtt < item.CurrentSessionMin) item.CurrentSessionMin = rtt;
                    if (rtt > item.CurrentSessionMax) item.CurrentSessionMax = rtt;

                    // Jitter2 (Packet Pair) 計算
                    // 前回のRTTが有効(=成功していた)場合のみ計算
                    if (item.PreviousRtt >= 0)
                    {
                        long diff = Math.Abs(rtt - item.PreviousRtt);
                        item.JitterDiffSum += diff;
                        item.JitterDiffCount++;
                        item.Jitter2_PktPair = (double)item.JitterDiffSum / item.JitterDiffCount;
                    }
                }

                // 前回のRTTを更新
                item.PreviousRtt = rtt;

                // 平均計算
                item.平均値ms = (double)item.CurrentSessionUpTimeMs / item.CurrentSessionSuccessCount;

                // Jitter1 (Max - Min) 計算
                item.Jitter1_MaxMin = item.CurrentSessionMax - item.CurrentSessionMin;

                // 標準偏差 計算 (Population StdDev)
                if (item.CurrentSessionSuccessCount > 0)
                {
                    // 分散 = (二乗和 / N) - (平均^2)
                    double mean = item.平均値ms;
                    double variance = (item.CurrentSessionSumSquares / item.CurrentSessionSuccessCount) - (mean * mean);
                    // 浮動小数点の誤差でマイナスになるのを防ぐ
                    item.StdDev = Math.Sqrt(Math.Max(0, variance));
                }
                else
                {
                    item.StdDev = 0.0;
                }

                // UI表示用の基本値更新
                item.最小値ms = item.CurrentSessionMin;
                item.最大値ms = item.CurrentSessionMax;

                // ログのアクティブ行がある場合、復旧後の統計をリアルタイム更新
                if (item.ActiveLogItem != null)
                {
                    item.ActiveLogItem.復旧後平均ms = item.平均値ms;
                    item.ActiveLogItem.復旧後最小ms = item.最小値ms;
                    item.ActiveLogItem.復旧後最大ms = item.最大値ms;
                    item.ActiveLogItem.復旧後Jitter1 = item.Jitter1_MaxMin;
                    item.ActiveLogItem.復旧後Jitter2 = item.Jitter2_PktPair;
                    item.ActiveLogItem.復旧後StdDev = item.StdDev;
                }
            }
            else
            {
                // ========== 失敗時 ==========
                if (!item.IsCurrentlyDown)
                {
                    // Down開始時点のスナップショットを保存
                    item.ActiveLogItem = null;
                    item.SnapAvg = item.平均値ms;
                    item.SnapMin = item.最小値ms;
                    item.SnapMax = item.最大値ms;

                    // ★追加: 拡張統計スナップショット
                    item.SnapJitter1 = item.Jitter1_MaxMin;
                    item.SnapJitter2 = item.Jitter2_PktPair;
                    item.SnapStdDev = item.StdDev;

                    item.ContinuousDownStartTime = DateTime.Now;
                    item.IsCurrentlyDown = true;
                    item.CurrentDisruptionFailureCount = 0;
                }

                item.失敗回数++;
                item.CurrentDisruptionFailureCount++;
                item.時間ms = 0;
                item.ステータス = "Down";
                item.IsUp = false;
                item.HasBeenDown = true;
                item.連続失敗回数++;

                // 失敗したので、Jitter2用の「前回のRTT」チェーンは切れる
                item.PreviousRtt = -1;

                TimeSpan currentDuration = DateTime.Now - item.ContinuousDownStartTime.Value;
                item.連続失敗時間s = currentDuration.ToString(@"mm\:ss");

                if (currentDuration > item.MaxDisruptionDuration)
                {
                    item.MaxDisruptionDuration = currentDuration;
                    item.最大失敗時間s = item.MaxDisruptionDuration.ToString(@"mm\:ss");
                }
            }
        }
    }
}