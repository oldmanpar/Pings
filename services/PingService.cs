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
        /// 統計情報の更新ロジック (元PingMonitorItemにあったロジック)
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
                            Down前平均ms = item.SnapAvg,
                            Down前最小ms = item.SnapMin,
                            Down前最大ms = item.SnapMax,
                            復旧後平均ms = rtt,
                            復旧後最小ms = rtt,
                            復旧後最大ms = rtt
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

                if (item.CurrentSessionSuccessCount == 1)
                {
                    item.CurrentSessionMin = rtt;
                    item.CurrentSessionMax = rtt;
                }
                else
                {
                    if (rtt < item.CurrentSessionMin) item.CurrentSessionMin = rtt;
                    if (rtt > item.CurrentSessionMax) item.CurrentSessionMax = rtt;
                }

                item.平均値ms = (double)item.CurrentSessionUpTimeMs / item.CurrentSessionSuccessCount;
                item.最小値ms = item.CurrentSessionMin;
                item.最大値ms = item.CurrentSessionMax;

                if (item.ActiveLogItem != null)
                {
                    item.ActiveLogItem.復旧後平均ms = item.平均値ms;
                    item.ActiveLogItem.復旧後最小ms = item.最小値ms;
                    item.ActiveLogItem.復旧後最大ms = item.最大値ms;
                }
            }
            else
            {
                // ========== 失敗時 ==========
                if (!item.IsCurrentlyDown)
                {
                    item.ActiveLogItem = null;
                    item.SnapAvg = item.平均値ms;
                    item.SnapMin = item.最小値ms;
                    item.SnapMax = item.最大値ms;
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