using System;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using Pings.Models;

namespace Pings.Services
{
    /// <summary>
    /// Ping ループの実行と、復旧ログ発生時のコールバック通知を担うサービスクラス。
    /// 統計計算は PingMonitorItem 側に委譲している。
    /// </summary>
    public class PingService
    {
        private readonly Action<DisruptionLogItem> _onLogCreated;

        public PingService(Action<DisruptionLogItem> onLogCreated)
        {
            _onLogCreated = onLogCreated;
        }

        /// <summary>単一アイテムの Ping ループを実行する</summary>
        public async Task RunPingLoopAsync(PingMonitorItem item, CancellationToken token)
        {
            using (var ping = new Ping())
            {
                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        PingReply reply = await ping.SendPingAsync(item.対象アドレス, item.タイムアウトms);

                        if (reply.Status == IPStatus.Success)
                        {
                            var logItem = item.RecordSuccess(reply.RoundtripTime);
                            if (logItem != null) _onLogCreated?.Invoke(logItem);
                        }
                        else
                        {
                            item.RecordFailure();
                        }

                        await Task.Delay(item.送信間隔ms, token);
                    }
                    catch (TaskCanceledException)
                    {
                        break;
                    }
                    catch
                    {
                        item.RecordFailure();
                        await Task.Delay(item.送信間隔ms, token);
                    }
                }
            }
        }
    }
}
