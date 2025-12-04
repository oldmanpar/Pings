using System;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Pings.Services
{
    /// <summary>
    /// Tracerouteの実行を行うサービスクラス
    /// </summary>
    public class TracerouteService
    {
        public async Task RunTracerouteAsync(string address, int timeoutMs, CancellationToken token, Action<string> onOutput)
        {
            // Windows tracert の引数: -d (名前解決なし), -w (タイムアウトms)
            int tryTimeoutMs = Math.Max(100, timeoutMs);
            string arguments = $"-d -w {tryTimeoutMs} {address}";

            onOutput($"--- tracert {address} (timeout={tryTimeoutMs}ms) ---\r\n");

            var psi = new ProcessStartInfo
            {
                FileName = "tracert",
                Arguments = arguments,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.GetEncoding(932)
            };

            using (var proc = new Process { StartInfo = psi })
            {
                try
                {
                    proc.Start();

                    using (var reader = proc.StandardOutput)
                    {
                        using (token.Register(() => { try { if (!proc.HasExited) proc.Kill(); } catch { } }))
                        {
                            while (!reader.EndOfStream && !token.IsCancellationRequested)
                            {
                                string line;
                                try { line = await reader.ReadLineAsync().ConfigureAwait(false); }
                                catch (Exception ex) { line = $"(出力取得エラー: {ex.Message})"; }

                                if (line == null) break;
                                onOutput(line + Environment.NewLine);
                            }
                        }
                    }

                    if (!proc.HasExited)
                    {
                        try { proc.WaitForExit(1000); } catch { }
                    }
                }
                catch (Exception ex)
                {
                    onOutput($"(tracert 実行エラー: {ex.Message})\r\n");
                }
                finally
                {
                    if (!proc.HasExited) try { proc.Kill(); } catch { }
                }
            }
        }
    }
}