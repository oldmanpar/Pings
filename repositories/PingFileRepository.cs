using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Pings.Models;

namespace Pings.Repositories
{
    /// <summary>
    /// ファイルの保存・読み込みに関する責務を持つクラス
    /// </summary>
    public class PingFileRepository
    {
        // Shift-JIS エンコーディング (Encoding 932)
        private readonly Encoding _encoding = Encoding.GetEncoding(932);

        /// <summary>
        /// 監視対象アドレスをCSVに保存
        /// </summary>
        public void SaveAddresses(string filePath, IEnumerable<PingMonitorItem> items)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false, _encoding))
            {
                sw.WriteLine("対象アドレス,Host名");
                foreach (var item in items.Where(i => !string.IsNullOrEmpty(i.対象アドレス) || !string.IsNullOrEmpty(i.Host名)))
                {
                    sw.WriteLine($"{item.対象アドレス},{item.Host名}");
                }
            }
        }

        /// <summary>
        /// 監視対象アドレスを読み込み (複雑なフォーマット判定ロジックをカプセル化)
        /// </summary>
        public List<PingMonitorItem> LoadAddresses(string filePath)
        {
            var newItems = new List<PingMonitorItem>();
            int index = 1;
            var allLines = new List<string>();

            using (StreamReader sr = new StreamReader(filePath, _encoding))
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
                        hostName = Regex.Replace(hostName, @"\s+", " ").Trim();
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
                newItems.Add(new PingMonitorItem(index++, address, hostName));
            }

            return newItems;
        }

        /// <summary>
        /// 監視結果(統計とログ)をCSV保存
        /// </summary>
        public void SavePingResults(string filePath, IEnumerable<PingMonitorItem> monitors, IEnumerable<DisruptionLogItem> logs, string startTime, string endTime, string interval, string timeout)
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            using (StreamWriter sw = new StreamWriter(filePath, true, _encoding))
            {
                sw.WriteLine("================================================================");
                sw.WriteLine($"Saved: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                sw.WriteLine();

                sw.WriteLine($"開始日時　：　{startTime}");
                sw.WriteLine($"終了日時　：　{endTime}");
                sw.WriteLine($"送信間隔　：　{interval} [ms]     タイムアウト　：　{timeout} [ms]");
                sw.WriteLine();

                sw.WriteLine("--- 監視統計 ---");
                string header = "ｽﾃｰﾀｽ,順番,対象アドレス,Host名,送信回数,失敗回数,連続失敗回数,連続失敗時間[hh:mm:ss],最大失敗時間[hh:mm:ss],時間[ms],平均値[ms],最小値[ms],最大値[ms],ジッタ①[ms],ジッタ②[ms],標準偏差";
                sw.WriteLine(header);

                foreach (var item in monitors.OrderBy(i => i.順番))
                {
                    string line = string.Join(",",
                        item.ステータス,
                        item.順番,
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
                        item.最大値ms,
                        item.Jitter1_MaxMin,
                        $"{item.Jitter2_PktPair:F1}",
                        $"{item.StdDev:F2}"
                    );
                    sw.WriteLine(line);
                }

                sw.WriteLine();
                sw.WriteLine();

                sw.WriteLine("--- 障害イベントログ ---");
                if (logs.Any())
                {
                    // 順序を UI に合わせて変更: Down前基本 → 復旧後基本 → Down前拡張 → 復旧後拡張
                    string logHeader = "対象アドレス,Host名,Down開始日時,復旧日時,失敗回数,失敗時間[hh:mm:ss]," +
                                       "Down前平均[ms],Down前最小[ms],Down前最大[ms]," +
                                       "復旧後平均[ms],復旧後最小[ms],復旧後最大[ms]," +
                                       "Down前ジッタ①,Down前ジッタ②,Down前標準偏差," +
                                       "復旧後ジッタ①,復旧後ジッタ②,復旧後標準偏差";
                    sw.WriteLine(logHeader);

                    foreach (var logItem in logs.OrderBy(i => i.復旧日時))
                    {
                        string logLine = string.Join(",",
                            logItem.対象アドレス,
                            logItem.Host名,
                            logItem.Down開始日時.ToString("yyyy/MM/dd HH:mm:ss"),
                            logItem.復旧日時.ToString("yyyy/MM/dd HH:mm:ss"),
                            logItem.失敗回数,
                            logItem.失敗時間mmss,
                            // Down前 基本統計
                            $"{logItem.Down前平均ms:F1}",
                            logItem.Down前最小ms,
                            logItem.Down前最大ms,
                            // 復旧後 基本統計
                            $"{logItem.復旧後平均ms:F1}",
                            logItem.復旧後最小ms,
                            logItem.復旧後最大ms,
                            // Down前 拡張統計
                            logItem.Down前Jitter1,
                            $"{logItem.Down前Jitter2:F1}",
                            $"{logItem.Down前StdDev:F2}",
                            // 復旧後 拡張統計
                            logItem.復旧後Jitter1,
                            $"{logItem.復旧後Jitter2:F1}",
                            $"{logItem.復旧後StdDev:F2}"
                        );
                        sw.WriteLine(logLine);
                    }
                }
                else
                {
                    sw.WriteLine("障害イベントは記録されていません。");
                }
                sw.WriteLine();
            }
        }

        /// <summary>
        /// Traceroute結果を保存
        /// </summary>
        public void SaveTracerouteResult(string folderPath, string address, string hostName, string content)
        {
            if (string.IsNullOrEmpty(content)) return;
            Directory.CreateDirectory(folderPath);

            string safeAddress = SanitizeFileName(address);
            string safeHost = SanitizeFileName(hostName);
            string fileName = Path.Combine(folderPath, $"Traceroute_result_{DateTime.Now:yyyyMMdd}_{safeAddress}_{safeHost}.log");

            using (var sw = new StreamWriter(fileName, true, _encoding))
            {
                sw.WriteLine("----------------------------------------------------------------");
                sw.WriteLine($"Saved: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                sw.WriteLine(content);
                sw.WriteLine();
            }
        }

        public string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "unknown";
            var invalid = Path.GetInvalidFileNameChars();
            var s = new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
            if (s.Length > 120) s = s.Substring(0, 120);
            return s;
        }
    }
}