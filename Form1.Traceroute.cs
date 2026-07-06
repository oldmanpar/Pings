using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pings
{
    public partial class Form1
    {
        // ====================================================================
        // Traceroute ハンドラ
        // ====================================================================

        private async void BtnTraceroute_Click(object sender, EventArgs e)
        {
            var targets = monitorList.Where(i => i.Trace && !string.IsNullOrWhiteSpace(i.対象アドレス))
                                     .Select(i => i.対象アドレス).Distinct().ToList();

            if (!targets.Any())
            {
                MessageBox.Show("Traceroute対象がありません。", "実行不可", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _isTracerouteRunning = true;
            UpdateUiState(CurrentUiState);
            UpdateTracerouteButtons();
            tabControl.SelectedTab = traceroutePage;

            SetupTracerouteUI(targets);

            AppendTracerouteOutputToAll("================================================================\r\n", false);
            AppendTracerouteOutputToAll($"== Traceroute: {DateTime.Now:yyyy/MM/dd HH:mm:ss} ==\r\n", false);

            tracerouteCts?.Dispose();
            tracerouteCts = CancellationTokenSource.CreateLinkedTokenSource(cts?.Token ?? CancellationToken.None);

            _tracerouteCompletion = new System.Collections.Concurrent.ConcurrentDictionary<string, bool>();
            _tracerouteStoppedByUser = new System.Collections.Concurrent.ConcurrentDictionary<string, bool>();
            foreach (var t in targets)
            {
                _tracerouteCompletion[t] = false;
                _tracerouteStoppedByUser[t] = false;
            }

            try
            {
                int timeout = 4000;
                int.TryParse(cmbTimeout.Text, out timeout);

                bool noResolve = mnuTracerouteNoResolve != null && mnuTracerouteNoResolve.Checked;

                var tasks = targets.Select(async address =>
                {
                    await tracerouteSemaphore.WaitAsync(tracerouteCts.Token);
                    try
                    {
                        await _tracerouteService.RunTracerouteAsync(address, timeout, noResolve, tracerouteCts.Token, (text) =>
                        {
                            AppendTracerouteOutput(address, text, false);
                        });
                    }
                    finally
                    {
                        tracerouteSemaphore.Release();
                        if (_tracerouteCompletion != null) _tracerouteCompletion[address] = true;

                        bool stopped = _tracerouteStoppedByUser.ContainsKey(address) && _tracerouteStoppedByUser[address];
                        if (!stopped && !tracerouteCts.IsCancellationRequested)
                            AppendTracerouteOutput(address, "=== Traceroute 完了 ===\r\n\r\n", false);
                    }
                });

                await Task.WhenAll(tasks);
            }
            catch (OperationCanceledException)
            {
                HandleTracerouteCancel(targets);
            }
            finally
            {
                if (tracerouteCts != null && tracerouteCts.IsCancellationRequested) HandleTracerouteCancel(targets);

                AppendTracerouteOutputToAll("================================================================\r\n\r\n", false);
                _isTracerouteRunning = false;
                UpdateUiState(CurrentUiState);
                UpdateTracerouteButtons();

                if (mnuAutoSaveTraceroute.Checked) SaveTracerouteOutputsAutoAppend();
                tracerouteCts?.Dispose();
                tracerouteCts = null;
            }
        }

        private void BtnStopTraceroute_Click(object sender, EventArgs e)
        {
            if (_tracerouteCompletion != null)
            {
                foreach (var kv in _tracerouteCompletion)
                {
                    if (!kv.Value && !_tracerouteStoppedByUser[kv.Key])
                    {
                        AppendTracerouteOutput(kv.Key, "=== Tracerouteは途中で停止されました。 ===\r\n\r\n", false);
                        _tracerouteStoppedByUser[kv.Key] = true;
                    }
                }
            }
            tracerouteCts?.Cancel();
            _isTracerouteRunning = false;
            UpdateTracerouteButtons();
        }

        private void BtnSaveTraceroute_Click(object sender, EventArgs e)
        {
            string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Traceroute_Result");
            try
            {
                foreach (var kv in tracerouteTextBoxes)
                {
                    string host = monitorList.FirstOrDefault(i => i.対象アドレス == kv.Key)?.Host名 ?? "";
                    _repository.SaveTracerouteResult(folder, kv.Key, host, kv.Value.Text);
                    kv.Value.Clear();
                }
                _tracerouteHasOutput = false;
                MessageBox.Show("保存しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            UpdateTracerouteButtons();
        }

        private void BtnClearTraceroute_Click(object sender, EventArgs e)
        {
            if (tracerouteTextBoxes != null)
            {
                foreach (var tb in tracerouteTextBoxes.Values) tb.Clear();
                _tracerouteHasOutput = false;
                UpdateTracerouteButtons();
            }
        }

        // ====================================================================
        // Traceroute ヘルパー
        // ====================================================================

        private void SetupTracerouteUI(List<string> targets)
        {
            bool reuse = tracerouteTextBoxes != null && tracerouteTextBoxes.Count == targets.Count
                         && targets.All(t => tracerouteTextBoxes.ContainsKey(t));

            if (!reuse)
            {
                // 多拠点時のレイアウト再計算を止めてから一括構築する
                traceroutePanel.SuspendLayout();
                try
                {
                    tracerouteTextBoxes.Clear();
                    traceroutePanel.Controls.Clear();
                    _tracerouteHasOutput = false;
                    traceroutePanel.ColumnCount = Math.Max(1, targets.Count);
                    traceroutePanel.RowCount = 1;
                    traceroutePanel.ColumnStyles.Clear();
                    for (int i = 0; i < traceroutePanel.ColumnCount; i++)
                        traceroutePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, TracerouteColumnWidth));

                    for (int i = 0; i < targets.Count; i++)
                    {
                        string address = targets[i];
                        string hostName = monitorList.FirstOrDefault(m => m.対象アドレス == address)?.Host名 ?? "";

                        var container = new Panel { Dock = DockStyle.Fill, Padding = new Padding(2), Width = TracerouteColumnWidth };
                        var lbl = new Label { Text = $"{address}_{hostName}", Dock = DockStyle.Top, Height = 18, AutoEllipsis = true };
                        var tb = new TextBox
                        {
                            Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
                            WordWrap = false, Dock = DockStyle.Fill, BackColor = System.Drawing.SystemColors.Window,
                            Font = new Font(FontFamily.GenericMonospace, 9f)
                        };

                        container.Controls.Add(tb);
                        container.Controls.Add(lbl);
                        traceroutePanel.Controls.Add(container, i, 0);
                        tracerouteTextBoxes[address] = tb;
                    }
                }
                finally
                {
                    traceroutePanel.ResumeLayout(true);
                }
            }
        }

        private void HandleTracerouteCancel(List<string> targets)
        {
            foreach (var addr in targets)
            {
                bool done = _tracerouteCompletion != null && _tracerouteCompletion.TryGetValue(addr, out bool d) && d;
                bool stopped = _tracerouteStoppedByUser != null && _tracerouteStoppedByUser.TryGetValue(addr, out bool s) && s;
                if (!done && !stopped)
                {
                    AppendTracerouteOutput(addr, "=== Tracerouteは途中で停止されました。 ===\r\n\r\n", false);
                    if (_tracerouteStoppedByUser != null) _tracerouteStoppedByUser[addr] = true;
                }
            }
        }

        private void SaveTracerouteOutputsAutoAppend()
        {
            string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Traceroute_Result");
            foreach (var kv in tracerouteTextBoxes)
            {
                string host = monitorList.FirstOrDefault(i => i.対象アドレス == kv.Key)?.Host名 ?? "";
                _repository.SaveTracerouteResult(folder, kv.Key, host, kv.Value.Text);
            }
        }

        private void AppendTracerouteOutput(string address, string text, bool keepTimestamp)
        {
            if (tracerouteTextBoxes != null && tracerouteTextBoxes.ContainsKey(address))
            {
                var tb = tracerouteTextBoxes[address];
                _tracerouteHasOutput = true;
                // BeginInvoke: ワーカースレッドをUIの応答待ちでブロックしない。
                // ボタン更新は毎行行わず、開始・終了時にまとめて行う
                if (tb.InvokeRequired)
                    tb.BeginInvoke(new Action(() => { tb.AppendText(text); tb.ScrollToCaret(); }));
                else
                {
                    tb.AppendText(text);
                    tb.ScrollToCaret();
                }
            }
        }

        private void AppendTracerouteOutputToAll(string text, bool keepTimestamp)
        {
            if (tracerouteTextBoxes != null)
            {
                foreach (var addr in tracerouteTextBoxes.Keys)
                    AppendTracerouteOutput(addr, text, keepTimestamp);
            }
        }
    }
}
