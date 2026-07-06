using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Pings.Models;

namespace Pings
{
    public partial class Form1
    {
        // ====================================================================
        // 監視の開始・停止・クリア
        // ====================================================================

        private void StartMonitoring()
        {
            _allowEditAfterStop = false;
            _allowIntervalTimeoutEdit = false;
            cts = new System.Threading.CancellationTokenSource();
            txtStartTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            txtEndTime.Text = "";

            int interval = int.Parse(cmbInterval.Text);
            int timeout = int.Parse(cmbTimeout.Text);

            UpdateUiState(UiState.Running);
            uiUpdateTimer.Start();
            ResetLogSortIndicators();

            foreach (var item in monitorList.Where(i => !string.IsNullOrEmpty(i.対象アドレス) && i.順番 > 0))
            {
                item.送信間隔ms = interval;
                item.タイムアウトms = timeout;
                item.ResetData();
                Task.Run(() => _pingService.RunPingLoopAsync(item, cts.Token));
            }
            dgvMonitor.AllowUserToDeleteRows = false;
        }

        private void StopMonitoring()
        {
            if (cts != null)
            {
                uiUpdateTimer.Stop();
                cts.Cancel();
                cts.Dispose();
                cts = null;
                txtEndTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                _allowEditAfterStop = false;

                if (mnuAutoSaveAllPing.Checked)
                {
                    try
                    {
                        string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ping_Result", $"Ping_Result_{DateTime.Now:yyyyMMdd}.csv");
                        _repository.SavePingResults(path, monitorList, disruptionLogList, txtStartTime.Text, txtEndTime.Text, cmbInterval.Text, cmbTimeout.Text);
                        MessageBox.Show($"自動保存しました：\n{path}", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        _autoSaveFailed = true;
                        MessageBox.Show($"自動保存失敗: {ex.Message}\n手動保存してください。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                UpdateUiState(UiState.Stopped);
                btnPingStart.Enabled = false;
            }
        }

        private void ClearView()
        {
            foreach (var item in monitorList) item.ResetData();
            disruptionLogList.Clear();
            txtStartTime.Text = "";
            txtEndTime.Text = "";
            monitorList.ResetBindings();
            UpdateUiState(UiState.Initial);
            ResetLogSortIndicators();
            if (tracerouteTextBoxes != null)
            {
                tracerouteTextBoxes.Clear();
                traceroutePanel.Controls.Clear();
            }
        }

        // ====================================================================
        // UI状態管理
        // ====================================================================

        private void UpdateUiState(UiState state)
        {
            if (MainMenuStrip != null && MainMenuStrip.Items.Count > 0)
            {
                var fileMenu = MainMenuStrip.Items[0] as ToolStripMenuItem;
                if (fileMenu != null)
                {
                    fileMenu.DropDownItems[0].Enabled = (state == UiState.Initial || state == UiState.Stopped);
                    fileMenu.DropDownItems[1].Enabled = (state == UiState.Initial);
                }
            }

            dgvMonitor.ReadOnly = false;
            dgvMonitor.AllowUserToAddRows = (state == UiState.Initial);

            foreach (DataGridViewColumn col in dgvMonitor.Columns)
            {
                if (state == UiState.Initial)
                {
                    col.ReadOnly = !(col.DataPropertyName == "対象アドレス" || col.DataPropertyName == "Host名" || col.DataPropertyName == "Trace");
                }
                else if (state == UiState.Running)
                {
                    col.ReadOnly = col.DataPropertyName != "Trace";
                }
                else // Stopped
                {
                    if (col.DataPropertyName == "Trace") col.ReadOnly = false;
                    else if (col.DataPropertyName == "対象アドレス" || col.DataPropertyName == "Host名") col.ReadOnly = !_allowEditAfterStop;
                    else col.ReadOnly = true;
                }
            }

            if (state != UiState.Running) btnTraceroute.Enabled = false;
            UpdateTracerouteButtons();

            bool allowCombos = !_isTracerouteRunning && state != UiState.Running && _allowIntervalTimeoutEdit;
            cmbInterval.Enabled = allowCombos;
            cmbTimeout.Enabled = allowCombos;

            bool autoSave = mnuAutoSaveAllPing?.Checked ?? false;

            switch (state)
            {
                case UiState.Initial:
                    btnPingStart.Enabled = true;
                    btnStop.Enabled = false;
                    btnClear.Enabled = false;
                    btnSave.Enabled = false;
                    break;
                case UiState.Running:
                    btnPingStart.Enabled = false;
                    btnStop.Enabled = true;
                    btnClear.Enabled = false;
                    btnSave.Enabled = false;
                    break;
                case UiState.Stopped:
                    btnPingStart.Enabled = true;
                    btnStop.Enabled = false;
                    btnClear.Enabled = true;
                    btnSave.Enabled = !autoSave || _autoSaveFailed;
                    break;
            }

            if (_isTracerouteRunning)
            {
                btnPingStart.Enabled = false;
                btnStop.Enabled = false;
            }
        }

        private void UpdateTracerouteButtons()
        {
            if (btnTraceroute == null) return;
            if (InvokeRequired) { Invoke((MethodInvoker)UpdateTracerouteButtons); return; }

            bool hasChecked = monitorList?.Any(i => i.Trace && !string.IsNullOrWhiteSpace(i.対象アドレス)) ?? false;
            btnTraceroute.Enabled = hasChecked && !_isTracerouteRunning;

            bool hasOutput = tracerouteTextBoxes != null && tracerouteTextBoxes.Values.Any(t => !string.IsNullOrEmpty(t.Text));
            bool autoSave = mnuAutoSaveTraceroute?.Checked ?? false;

            btnSaveTraceroute.Enabled = !_isTracerouteRunning && hasOutput && !autoSave;
            btnClearTraceroute.Enabled = !_isTracerouteRunning && hasOutput;
            btnStopTraceroute.Enabled = _isTracerouteRunning;
        }

        // ====================================================================
        // グリッド更新・ソート・行番号
        // ====================================================================

        private void AddDisruptionLogItem(DisruptionLogItem item)
        {
            disruptionLogList.Add(item);
            if (currentSortColumn == null ||
                (currentSortColumn.DataPropertyName == "復旧日時" && currentSortDirection == System.ComponentModel.ListSortDirection.Ascending))
            {
                if (dgvLog.RowCount > 0) dgvLog.FirstDisplayedScrollingRowIndex = dgvLog.RowCount - 1;
            }
        }

        private void UiUpdateTimer_Tick(object sender, EventArgs e)
        {
            if (dgvMonitor != null && dgvMonitor.IsHandleCreated) dgvMonitor.Refresh();
            if (dgvLog != null && dgvLog.IsHandleCreated) dgvLog.Refresh();
        }

        private void RecalculateOrderNumbers()
        {
            if (_recalculating) return;
            _recalculating = true;
            try
            {
                monitorList.ListChanged -= MonitorList_ListChanged;
                for (int i = 0; i < monitorList.Count; i++) monitorList[i].順番 = i + 1;
                _nextIndex = monitorList.Count + 1;
                monitorList.ResetBindings();
                monitorList.ListChanged += MonitorList_ListChanged;
            }
            finally
            {
                _recalculating = false;
            }
        }

        private void ResetLogSortIndicators()
        {
            if (currentSortColumn != null)
            {
                currentSortColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                currentSortColumn = null;
                currentSortDirection = System.ComponentModel.ListSortDirection.Ascending;
            }
        }
    }
}
