using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Pings.Models;

namespace Pings
{
    public partial class Form1
    {
        // ====================================================================
        // Ping操作ボタン
        // ====================================================================

        private void BtnPingStart_Click(object sender, EventArgs e)
        {
            StopMonitoring();
            ClearView();
            dgvMonitor.EndEdit();

            // 空行削除
            var validItems = monitorList
                .Where(i => !string.IsNullOrWhiteSpace(i.対象アドレス) || !string.IsNullOrWhiteSpace(i.Host名))
                .ToList();

            monitorList.ListChanged -= MonitorList_ListChanged;
            monitorList.Clear();
            foreach (var item in validItems) monitorList.Add(item);
            monitorList.ListChanged += MonitorList_ListChanged;

            _nextIndex = monitorList.Any() ? monitorList.Max(i => i.順番) + 1 : 1;

            StartMonitoring();
        }

        private void BtnStop_Click(object sender, EventArgs e) => StopMonitoring();

        private void BtnClear_Click(object sender, EventArgs e)
        {
            StopMonitoring();
            ClearView();
            _allowEditAfterStop = true;
            _allowIntervalTimeoutEdit = true;
            UpdateUiState(UiState.Initial);
            btnPingStart.Enabled = true;
            dgvMonitor.AllowUserToDeleteRows = true;
        }

        private void BtnSaveResult_Click(object sender, EventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中は結果を保存できません。", "保存不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string baseFolder = AppDomain.CurrentDomain.BaseDirectory;
            string defaultPath = Path.Combine(baseFolder, "Ping_Result", $"Ping_Result_{DateTime.Now:yyyyMMdd}.csv");

            try
            {
                ExecuteSave(defaultPath);
                MessageBox.Show($"監視結果を保存しました：\n{defaultPath}", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch
            {
                // 既定パス失敗時は手動保存へ
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSVファイル (*.csv)|*.csv|All|*.*";
                sfd.FileName = Path.GetFileName(defaultPath);
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExecuteSave(sfd.FileName);
                        MessageBox.Show($"監視結果を保存しました：\n{sfd.FileName}", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"保存エラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            void ExecuteSave(string path)
            {
                _repository.SavePingResults(path, monitorList, disruptionLogList, txtStartTime.Text, txtEndTime.Text, cmbInterval.Text, cmbTimeout.Text);
                _autoSaveFailed = false;
                _allowEditAfterStop = true;
                _allowIntervalTimeoutEdit = true;
                UpdateUiState(UiState.Stopped);
                btnPingStart.Enabled = true;
            }
        }

        private void BtnSaveAddress_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSVファイル (*.csv)|*.csv|All|*.*";
                sfd.FileName = "PingTargets.csv";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _repository.SaveAddresses(sfd.FileName, monitorList);
                        MessageBox.Show("監視対象アドレスを保存しました。", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"保存エラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnLoadAddress_Click(object sender, EventArgs e)
        {
            if (cts != null) return;

            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "アドレスファイル (*.csv;*.txt;*.log)|*.csv;*.txt;*.log|All|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var items = _repository.LoadAddresses(ofd.FileName);
                        monitorList = new BindingList<PingMonitorItem>(items);
                        monitorList.ListChanged += MonitorList_ListChanged;
                        dgvMonitor.DataSource = monitorList;
                        _nextIndex = items.Count + 1;
                        UpdateUiState(UiState.Initial);
                        MessageBox.Show($"監視対象アドレスを {items.Count}件 読み込みました。", "読込完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"読込エラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // ====================================================================
        // フォーム・メニューイベント
        // ====================================================================

        private void BtnExit_Click(object sender, EventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中は終了できません。", "終了不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            StopMonitoring();
            Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中は終了できません。", "終了不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
            else StopMonitoring();
        }

        private void VersionItem_Click(object sender, EventArgs e)
        {
            using (var ab = new AboutBox()) ab.ShowDialog(this);
        }

        private void HelpUsageItem_Click(object sender, EventArgs e)
        {
            // モードレスで単一インスタンス表示（既に開いていれば前面へ）
            if (helpForm == null || helpForm.IsDisposed)
            {
                helpForm = new HelpForm();
                helpForm.Show(this);
            }
            else
            {
                helpForm.Activate();
            }
        }

        // ====================================================================
        // DataGridView イベント
        // ====================================================================

        private void MonitorList_ListChanged(object sender, ListChangedEventArgs e)
        {
            if (e.ListChangedType == ListChangedType.ItemAdded ||
                e.ListChangedType == ListChangedType.ItemDeleted ||
                e.ListChangedType == ListChangedType.Reset)
                RecalculateOrderNumbers();
        }

        private void DgvMonitor_UserDeletedRow(object sender, DataGridViewRowEventArgs e) => RecalculateOrderNumbers();
        private void DgvMonitor_RowValidated(object sender, DataGridViewCellEventArgs e) => RecalculateOrderNumbers();

        private void DgvMonitor_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            if (e.Row.IsNewRow) e.Row.Cells[2].Value = _nextIndex++;
        }

        private void DgvMonitor_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0 && dgvMonitor.Columns[e.ColumnIndex].DataPropertyName == "Trace")
                UpdateTracerouteButtons();
        }

        private void DgvMonitor_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvMonitor.IsCurrentCellDirty)
                dgvMonitor.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void DgvMonitor_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex >= 0 && dgvMonitor.Columns[e.ColumnIndex].DataPropertyName == "Trace")
            {
                bool val = monitorList.Any(i => !i.Trace);
                foreach (var item in monitorList) item.Trace = val;
                monitorList.ResetBindings();
                UpdateTracerouteButtons();
            }
        }

        private void DgvMonitor_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvMonitor.Rows[e.RowIndex].DataBoundItem is PingMonitorItem item)
            {
                if (string.IsNullOrWhiteSpace(item.対象アドレス)) return;
                var terminal = new PingTerminalForm(item.対象アドレス, item.Host名);
                terminal.Show(this);
            }
        }

        private void dgvLog_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var col = dgvLog.Columns[e.ColumnIndex];
            var dir = (currentSortColumn == col && currentSortDirection == System.ComponentModel.ListSortDirection.Ascending)
                ? System.ComponentModel.ListSortDirection.Descending
                : System.ComponentModel.ListSortDirection.Ascending;
            var prop = System.ComponentModel.TypeDescriptor.GetProperties(typeof(DisruptionLogItem)).Find(col.DataPropertyName, true);
            if (prop != null)
            {
                if (currentSortColumn != null) currentSortColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                disruptionLogList.Sort(prop, dir);
                col.HeaderCell.SortGlyphDirection = (dir == System.ComponentModel.ListSortDirection.Ascending) ? SortOrder.Ascending : SortOrder.Descending;
                currentSortColumn = col;
                currentSortDirection = dir;
            }
        }
    }
}
