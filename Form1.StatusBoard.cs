using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Pings.Models;

namespace Pings
{
    public partial class Form1
    {
        // ====================================================================
        // ステータスボード
        // 監視対象1件 = 1行チップ。縦優先（上→下、列が埋まったら右へ折り返し）
        // で配置し、多拠点をスクロールなしで一望できるようにする。
        // ====================================================================

        /// <summary>チップ1枚分のコントロールと差分更新用の前回値</summary>
        private class StatusChip
        {
            public PingMonitorItem Item;
            public Panel Panel;
            public Label Name;
            public Label Value;
            public string LastValue;
            public Color LastColor;
        }

        private readonly List<StatusChip> _statusChips = new List<StatusChip>();

        private const int ChipWidth = 236;
        private const int ChipHeight = 22;

        /// <summary>ステータスボードタブを構築して返す（Setup から呼ぶ）</summary>
        private TabPage CreateStatusBoardTab()
        {
            statusBoardPage = new TabPage("ステータスボード");

            // サマリーバー（OK/Down/復旧の件数）
            var summaryFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 30,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(6, 4, 6, 0)
            };
            lblBoardTotal    = CreateSummaryLabel("監視中: 0", SystemColors.ControlLight);
            lblBoardOk       = CreateSummaryLabel("OK: 0", Color.LightGreen);
            lblBoardDown     = CreateSummaryLabel("Down: 0", Color.MistyRose);
            lblBoardRecovery = CreateSummaryLabel("復旧: 0", Color.LightSkyBlue);
            summaryFlow.Controls.AddRange(new Control[] { lblBoardTotal, lblBoardOk, lblBoardDown, lblBoardRecovery });

            // チップ配置エリア。TopDown + 高さ固定で縦優先の折り返しを実現し、
            // 横方向のあふれのみ親パネルのスクロールで対応する
            statusBoardHost = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(4) };
            statusBoardFlow = new FlowLayoutPanel
            {
                Location = new Point(4, 4),
                FlowDirection = FlowDirection.TopDown,
                WrapContents = true,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            statusBoardHost.Controls.Add(statusBoardFlow);
            statusBoardHost.Resize += (s, e) => SyncStatusBoardHeight();

            statusBoardPage.Controls.Add(statusBoardHost);
            statusBoardPage.Controls.Add(summaryFlow);
            return statusBoardPage;
        }

        private Label CreateSummaryLabel(string text, Color backColor)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                BackColor = backColor,
                Padding = new Padding(8, 4, 8, 4),
                Margin = new Padding(3)
            };
        }

        /// <summary>FlowLayoutPanel の折り返し高さをホストパネルの高さに合わせる</summary>
        private void SyncStatusBoardHeight()
        {
            if (statusBoardHost == null || statusBoardFlow == null) return;
            int h = statusBoardHost.ClientSize.Height - statusBoardHost.Padding.Vertical;
            if (h < ChipHeight) h = ChipHeight;
            // MaximumSize.Height で折り返し位置を制御（幅0 = 無制限）
            statusBoardFlow.MaximumSize = new Size(0, h);
            statusBoardFlow.MinimumSize = new Size(0, h);
        }

        /// <summary>監視開始時にチップを作り直す（StartMonitoring から呼ぶ）</summary>
        private void BuildStatusBoard()
        {
            statusBoardFlow.SuspendLayout();
            statusBoardFlow.Controls.Clear();
            _statusChips.Clear();

            foreach (var item in monitorList.Where(i => !string.IsNullOrEmpty(i.対象アドレス) && i.順番 > 0))
            {
                var chip = new StatusChip { Item = item };

                chip.Panel = new Panel
                {
                    Width = ChipWidth,
                    Height = ChipHeight,
                    Margin = new Padding(2),
                    BackColor = SystemColors.Control
                };
                // 列幅を固定してアドレス・Host名の開始位置を全チップで揃える
                // | 順番(22) | 対象アドレス(102) | Host名(47) | RTT/状態(50) |
                var lblNum = new Label
                {
                    Text = item.順番.ToString(),
                    Location = new Point(3, 3),
                    Size = new Size(22, 16),
                    TextAlign = ContentAlignment.MiddleRight
                };
                var lblAddr = new Label
                {
                    Text = item.対象アドレス,
                    Location = new Point(27, 3),
                    Size = new Size(102, 16),
                    AutoEllipsis = true
                };
                chip.Name = new Label
                {
                    Text = item.Host名,
                    Location = new Point(131, 3),
                    Size = new Size(47, 16),
                    AutoEllipsis = true
                };
                chip.Value = new Label
                {
                    Text = "",
                    Location = new Point(181, 3),
                    Size = new Size(50, 16),
                    TextAlign = ContentAlignment.MiddleRight
                };
                chip.Panel.Controls.Add(lblNum);
                chip.Panel.Controls.Add(lblAddr);
                chip.Panel.Controls.Add(chip.Name);
                chip.Panel.Controls.Add(chip.Value);

                // ダブルクリックで監視統計タブの該当行へ移動
                var target = item;
                chip.Panel.DoubleClick += (s, e) => NavigateToMonitorRow(target);
                lblNum.DoubleClick     += (s, e) => NavigateToMonitorRow(target);
                lblAddr.DoubleClick    += (s, e) => NavigateToMonitorRow(target);
                chip.Name.DoubleClick  += (s, e) => NavigateToMonitorRow(target);
                chip.Value.DoubleClick += (s, e) => NavigateToMonitorRow(target);

                _statusChips.Add(chip);
                statusBoardFlow.Controls.Add(chip.Panel);
            }

            SyncStatusBoardHeight();
            statusBoardFlow.ResumeLayout();
        }

        /// <summary>チップとサマリーを差分更新する（1秒タイマーから呼ぶ）</summary>
        private void UpdateStatusBoard()
        {
            if (_statusChips.Count == 0) return;

            int ok = 0, down = 0, recovery = 0;

            foreach (var chip in _statusChips)
            {
                string status = chip.Item.ステータス ?? "";
                string value;
                Color color;

                if (status.StartsWith("Down", StringComparison.OrdinalIgnoreCase))
                {
                    down++;
                    value = status;
                    color = Color.MistyRose;
                }
                else if (status.StartsWith("復旧"))
                {
                    recovery++;
                    value = status;
                    color = Color.LightSkyBlue;
                }
                else if (status.Equals("OK", StringComparison.OrdinalIgnoreCase))
                {
                    ok++;
                    value = $"{chip.Item.時間ms}ms";
                    color = Color.LightGreen;
                }
                else
                {
                    value = "";
                    color = SystemColors.Control;
                }

                // 変わったときだけ書き換える（200拠点でのチラつき防止）
                if (chip.LastValue != value)
                {
                    chip.Value.Text = value;
                    chip.LastValue = value;
                }
                if (chip.LastColor != color)
                {
                    chip.Panel.BackColor = color;
                    chip.LastColor = color;
                }
            }

            UpdateSummaryLabel(lblBoardTotal, $"監視中: {_statusChips.Count}");
            UpdateSummaryLabel(lblBoardOk, $"OK: {ok}");
            UpdateSummaryLabel(lblBoardDown, $"Down: {down}");
            UpdateSummaryLabel(lblBoardRecovery, $"復旧: {recovery}");
        }

        private void UpdateSummaryLabel(Label label, string text)
        {
            if (label.Text != text) label.Text = text;
        }

        /// <summary>ボードをクリアする（ClearView から呼ぶ）</summary>
        private void ClearStatusBoard()
        {
            statusBoardFlow.Controls.Clear();
            _statusChips.Clear();
            lblBoardTotal.Text = "監視中: 0";
            lblBoardOk.Text = "OK: 0";
            lblBoardDown.Text = "Down: 0";
            lblBoardRecovery.Text = "復旧: 0";
        }

        /// <summary>監視統計タブへ切り替え、該当アイテムの行を選択・表示する</summary>
        private void NavigateToMonitorRow(PingMonitorItem item)
        {
            tabControl.SelectedTab = statsPage;

            // タブ切り替え直後はグリッドの表示情報が未確定のため、
            // レイアウト完了後（BeginInvoke）に行選択・スクロールを行う
            BeginInvoke(new Action(() =>
            {
                dgvMonitor.Focus();
                foreach (DataGridViewRow row in dgvMonitor.Rows)
                {
                    if (ReferenceEquals(row.DataBoundItem, item))
                    {
                        dgvMonitor.ClearSelection();
                        row.Selected = true;
                        var firstVisible = dgvMonitor.Columns.GetFirstColumn(DataGridViewElementStates.Visible);
                        if (firstVisible != null)
                        {
                            try { dgvMonitor.CurrentCell = row.Cells[firstVisible.Index]; } catch { }
                        }
                        try
                        {
                            // 既に画面内に表示されている行はスクロール不要
                            if (!row.Displayed && row.Index >= 0 && row.Index < dgvMonitor.RowCount)
                                dgvMonitor.FirstDisplayedScrollingRowIndex = row.Index;
                        }
                        catch { }
                        break;
                    }
                }
            }));
        }
    }
}
