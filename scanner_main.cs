using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

class Program
{
    static readonly string[] BUILTIN_REASONS = {
        "版本問題", "治具跳錯誤訊息", "LED4閃", "自動校準未過",
        "人員偵測NG", "機台測試NG", "零件組裝問題", "未過電",
        "序號問題", "速度未過", "揚升未過", "無法辨別\"DE\"",
        "過電不動", "睡眠未過", "其他"
    };

    static List<string> customReasons = new List<string>();
    static string customFile = "";

    static string[] AllReasons()
    {
        var list = new List<string>(BUILTIN_REASONS);
        list.AddRange(customReasons);
        return list.ToArray();
    }

    static void LoadCustomReasons()
    {
        if (File.Exists(customFile))
            customReasons = File.ReadAllLines(customFile, Encoding.UTF8)
                .Where(l => l.Trim() != "").ToList();
    }

    static void SaveCustomReasons()
    {
        File.WriteAllLines(customFile, customReasons, Encoding.UTF8);
    }

    static dynamic excel = null;
    static dynamic wb = null;
    static dynamic ws = null;
    static readonly Dictionary<int, string> known = new Dictionary<int, string>();
    static bool busy = false;
    const int DATA_START = 3;
    const int DATA_END = 300;

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        customFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "custom_reasons.txt");
        LoadCustomReasons();

        try { excel = Marshal.GetActiveObject("Excel.Application"); }
        catch
        {
            MessageBox.Show("請先開啟 Excel 不良明細檔案，再執行此程式！",
                "找不到 Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        wb = null;
        int wbCount = excel.Workbooks.Count;
        for (int i = 1; i <= wbCount; i++)
        {
            dynamic book = excel.Workbooks.Item[i];
            string name = (string)book.Name;
            if (name.Contains("不良明細") || name.Contains("TM23PL"))
            { wb = book; break; }
        }
        if (wb == null && wbCount > 0)
            wb = excel.Workbooks.Item[1];

        if (wb == null)
        {
            MessageBox.Show("找不到任何 Excel 活頁簿！", "錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try { ws = wb.Sheets.Item["RAW DATA"]; }
        catch
        {
            MessageBox.Show("找不到工作表 'RAW DATA'！", "錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            object[,] init = (object[,])ws.Range["B" + DATA_START + ":C" + DATA_END].Value2;
            if (init != null)
            {
                for (int i = 1; i <= DATA_END - DATA_START + 1; i++)
                {
                    object s = init[i, 1], c = init[i, 2];
                    int r = DATA_START + i - 1;
                    if (s != null && s.ToString() != "" && c != null && c.ToString() != "")
                        known[r] = s.ToString();
                }
            }
        }
        catch { }

        var mainForm = new Form
        {
            Text = "掃描監聽器",
            ClientSize = new Size(280, 80),
            StartPosition = FormStartPosition.Manual,
            Location = new Point(5, 5),
            FormBorderStyle = FormBorderStyle.FixedToolWindow,
            TopMost = true,
            BackColor = Color.FromArgb(235, 245, 255)
        };

        var lblStatus = new Label
        {
            Font = new Font("Microsoft JhengHei UI", 9),
            Location = new Point(10, 8),
            Size = new Size(260, 62),
            Text = "✓ 監聽中...\n掃描序號後自動彈出選擇視窗"
        };
        mainForm.Controls.Add(lblStatus);

        var timer = new Timer { Interval = 250 };
        timer.Tick += (sender, e) =>
        {
            if (busy) return;
            try
            {
                object[,] vals = (object[,])ws.Range["B" + DATA_START + ":C" + DATA_END].Value2;
                if (vals == null) return;
                for (int i = 1; i <= DATA_END - DATA_START + 1; i++)
                {
                    object serialObj = vals[i, 1];
                    object reasonObj = vals[i, 2];
                    if (serialObj == null || serialObj.ToString() == "") continue;
                    string serial = serialObj.ToString().Trim();
                    int r = DATA_START + i - 1;
                    bool isNew = !known.ContainsKey(r) || known[r] != serial;
                    bool noReason = reasonObj == null || reasonObj.ToString() == "";
                    if (isNew && noReason)
                    {
                        known[r] = serial;
                        timer.Stop();
                        lblStatus.Text = "偵測到序號: " + serial + "\n等待選擇原因...";
                        string sel = ShowSelection(serial);
                        if (sel != null)
                        {
                            ws.Cells.Item[r, 3].Value2 = sel;
                            try { ws.Cells.Item[r + 1, 2].Select(); } catch { }
                            try { SetForegroundWindow(new IntPtr(Convert.ToInt64(excel.Hwnd))); } catch { }
                            lblStatus.Text = "✓ Row " + r + " 已填入: " + sel + "\n監聽中...";
                        }
                        else
                            lblStatus.Text = "⚠ Row " + r + " 取消，未填入\n監聽中...";
                        timer.Start();
                        break;
                    }
                }
            }
            catch { }
        };

        mainForm.FormClosing += (s, e) => timer.Stop();
        timer.Start();
        Application.Run(mainForm);
    }

    static string ShowSelection(string serialNum)
    {
        busy = true;
        string result = null;

        while (true)
        {
            bool doAddNew = false;
            string[] allReasons = AllReasons();

            const int bW = 128, bH = 50, sX = 10, sY = 46, gap = 4, cols = 3;
            int rows = (allReasons.Length + cols - 1) / cols;
            int formWidth = sX + cols * bW + (cols - 1) * gap + sX;   // 412
            int bottomY = sY + rows * (bH + gap) + 10;
            int formHeight = bottomY + 36 + 10;

            string[] keys = { "1","2","3","4","5","6","7","8","9","A","B","C","D","E","F" };

            var f = new Form
            {
                Text = "選擇產線不良原因",
                ClientSize = new Size(formWidth, formHeight),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true,
                BackColor = Color.FromArgb(245, 248, 255),
                KeyPreview = true
            };

            f.Controls.Add(new Label
            {
                Text = "序號：" + serialNum,
                Font = new Font("Microsoft JhengHei UI", 13, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 70, 140),
                Location = new Point(sX, 8),
                Size = new Size(formWidth - 2 * sX, 30)
            });

            for (int i = 0; i < allReasons.Length; i++)
            {
                string reason = allReasons[i];
                int col = i % cols, row = i / cols;
                string prefix = (i < keys.Length) ? "[" + keys[i] + "] " : "    ";
                var btn = new Button
                {
                    Text = prefix + reason,
                    Font = new Font("Microsoft JhengHei UI", 11),
                    Location = new Point(sX + col * (bW + gap), sY + row * (bH + gap)),
                    Size = new Size(bW, bH),
                    BackColor = Color.FromArgb(210, 228, 255),
                    FlatStyle = FlatStyle.Flat,
                    Tag = reason
                };
                btn.FlatAppearance.BorderColor = Color.FromArgb(100, 150, 220);
                btn.Click += (s, ev) => { result = ((Button)s).Tag.ToString(); f.Close(); };
                btn.MouseEnter += (s, ev) => ((Button)s).BackColor = Color.FromArgb(150, 190, 255);
                btn.MouseLeave += (s, ev) => ((Button)s).BackColor = Color.FromArgb(210, 228, 255);
                f.Controls.Add(btn);
            }

            var btnCancel = new Button
            {
                Text = "取消 (不填入)",
                Font = new Font("Microsoft JhengHei UI", 10),
                Location = new Point(sX, bottomY),
                Size = new Size(190, 36),
                BackColor = Color.LightGray,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.Click += (s, ev) => f.Close();
            f.Controls.Add(btnCancel);

            var btnAdd = new Button
            {
                Text = "＋ 新增選項",
                Font = new Font("Microsoft JhengHei UI", 10),
                Location = new Point(sX + 198, bottomY),
                Size = new Size(formWidth - sX - 198 - sX, 36),
                BackColor = Color.FromArgb(200, 240, 210),
                FlatStyle = FlatStyle.Flat
            };
            btnAdd.FlatAppearance.BorderColor = Color.FromArgb(100, 180, 120);
            btnAdd.Click += (s, ev) => { doAddNew = true; f.Close(); };
            f.Controls.Add(btnAdd);

            f.KeyDown += (s, e) =>
            {
                int idx = -1;
                if (e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D9)
                    idx = (int)e.KeyCode - (int)Keys.D1;
                else if (e.KeyCode >= Keys.NumPad1 && e.KeyCode <= Keys.NumPad9)
                    idx = (int)e.KeyCode - (int)Keys.NumPad1;
                else if (e.KeyCode >= Keys.A && e.KeyCode <= Keys.F)
                    idx = 9 + (int)e.KeyCode - (int)Keys.A;
                else if (e.KeyCode == Keys.Escape)
                { f.Close(); return; }
                if (idx >= 0 && idx < allReasons.Length)
                { result = allReasons[idx]; f.Close(); }
            };

            f.Shown += (s, ev) => f.Activate();
            f.ShowDialog();

            if (!doAddNew) break;  // 選了原因或取消 → 離開迴圈

            // 使用者點了「新增選項」
            string newReason = ShowInputDialog();
            if (newReason != null && newReason.Trim() != "")
            {
                customReasons.Add(newReason.Trim());
                SaveCustomReasons();
            }
            // 重新顯示選擇視窗（含新增的選項）
        }

        busy = false;
        return result;
    }

    static string ShowInputDialog()
    {
        string result = null;
        var f = new Form
        {
            Text = "新增不良原因",
            ClientSize = new Size(340, 122),
            StartPosition = FormStartPosition.CenterScreen,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            TopMost = true,
            BackColor = Color.FromArgb(250, 252, 255)
        };

        f.Controls.Add(new Label
        {
            Text = "輸入新的不良原因：",
            Font = new Font("Microsoft JhengHei UI", 11),
            Location = new Point(10, 12),
            Size = new Size(320, 26)
        });

        var txt = new TextBox
        {
            Location = new Point(10, 44),
            Size = new Size(320, 28),
            Font = new Font("Microsoft JhengHei UI", 12)
        };
        f.Controls.Add(txt);

        var btnOK = new Button
        {
            Text = "確定新增",
            Font = new Font("Microsoft JhengHei UI", 10),
            Location = new Point(10, 82),
            Size = new Size(150, 30),
            BackColor = Color.FromArgb(200, 240, 210),
            FlatStyle = FlatStyle.Flat
        };
        var btnCancel = new Button
        {
            Text = "取消",
            Font = new Font("Microsoft JhengHei UI", 10),
            Location = new Point(170, 82),
            Size = new Size(150, 30),
            BackColor = Color.LightGray,
            FlatStyle = FlatStyle.Flat
        };

        btnOK.Click += (s, e) => { result = txt.Text.Trim(); f.Close(); };
        btnCancel.Click += (s, e) => f.Close();
        f.Controls.Add(btnOK);
        f.Controls.Add(btnCancel);
        f.AcceptButton = btnOK;
        f.CancelButton = btnCancel;
        f.Shown += (s, e) => { f.Activate(); txt.Focus(); };
        f.ShowDialog();
        return result;
    }
}
