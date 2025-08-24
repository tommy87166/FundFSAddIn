using System;
using System.Windows.Forms;

namespace FundFSAddIn
{
    public partial class ProgressDialog : Form
    {
        public ProgressDialog(string title)
        {
            InitializeComponent();
            this.Text = title;
        }

        public void UpdateProgress(int current, int total, string additionalInfo = null)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<int, int, string>(UpdateProgress), current, total, additionalInfo);
                return;
            }

            // 計算百分比
            int percentage = total > 0 ? Math.Min((int)(((double)current / total) * 100), 100) : 0;
            
            // 更新進度條
            progressBar.Value = percentage;
            
            // 更新狀態文字
            string status = $"處理中: {current}/{total} ({percentage}%)";
            if (!string.IsNullOrEmpty(additionalInfo))
                status += $" - {additionalInfo}";
            
            lblStatus.Text = status;
        }

        public void CompleteProgress(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(CompleteProgress), message);
                return;
            }
            progressBar.Value = 100;
            lblStatus.Text = message;
            btnClose.Enabled = true;
            btnClose.Visible = true;
            btnClose.Text = "好";
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
