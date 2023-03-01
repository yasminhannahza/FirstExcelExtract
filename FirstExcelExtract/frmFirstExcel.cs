using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FirstExcelExtract
{
    public partial class frmFirstExcel : Form
    {
        public frmFirstExcel()
        {
            InitializeComponent();
        }

        public async Task UpdateConsole(string message, Exception ex = null)
        {
            string outputText = $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}:\r\n" +
                $"{message}\r\n\r\n";

            if (ex != null)
            {
                outputText += $"{ex}\r\n\r\n";
            }

            outputText += "\r\n";

            this.Invoke(new Action(() =>
            {
                txtConsole.AppendText(outputText);
            }));
        }

        public async Task UpdateProgressBar(int value, int maxValue, bool visible = false)
        {
            this.Invoke(new Action(() =>
            {
                pbProgress.Visible = visible;
                pbProgress.Maximum = maxValue;
                pbProgress.Value = value;
            }));
        }
        private void btnExecute_Click(object sender, EventArgs e)
        {
            Task.Run(async () =>
            {
                try
                {
                    await ExcelLogic();
                }
                catch (Exception ex)
                {
                    await UpdateConsole("ExcelLogic Routine Error", ex);
                }
            });
        }
    }
}
