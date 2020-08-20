using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelController
{
    public partial class Form1 : Form
    {
        private const int ColorByteRed      = 0;
        private const int ColorByteGreen    = 8;
        private const int ColorByteBlue     = 16;

        public Form1()
        {   
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
				FileNameTextBox.Text = openFileDialog1.FileName;
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            string xlBookFilePass = FileNameTextBox.Text;

            ExcelAppAccessor _App = new ExcelAppAccessor();
            ExcelBookAccessor _Book = null;
            ExcelSheetAccessor _Sheet = null;
            Excel.Range _Range1 = null;
            Excel.Range _Range2 = null;
            Excel.Range _Range3 = null;

            try
            {
                _Book = _App.Open(xlBookFilePass);
                _Sheet = _Book.Open(1);

                _Range1 = _Sheet.GetRange("A1", "M1");
                _Range2 = _Sheet.GetRange("A2", "A23");
                _Range3 = _App.Union(_Range1, _Range2);

                _Range3.Interior.Color = ((0x00 << ColorByteBlue) | (0x00 << ColorByteGreen) | (0xFF << ColorByteRed));

                string label_txt = "";

                label_txt = _Range1.Row.ToString() + "," + _Range1.Column.ToString();

				ResultTextBox.Text = label_txt;

                /*xlSheet = _Book.GetSheet("チェックリスト") as Excel.Worksheet;
    
                object[,] cellData = new object[10, 5]; //受け取るデータの個数分（行、列セル数を指定）
                Excel.Range xlRange = xlSheet.Range["A4", "AI500"];
                cellData = xlRange.Value;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);

                string label_txt = "";
                //データを出力 要素数は1から格納されています
                for (int r = 1; r <= cellData.GetLength(0); r++)
                {
                    for (int c = 1; c <= cellData.GetLength(1); c++)
                    {
                        label_txt += string.Format(" {0} ", cellData[r, c]);
                    }

                    label_txt += "\n";
                }
                   
                label1.Text = label_txt;

                xlRange = xlSheet.Range["A1"];
                xlRange.Value = "Model : A";
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                */
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetType().FullName + "\n" + ex.Message);
            }
            finally
            {
                if (_Range1 != null)
                {
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(_Range1) > 0);
                }

                if (_Range2 != null)
                {
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(_Range2) > 0) ;
                }

                if (_Range3 != null)
                {
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(_Range3) > 0) ;
                }

                if (_Sheet != null) _Sheet.Dispose();
                if (_Book != null)  _Book.Dispose();
                if (_App != null)   _App.Dispose();

                //全てのEXCELのバックグラウンドプロセスを直接開放する
                foreach (Process _Process in Process.GetProcessesByName("EXCEL"))
                {
                    if (_Process.MainWindowTitle == "") _Process.Kill();
                }
            }
        }

		private void Button3_Click(object sender, EventArgs e)
		{

		}

		private void TextBox3_TextChanged(object sender, EventArgs e)
		{

		}

		private void OpenFileDialog1_FileOk(object sender, CancelEventArgs e)
		{

		}
	}
}
