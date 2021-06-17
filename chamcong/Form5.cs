using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace chamcong
{
    public partial class Form5 : Form
    {
        string exeFile;
        public Form5()
        {
            InitializeComponent();
            initTable();
        }
        private void initTable()
        {
            customDataGridView1.ReadOnly = true;
            customDataGridView1.AllowUserToAddRows = false;
            customDataGridView1.Columns.Add("a", "Họ và tên");
            customDataGridView1.Columns.Add("b", "Chức vụ");
            customDataGridView1.Columns.Add("c", "Đội công tác/ vị trí công tác");
            customDataGridView1.Columns.Add("d", "Số ngày làm việc");
            customDataGridView1.Columns.Add("e", "Số ngày nghỉ có lý do");
            customDataGridView1.Columns.Add("f", "Số lần vi phạm quy chế, quy định");
            customDataGridView1.Columns.Add("g", "Hình thức kỉ luật");
            customDataGridView1.Columns.Add("h", "Kết quả xếp loại");
            customDataGridView1.Columns.Add("i", "Lý do nghỉ");
            customDataGridView1.Columns.Add("j", "Số giờ làm thêm");
        }

        private void populateGrid(StreamReader a)
        {
            string line = a.ReadLine();
            int j = customDataGridView1.RowCount;
            while (line != null)
            {
                customDataGridView1.Rows.Add();
                int i = 0;
                for (int k = 0; k < line.Length; k++)
                {
                    if (line[k] != ',')
                    {
                        customDataGridView1[i, j].Value += line[k].ToString();
                    }
                    else
                    {
                        i++;
                    }
                }
                line = a.ReadLine();
                j++;
            }
        }

        private void menuItem1_Click(object sender, EventArgs e)
        {
            OpenFileDialog saveFileDialog1 = new OpenFileDialog();
            saveFileDialog1.Title = "Open";
            saveFileDialog1.Filter = "csv | *.csv";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                string a = saveFileDialog1.FileName;
                FileStream fs = File.Open(a, FileMode.Open);
                populateGrid(new StreamReader(fs));
                fs.Close();
            }
        }
        private System.Data.DataTable GetDataGridViewAsDataTable(DataGridView _DataGridView)
        {
            try
            {
                if (_DataGridView.ColumnCount == 0) return null;
                System.Data.DataTable dtSource = new System.Data.DataTable();
                //////create columns
                foreach (DataGridViewColumn col in _DataGridView.Columns)
                {
                    if (col.ValueType == null) dtSource.Columns.Add(col.Name, typeof(string));
                    else dtSource.Columns.Add(col.Name, col.ValueType);
                    dtSource.Columns[col.Name].Caption = col.HeaderText;
                }
                ///////insert row data
                int count = -1;
                foreach (DataGridViewRow row in _DataGridView.Rows)
                {
                    count++;
                    DataRow drNewRow = dtSource.NewRow();
                    foreach (DataColumn col in dtSource.Columns)
                    {
                        drNewRow[col.ColumnName] = row.Cells[col.ColumnName].Value;
                    }
                    dtSource.Rows.Add(drNewRow);
                }
                return dtSource;
            }
            catch
            {
                return null;
            }
        }
        static T[,] To2D<T>(T[][] source)
        {
            try
            {
                int FirstDim = source.Length;
                int SecondDim = source.GroupBy(row => row.Length).Single().Key; // throws InvalidOperationException if source is not rectangular

                var result = new T[FirstDim, SecondDim];
                for (int i = 0; i < FirstDim; ++i)
                    for (int j = 0; j < SecondDim; ++j)
                        result[i, j] = source[i][j];

                return result;
            }
            catch (InvalidOperationException)
            {
                throw new InvalidOperationException("The given jagged array is not rectangular.");
            }
        }
        private void readExcel(string sFile)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlApp.ScreenUpdating = false;
            xlWorkBook = xlApp.Workbooks.Open(sFile);
            xlWorkSheet = xlWorkBook.Worksheets["MAU 02"];
            xlWorkSheet.Cells[4, 1].value = "Tháng " + DateTime.Now.Month.ToString() + " Năm " + DateTime.Now.Year.ToString();
            for (int iRow = customDataGridView1.RowCount - 1; iRow >= 0; iRow--)
            {
                Range line = (Range)xlWorkSheet.Rows[8];
                line.Insert();
                line = (Range)xlWorkSheet.Range["A8", "K8"];
                line.Borders.LineStyle = 1;
                line.Borders.Weight = 2;
                line.Font.Bold = false;
                line.Font.Italic = false;
                line.Font.Size = 11.5;
                xlWorkSheet.Cells[8, 1].value = iRow + 1;
            }
            Excel.Range range = (Excel.Range)xlWorkSheet.Cells[8, 2];
            range = range.get_Resize(customDataGridView1.RowCount, customDataGridView1.ColumnCount);
            object[][] tableArray = GetDataGridViewAsDataTable(customDataGridView1).AsEnumerable().Select(x => x.ItemArray).ToArray();
            object[,] ttableArray = To2D(tableArray);
            range.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, ttableArray);
            xlApp.ScreenUpdating = true;

            string saveTo = Path.Combine(exeFile, "TongHopCacDoi.xls");
            xlWorkBook.SaveAs(saveTo, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges
                    , Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
            xlWorkBook.Close();
            xlApp.Quit();
            excelAppShow(saveTo);
        }
        private void excelAppShow(string a)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = true;
            xlWorkBook = xlApp.Workbooks.Open(a);
            xlApp.Visible = true;
        }
        private void menuItem2_Click(object sender, EventArgs e)
        {
            exeFile = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            string fullPath = Path.Combine(exeFile, "b.xls");
            readExcel(fullPath);
        }
    }
}
