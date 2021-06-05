using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace chamcong
{
    public partial class Form3 : Form
    {
        DateTime dateTime;
        String savePath;
        public Form3(DateTime a, String b)
        {
            dateTime = a;
            savePath = b;
            InitializeComponent();
        }
        private void readName(StreamReader a)
        {
            string line = a.ReadLine();
            int j = 0;
            while (line != null)
            {
                int i = 0;
                customDataGridView1.Rows.Add();
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

        private void readDate(StreamReader a)
        {
            string line = a.ReadLine();
            while (line != null)
            {
                string[] k = line.Split(',');
                for (int i = 0; i < customDataGridView1.Rows.Count; i++)
                {
                    if (customDataGridView1[0, i].Value != null)
                    {
                        if (customDataGridView1[0, i].Value.ToString() == k[0])
                        {
                            customDataGridView1["d", i].Value = k[DateTime.DaysInMonth(dateTime.Year,dateTime.Month) + 2];
                            customDataGridView1["e", i].Value = k[DateTime.DaysInMonth(dateTime.Year, dateTime.Month) + 4];
                            customDataGridView1["h", i].Value = k[DateTime.DaysInMonth(dateTime.Year, dateTime.Month) + 1];
                            customDataGridView1["j", i].Value = k[DateTime.DaysInMonth(dateTime.Year, dateTime.Month) + 3];
                        }
                    }
                }
                line = a.ReadLine();
            }
        }
        private void readAdditional(StreamReader a)
        {
            string line = a.ReadLine();
            while (line != null)
            {
                string[] k = line.Split(',');
                for (int i = 0; i < customDataGridView1.Rows.Count; i++)
                {
                    if (customDataGridView1[0, i].Value != null)
                    {
                        if (customDataGridView1[0, i].Value.ToString() == k[0])
                        {
                            customDataGridView1[5, i].Value = k[5];
                            customDataGridView1[6, i].Value = k[6];
                            customDataGridView1[8, i].Value = k[8];
                        }
                    }
                }
                line = a.ReadLine();
            }
        }
        public void fillDefault()
        {
            for (int i = 0; i < customDataGridView1.RowCount; i++)
            {
                customDataGridView1[5, i].Value = "0";
                customDataGridView1[6, i].Value = "0";
                if (customDataGridView1[4, i].Value.ToString() != "0")
                {
                    customDataGridView1[8, i].Value = "Nghỉ phép";
                }
            }
        }
        public void saveToEntry(System.Data.DataTable dtDataTable, Stream strFilePath)
        {
            try
            {
                StreamWriter sw = new StreamWriter(strFilePath, Encoding.UTF8);

                foreach (DataRow dr in dtDataTable.Rows)
                {
                    for (int i = 0; i < dtDataTable.Columns.Count; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            string value = dr[i].ToString();
                            if (value.Contains(','))
                            {
                                value = String.Format("\"{0}\"", value);
                                sw.Write(value);
                            }
                            else
                            {
                                sw.Write(dr[i].ToString());
                            }
                        }
                        if (i < dtDataTable.Columns.Count - 1)
                        {
                            sw.Write(",");
                        }
                    }
                    sw.Write(sw.NewLine);
                }
                sw.Close();
            }
            catch
            {
                const string message = "File này đang được mở ở chương trình khác nên không thể lưu đè lên được, xin hãy đóng nó trước khi lưu";
                const string caption = "Lỗi";
                var result = MessageBox.Show(message, caption,
                                             MessageBoxButtons.OK,
                                             MessageBoxIcon.Question);
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
        public string datetostring(DateTime inp)
        {
            return "Thang" + dateTime.Month.ToString() + "Nam" + dateTime.Year.ToString();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
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

            customDataGridView1.Columns["a"].ReadOnly = true;
            customDataGridView1.Columns["b"].ReadOnly = true;
            customDataGridView1.Columns["c"].ReadOnly = true;
            customDataGridView1.Columns["d"].ReadOnly = true;
            customDataGridView1.Columns["h"].ReadOnly = true;
            customDataGridView1.Columns["e"].ReadOnly = true;
            customDataGridView1.Columns["j"].ReadOnly = true;
            using (FileStream zipToOpen = new FileStream(savePath, FileMode.Open))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry nameEntry = archive.GetEntry("DanhSach.csv");
                    if (nameEntry == null)
                    {
                        nameEntry = archive.CreateEntry("DanhSach.csv");
                    }
                    readName(new StreamReader(nameEntry.Open()));

                    ZipArchiveEntry dateEntry = archive.GetEntry(datetostring(dateTime) + ".csv");
                    if (dateEntry == null)
                    {
                        dateEntry = archive.CreateEntry("TongHop" + datetostring(dateTime) + ".csv");
                    }
                    readDate(new StreamReader(dateEntry.Open()));


                    ZipArchiveEntry readmeEntry = archive.GetEntry("TongHop" + datetostring(dateTime) + ".csv");
                    if (readmeEntry == null)
                    {
                        readmeEntry = archive.CreateEntry("TongHop" + datetostring(dateTime) + ".csv");
                        fillDefault();
                    }
                    else
                    {
                        readAdditional(new StreamReader(readmeEntry.Open()));
                    }
                }
            }
            for (int i = 0; i < customDataGridView1.RowCount; i++)
            {
                if (customDataGridView1[4, i].Value.ToString() != "0" && customDataGridView1[8, i].Value.ToString() == "")
                {
                    customDataGridView1[8, i].Value = "Nghỉ phép";
                }
            }
        }

        private void customDataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                foreach (DataGridViewCell a in customDataGridView1.SelectedCells)
                {
                    customDataGridView1[a.ColumnIndex, a.RowIndex].Value = "";
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (FileStream zipToOpen = new FileStream(savePath, FileMode.Open))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    if (archive.GetEntry("TongHop" + datetostring(dateTime) + ".csv") != null)
                    {
                        archive.GetEntry("TongHop" + datetostring(dateTime) + ".csv").Delete();
                    }
                    ZipArchiveEntry readmeEntry = archive.CreateEntry("TongHop" + datetostring(dateTime) + ".csv");
                    saveToEntry(GetDataGridViewAsDataTable(customDataGridView1), readmeEntry.Open());
                }
            }
            this.Close();
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            using (FileStream zipToOpen = new FileStream(savePath, FileMode.Open))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    if (archive.GetEntry("TongHop" + datetostring(dateTime) + ".csv") != null)
                    {
                        archive.GetEntry("TongHop" + datetostring(dateTime) + ".csv").Delete();
                    }
                    ZipArchiveEntry readmeEntry = archive.CreateEntry("TongHop" + datetostring(dateTime) + ".csv");
                    saveToEntry(GetDataGridViewAsDataTable(customDataGridView1), readmeEntry.Open());
                }
            }
        }

        
        string exeFile = "";
        private void excelAppShow(string a)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = true;
            xlWorkBook = xlApp.Workbooks.Open(a);
            xlApp.Visible = true;
        }
        private void readExcel(string sFile)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(sFile);
            xlWorkSheet = xlWorkBook.Worksheets["MAU 02"];
            xlWorkSheet.Cells[4, 1].value = "Tháng " + dateTime.Month.ToString() +" Năm " + dateTime.Year.ToString();
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
                xlWorkSheet.Cells[8, 2].value = customDataGridView1[0, iRow].Value;
                xlWorkSheet.Cells[8, 3].value = customDataGridView1[1, iRow].Value;
                xlWorkSheet.Cells[8, 4].value = customDataGridView1[2, iRow].Value;
                xlWorkSheet.Cells[8, 5].value = customDataGridView1[3, iRow].Value;
                xlWorkSheet.Cells[8, 6].value = customDataGridView1[4, iRow].Value;
                xlWorkSheet.Cells[8, 7].value = customDataGridView1[5, iRow].Value;
                xlWorkSheet.Cells[8, 8].value = customDataGridView1[6, iRow].Value;
                xlWorkSheet.Cells[8, 9].value = customDataGridView1[7, iRow].Value;
                xlWorkSheet.Cells[8, 10].value = customDataGridView1[8, iRow].Value;
                xlWorkSheet.Cells[8, 11].value = customDataGridView1[9, iRow].Value;
            }

            string saveTo = Path.Combine(exeFile, "TongHop.xls");
            xlWorkBook.SaveAs(saveTo, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges
                    , Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
            xlWorkBook.Close();
            xlApp.Quit();
            excelAppShow(saveTo);
        }

        private void menuItem1_Click(object sender, EventArgs e)
        {
            exeFile = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            string fullPath = Path.Combine(exeFile, "a.xls");
            label1.Text = fullPath;
            readExcel(fullPath);
        }
    }
}
