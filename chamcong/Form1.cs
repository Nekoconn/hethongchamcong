using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace chamcong
{
    public partial class Form1 : Form
    {
        DateTime dateTime;
        String savePath;
        
        public Form1()
        {
            InitializeComponent();
            dateTime = DateTime.Now;
            label1.Text = "Tháng " + dateTime.Month.ToString() + " Năm " + dateTime.Year.ToString();
            savePath = Properties.Settings.Default.savePathSetting;
            initDataGrid();
        }
        private string getWeekday(int a)
        {
            DateTime k = new DateTime(dateTime.Year, dateTime.Month, a);
            if (k.DayOfWeek != 0)
            {
                return "T" + ((int)k.DayOfWeek + 1).ToString();
            }
            else
            {
                return "CN";
            }
        }
        private void initDataGrid()
        {
            ((ISupportInitialize)dataGridView1).BeginInit();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("a", "Họ và tên");
            dataGridView1.Columns[0].ReadOnly = true;
            for (int i = 1; i <= DateTime.DaysInMonth(dateTime.Year, dateTime.Month); i++)
            {
                dataGridView1.Columns.Add(i.ToString(),  getWeekday(i) + "\n" +  i.ToString());
                dataGridView1.Columns[i.ToString()].Width = 30;
                if (getWeekday(i) == "CN" || getWeekday(i) == "T7")
                {
                    dataGridView1.Columns[i.ToString()].DefaultCellStyle.BackColor = Color.Gray;
                }
            }
            dataGridView1.Columns.Add("b", "ĐGPL");
            dataGridView1.Columns.Add("c", "Tổng ngày làm việc");
            dataGridView1.Columns.Add("d", "Số giờ làm thêm");
            dataGridView1.Columns.Add("e", "Nghỉ phép");
            dataGridView1.Columns.Add("f", "Ốm");
            dataGridView1.Columns["b"].Width = 40;
            dataGridView1.Columns["c"].Width = 80;
            dataGridView1.Columns["d"].Width = 80;
            dataGridView1.Columns["e"].Width = 80;
            dataGridView1.Columns["f"].Width = 80;
            for (int i = DateTime.DaysInMonth(dateTime.Year, dateTime.Month) + 2; i <= DateTime.DaysInMonth(dateTime.Year, dateTime.Month) + 5; i++)
            {
                dataGridView1.Columns[i].ReadOnly = true;
            }
            try
            {
                using (FileStream zipToOpen = new FileStream(savePath, FileMode.Open))
                {
                    using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                    {
                        if (archive.GetEntry("DanhSach.csv") != null)
                        {
                            ZipArchiveEntry readmeEntry = archive.GetEntry("DanhSach.csv");
                            populateGridName(new StreamReader(readmeEntry.Open()));
                            readmeEntry = archive.GetEntry(datetostring(dateTime) + ".csv");
                            if (readmeEntry != null)
                            {
                                populateWorkDay(new StreamReader(readmeEntry.Open()));
                            }
                            else
                            {
                                fillDefault();
                            }
                        }
                    }
                }
            }
            catch
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "New";
                saveFileDialog1.Filter = "Zip | *.zip";
                saveFileDialog1.ShowDialog();
                if (saveFileDialog1.FileName != "")
                {
                    using (FileStream zipToOpen = new FileStream(saveFileDialog1.FileName, FileMode.Create))
                    {
                        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                        {
                            ZipArchiveEntry readmeEntry = archive.CreateEntry("Readme.txt");
                            using (StreamWriter writer = new StreamWriter(readmeEntry.Open()))
                            {
                                writer.WriteLine("Information about this package.");
                                writer.WriteLine("========================");
                            }
                        }
                    }
                    savePath = saveFileDialog1.FileName;
                }
            }
            ((ISupportInitialize)dataGridView1).EndInit();
        }
        private void populateGridName(StreamReader a)
        {
            string line = a.ReadLine();
            int j = 0;
            while (line != null)
            {
                dataGridView1.Rows.Add();
                int i = 0;
                for (int k = 0; k < line.Length; k++)
                {
                    if (i > 0)
                    {
                        break;
                    }
                    if (line[k] != ',')
                    {
                        dataGridView1[i, j].Value += line[k].ToString();
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
        private void populateWorkDay(StreamReader a)
        {
            string line = a.ReadLine();
            while (line != null)
            {
                string[] k = line.Split(',');
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1[0, i].Value != null)
                    {
                        if (dataGridView1[0, i].Value.ToString() == k[0])
                        {
                            for (int j = 1; j < dataGridView1.ColumnCount; j++)
                            {
                                dataGridView1[j, i].Value = k[j];
                            }
                        }
                    }
                }
                line = a.ReadLine();
            }
        }
        private void saveCurrent()
        {
            using (FileStream zipToOpen = new FileStream(savePath, FileMode.Open))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    if (archive.GetEntry(datetostring(dateTime) + ".csv") != null)
                    {
                        archive.GetEntry(datetostring(dateTime) + ".csv").Delete();
                    }
                    ZipArchiveEntry readmeEntry = archive.CreateEntry(datetostring(dateTime) + ".csv");
                    saveToEntry(GetDataGridViewAsDataTable(dataGridView1), readmeEntry.Open());
                }
            }
        }

        private void fillDefault()
        {
            for (int i = 1; i <= DateTime.DaysInMonth(dateTime.Year, dateTime.Month); i++)
            {
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    if (dataGridView1[i.ToString(), j].Value == "" || dataGridView1[i.ToString(), j].Value == null)
                    {
                        if (getWeekday(i) == "CN" || getWeekday(i) == "T7")
                        {
                            dataGridView1[i.ToString(), j].Value = "";
                        }
                        else
                        {
                            dataGridView1[i.ToString(), j].Value = "X";
                        }
                    }
                }
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1["b", i].Value = "A";
            }
        }

        private DataTable GetDataGridViewAsDataTable(DataGridView _DataGridView)
        {
            try
            {
                if (_DataGridView.ColumnCount == 0) return null;
                DataTable dtSource = new DataTable();
                //////create columns
                foreach (DataGridViewColumn col in _DataGridView.Columns)
                {
                    if (col.ValueType == null) dtSource.Columns.Add(col.Name, typeof(string));
                    else dtSource.Columns.Add(col.Name, col.ValueType);
                    dtSource.Columns[col.Name].Caption = col.HeaderText;
                }
                ///////insert row data
                foreach (DataGridViewRow row in _DataGridView.Rows)
                {
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
        public void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            try
            {
                StreamWriter sw = new StreamWriter(
                    new FileStream(strFilePath, FileMode.Create, FileAccess.ReadWrite),
                    Encoding.UTF8
                );
                
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

        public void saveToEntry(DataTable dtDataTable, Stream strFilePath)
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

        public string datetostring(DateTime inp)
        {
            return "Thang" + dateTime.Month.ToString() + "Nam" + dateTime.Year.ToString();
        }

        //Save current table as csv
        private void menuItem3_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Save as";
            saveFileDialog1.Filter = "CSV | *.csv";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                DataTable dat = GetDataGridViewAsDataTable(dataGridView1);
                ToCSV(dat, saveFileDialog1.FileName);
            }
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                this.dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }
        }
        private void recalculate(int rrow)
        {
            int tong = 0;
            int lamthem = 0;
            int om = 0;
            int nghiphep = 0;
            for (int i = 0; i < dataGridView1.ColumnCount - 1; i++)
            {
                if (dataGridView1[i, rrow].Value != null)
                {
                    if (dataGridView1[i, rrow].Value.ToString() == "X" 
                        || dataGridView1[i, rrow].Value.ToString() == "H" 
                        || dataGridView1[i, rrow].Value.ToString() == "CT"
                        || dataGridView1[i, rrow].Value.ToString() == "Ts")
                    {
                        tong++;
                    }
                    if (dataGridView1[i, rrow].Value.ToString() == "O")
                    {
                        om++;
                    }
                    if (dataGridView1[i, rrow].Value.ToString() == "P")
                    {
                        nghiphep++;
                    }
                    if (dataGridView1[i, rrow].Value.ToString() == "NG")
                    {
                        lamthem+= 8;
                    }
                }

            }
            dataGridView1.Rows[rrow].Cells["c"].Value = tong;
            dataGridView1.Rows[rrow].Cells["d"].Value = lamthem;
            dataGridView1.Rows[rrow].Cells["e"].Value = nghiphep;
            dataGridView1.Rows[rrow].Cells["f"].Value = om;
        }
        //Next Month
        private void button3_Click(object sender, EventArgs e)
        {
            saveCurrent();
            dateTime = dateTime.AddMonths(1);
            label1.Text = "Tháng " + dateTime.Month.ToString() + " Năm " + dateTime.Year.ToString();
            initDataGrid();
        }
        //Previous Month
        private void button4_Click(object sender, EventArgs e)
        {
            saveCurrent();
            dateTime = dateTime.AddMonths(-1);
            label1.Text = "Tháng " + dateTime.Month.ToString() + " Năm " + dateTime.Year.ToString();

            initDataGrid();
        }

        //New file
        private void menuItem4_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "New";
            saveFileDialog1.Filter = "Zip | *.zip";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                using (FileStream zipToOpen = new FileStream(saveFileDialog1.FileName, FileMode.Create))
                {
                    using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                    {
                        ZipArchiveEntry readmeEntry = archive.CreateEntry("Readme.txt");
                        using (StreamWriter writer = new StreamWriter(readmeEntry.Open()))
                        {
                            writer.WriteLine("Information about this package.");
                            writer.WriteLine("========================");
                        }
                    }
                }
                savePath = saveFileDialog1.FileName;
                Properties.Settings.Default.savePathSetting = savePath;
                Properties.Settings.Default.Save();
            }
            initDataGrid();
        }

        //Save file
        private void menuItem6_Click(object sender, EventArgs e)
        {
            saveCurrent();
        }
        //Close form
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.savePathSetting = savePath;
            Properties.Settings.Default.Save();
            dataGridView1.Visible = false;
            DialogResult r = MessageBox.Show(this, "Có lưu không?", "Lưu",MessageBoxButtons.YesNoCancel);
            if (r == DialogResult.Yes)
            {
                saveCurrent();
            }
            else
            {
                if (r == DialogResult.No)
                {
                }
                else
                {
                    e.Cancel = true;
                    dataGridView1.Visible = true;
                }
            }
        }
        //Open File
        private void menuItem5_Click(object sender, EventArgs e)
        {
            OpenFileDialog saveFileDialog1 = new OpenFileDialog();
            saveFileDialog1.Title = "Open";
            saveFileDialog1.Filter = "Zip | *.zip";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                savePath = saveFileDialog1.FileName;
                Properties.Settings.Default.savePathSetting = savePath;
                Properties.Settings.Default.Save();
            }
            initDataGrid();
        }

        private void dataGridView1_Paint(object sender, PaintEventArgs e)
        {
            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                this.dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }
        }

        private void menuItem7_Click(object sender, EventArgs e)
        {
            Form2 a = new Form2(savePath);
            a.ShowDialog();
            initDataGrid();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            recalculate(e.RowIndex);
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                foreach (DataGridViewCell a in dataGridView1.SelectedCells)
                {
                    dataGridView1[a.ColumnIndex, a.RowIndex].Value = "";
                }
            }
        }

        private void menuItem8_Click(object sender, EventArgs e)
        {
            saveCurrent();
            Form3 a = new Form3(dateTime, savePath);
            a.Show();
        }



        private void menuItem9_Click(object sender, EventArgs e)
        {
            
        }
    }
}
