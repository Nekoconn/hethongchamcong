using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace chamcong
{
    public partial class Form2 : Form
    {
        string savePath;
        public Form2(string a)
        {
            InitializeComponent();
            savePath = a;
        }

        private void readAndUpdate(StreamReader a)
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
        private void readAndUpdateTab(StreamReader a)
        {
            string line = a.ReadLine();
            int j = 0;
            while (line != null)
            {
                int i = 0;
                customDataGridView1.Rows.Add();
                for (int k = 0; k < line.Length; k++)
                {
                    if (line[k] != '\t')
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
        private void Form2_Load(object sender, EventArgs e)
        {
            customDataGridView1.Columns.Add("a", "Họ và tên");
            customDataGridView1.Columns.Add("b", "Chức vụ");
            customDataGridView1.Columns.Add("c", "Đội công tác/ vị trí công tác");

            using (FileStream zipToOpen = new FileStream(savePath, FileMode.Open))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry = archive.GetEntry("DanhSach.csv");
                    if (readmeEntry == null)
                    {
                        archive.CreateEntry("DanhSach.csv");
                        readmeEntry = archive.GetEntry("DanhSach.csv");
                    }
                    readAndUpdate(new StreamReader(readmeEntry.Open()));
                }
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
                dtSource.Rows.RemoveAt(count);
                return dtSource;
            }
            catch
            {
                return null;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            using (FileStream zipToOpen = new FileStream(savePath, FileMode.Open))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    if (archive.GetEntry("DanhSach.csv") != null)
                    {
                        archive.GetEntry("DanhSach.csv").Delete();
                    }
                    ZipArchiveEntry readmeEntry = archive.CreateEntry("DanhSach.csv");
                    saveToEntry(GetDataGridViewAsDataTable(customDataGridView1), readmeEntry.Open());
                }
            }
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public static Stream GenerateStreamFromString(string s)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        private void Paste()
        {
            DataObject o = (DataObject)Clipboard.GetDataObject();
            
            if (o.GetDataPresent(DataFormats.StringFormat))
            {

                string s = Clipboard.GetText(TextDataFormat.UnicodeText);
                using (var stream = GenerateStreamFromString(s))
                {
                    StreamReader toRead = new StreamReader(stream);
                    readAndUpdateTab(toRead);
                }
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            customDataGridView1.Rows.Clear();
            Paste();
        }
    }
}
