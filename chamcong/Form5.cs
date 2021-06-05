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

namespace chamcong
{
    public partial class Form5 : Form
    {
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

        private void populateGridName(StreamReader a)
        {
            string line = a.ReadLine();
            int j = 0;
            while (line != null)
            {
                customDataGridView1.Rows.Add();
                int i = 0;
                for (int k = 0; k < line.Length; k++)
                {
                    if (i > 0)
                    {
                        break;
                    }
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

            }
        }
    }
}
