using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace Generated_bat
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string puch = @"pc_data.xlsx";
            //MessageBox.Show(puch);
            try
            {
                dataGridView1.Rows.Clear();
                var excelappworkbook = new XLWorkbook(puch);
                //работаем с первым листом добавляем заголовки
                var excelworksheet = excelappworkbook.Worksheet(1);
                int i = 1;
                while (excelworksheet.Row(i+1).Cell(1).Value.ToString() != "")
                {
                    dataGridView1.Rows.Add(1);
                    for (int z=0; z<8; z++)
                    {
                        dataGridView1.Rows[i - 1].Cells[z].Value = excelworksheet.Row(i + 1).Cell(z+1).Value.ToString();
                    }
                    i++;
                }

                excelworksheet.Row(1).Cell(15).Value = "Номер лицевого счета";
                excelworksheet.Row(1).Cell(16).Value = "ID лицевого счета";
                excelworksheet.Row(1).Cell(17).Value = "Наличие в списке";
            }
            catch(SystemException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void bat(string ip, string name, string password, string pc_name)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int z = 0; z < dataGridView1.RowCount; z++)
            {
                bat(dataGridView1[z, 4].Value.ToString(), dataGridView1[z, 5].Value.ToString(), dataGridView1[z, 6].Value.ToString(), dataGridView1[z, 2].Value.ToString());
                MessageBox.Show(dataGridView1[1, z].Value.ToString());
            }
        }
    }
}
