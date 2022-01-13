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
            Directory_data();
            string dataname = date();
            StreamWriter file = new StreamWriter(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\data\"+name +"_"+ dataname + ".bat", false, Encoding.UTF8);
            //file.WriteLine(strok_csv);
            file.WriteLine($"# прописать{ip} ip");
            file.WriteLine($"netsh interface ip set address name=\"Ethernet\" static {ip} 255.255.255.0 10.0.77.225 1");
            file.WriteLine($"netsh interface ip set dns \"Ethernet\" static 10.0.77.225");
            file.WriteLine($"#создание юзера");
            file.WriteLine($"net user {name} {password} /add");
            file.WriteLine($"#добавление в группу админов");
            file.WriteLine($"net localgroup Администраторы {name} /add");
            file.WriteLine($"#имя пк и рабочей группы");
            file.WriteLine($"wmic computersystem where name=\"%computername%\" call rename name=\"{pc_name}\"");
            file.WriteLine($"wmic computersystem where name = \"%computername%\" call joindomainorworkgroup name=\"INSP\"");
            file.WriteLine($"pause");
            file.WriteLine($"shutdown /r /t 0");
            file.Close();
        }
        public string date()
        {
            DateTime thisDay = DateTime.Today;
            return thisDay.ToString("dd-MM-yyyy");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            for (int z = 0; z <= dataGridView1.RowCount-2; z++)
            {
                //MessageBox.Show(dataGridView1[4, z].Value.ToString());
                //MessageBox.Show(dataGridView1[4, z].Value.ToString());
                //MessageBox.Show(dataGridView1[4, z].Value.ToString());
                bat(dataGridView1[4,z].Value.ToString(), dataGridView1[5,z].Value.ToString(), dataGridView1[6,z].Value.ToString(), dataGridView1[7,z].Value.ToString());
                
            }
        }
        private void Directory_data()
        {
            if (!File.Exists(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\data"))
            {
                //log("Создаем дирректрию: " + (Application.StartupPath) + @"\data");
                DirectoryInfo di = Directory.CreateDirectory((Application.StartupPath) + @"\data");
            }
        }
    }
}
