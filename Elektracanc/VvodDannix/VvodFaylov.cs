using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Data.Sql;
using System.Data.SqlClient;
using Microsoft.Win32;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office;
using System.Reflection;

namespace Elektracanc.VvodDannix
{
    public partial class VvodFaylov : Form
    {
        public VvodFaylov()
        {
            InitializeComponent();
        }

        SqlConnection sqlConnection;

        private async void VvodFaylov_Load(object sender, EventArgs e)
        {


            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";

            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();

            var column7 = new DataGridViewColumn();
            column7.HeaderText = "CounterID";
            column7.ReadOnly = true;
            column7.Name = "CounterID";
            column7.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column7);
            var column1 = new DataGridViewColumn();
            column1.HeaderText = "Year";
            column1.ReadOnly = true;
            column1.Name = "Year";
            column1.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column1);
            var column2 = new DataGridViewColumn();
            column2.HeaderText = "Month";
            column2.ReadOnly = true;
            column2.Name = "Month";
            column2.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column2);
            var column3 = new DataGridViewColumn();
            column3.HeaderText = "1Tarif";
            column3.ReadOnly = true;
            column3.Name = "1Tarif";
            column3.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column3);
            var column4 = new DataGridViewColumn();
            column4.HeaderText = "2Tarif";
            column4.ReadOnly = true;
            column4.Name = "2Tarif";
            column4.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column4);
            var column5 = new DataGridViewColumn();
            column5.HeaderText = "3Tarif";
            column5.ReadOnly = true;
            column5.Name = "3Tarif";
            column5.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column5);
            var column6 = new DataGridViewColumn();
            column6.HeaderText = "4Tarif";
            column6.ReadOnly = true;
            column6.Name = "4Tarif";
            column6.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column6);
        }



        private void button_openfile_Click(object sender, EventArgs e)
        {
            //listBox1.Items.Clear();
            s1.Clear();
            s2.Clear();
            s3.Clear();
            s4.Clear();
            s5.Clear();
            s6.Clear();
            s7.Clear();
            openFileDialog1.Filter = "bi files (*.bi)|*.bi";
          
            string s;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog1.FileName;
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                StreamReader w = new StreamReader(fs);
                //label_qanak.Text = s;
                int i = 7;

                while (w.Peek()>0 && i!=0)
                {
                    string st1 = w.ReadLine();
                    string st2 = w.ReadLine();
                    string st3 = w.ReadLine();
                    string st4 = w.ReadLine();
                    string st5 = w.ReadLine();
                    string st6 = w.ReadLine();
                    string st7 = w.ReadLine();
                    s1.Add(st1);
                    s2.Add(st2);
                    s3.Add(st3);
                    s4.Add(st4);
                    s5.Add(st5);
                    s6.Add(st6);
                    s7.Add(st7);
                    dataGridView1.Rows.Add(new string[] { st1,st2,st3,st4,st5,st6,st7});
                    i--;
                }
                w.Close();
            }
        }

        private void VvodFaylov_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }

        List<string> s1 = new List<string>();
        List<string> s2 = new List<string>();
        List<string> s3 = new List<string>();
        List<string> s4 = new List<string>();
        List<string> s5 = new List<string>();
        List<string> s6 = new List<string>();
        List<string> s7 = new List<string>();


        private async void button1_Click(object sender, EventArgs e)
        {

            if (dataGridView1.Rows.Count != 0)
            {

                SqlCommand command = new SqlCommand("INSERT INTO [CounterStatements] (CounterID,Month,Year,Tarif1,Tarif2,Tarif3,Tarif4) " +
                            "VALUES(@CounterID,@Month,@Year,@Tarif1,@Tarif2,@Tarif3,@Tarif4)", sqlConnection);


                for (int i = 0; i < s1.Count; i++)
                {
                    command.Parameters.AddWithValue("CounterID", s1[i]);
                    command.Parameters.AddWithValue("Month", s2[i]);
                    command.Parameters.AddWithValue("Year", s3[i]);
                    command.Parameters.AddWithValue("Tarif1", s4[i]);
                    command.Parameters.AddWithValue("Tarif2", s5[i]);
                    command.Parameters.AddWithValue("Tarif3", s6[i]);
                    command.Parameters.AddWithValue("Tarif4", s7[i]);
                    await command.ExecuteNonQueryAsync();
                }
                
                


                //command.Parameters.AddWithValue("CountersQuantity", numericUpDown1.Text);
                dataGridView1.Rows.Clear();

                MessageBox.Show("Hajoxutyamb katarvele avelacum@.", "Sozdanie");
            }
            else
            {
                MessageBox.Show("@ntreq fayl.", "Vvod faylov");
            }
            
        }
    }
}
