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

namespace Elektracanc
{
    public partial class Moduli_Sozdanie : Form
    {
        public Moduli_Sozdanie()
        {
            InitializeComponent();
        }

        SqlConnection sqlConnection;
        //DataTable dt;
        
        private async void Moduli_Sozdanie_Load(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;
            //Base SQL
            //string s = Environment.CurrentDirectory;
            //s += @"\Acba_database.mdf;";
            //Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True
            //string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\DataBaseEnergy.dbo;Integrated Security=True";
            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";

            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();

            //
            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    textBox1.Text = ((int)sqlReader["ShkafID"]+1).ToString("d6");
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString());
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
            //
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            errorProvider1.Clear();
            if (textBox2.Text.Trim() == "")
            {
                errorProvider1.SetError(textBox2, "Error set adress");
                return;
            }
            if (textBox3.Text.Trim() == "")
            {
                errorProvider1.SetError(textBox3, "Error set installer");
                return;
            }
            if (numericUpDown1.Text == "0")
            {
                errorProvider1.SetError(numericUpDown1, "Error set counter");
                return;
            }

            if (textBox5.Text == "")
            {
                errorProvider1.SetError(textBox5, "Error set password1");
                return;
            }

            if (textBox6.Text == "")
            {
                errorProvider1.SetError(textBox6, "Error set password2");
                return;
            }

            if (textBox7.Text == "")
            {
                errorProvider1.SetError(textBox7, "Error set password3");
                return;
            }
            SqlCommand command = new SqlCommand("INSERT INTO [Shkafs] (Address,InstallDate,ProverkaDate,Installer,CountersQuantity,Password1,Password2,Password3) VALUES(@Address,@InstallDate,@ProverkaDate,@Installer,@CountersQuantity,@Password1,@Password2,@Password3)", sqlConnection);

                command.Parameters.AddWithValue("Address", textBox2.Text);
                command.Parameters.AddWithValue("InstallDate", dateTimePicker1.Text);
                command.Parameters.AddWithValue("ProverkaDate", dateTimePicker2.Text);
                command.Parameters.AddWithValue("Installer", textBox3.Text);
                command.Parameters.AddWithValue("CountersQuantity", numericUpDown1.Text);
                command.Parameters.AddWithValue("Password1", textBox5.Text);
                command.Parameters.AddWithValue("Password2", textBox6.Text);
                command.Parameters.AddWithValue("Password3", textBox7.Text);

                await command.ExecuteNonQueryAsync();
            int k = int.Parse(textBox1.Text.ToString());
            k++;
            //MessageBox.Show(k.ToString());
            textBox1.Text = k.ToString("D6");
            textBox2.Text = "";
            textBox3.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            numericUpDown1.Text = "";
            dateTimePicker1.Text = DateTime.Now.ToString();
            dateTimePicker2.Text = DateTime.Now.ToString();
            errorProvider1.Clear();
            MessageBox.Show("Hajoxutyamb katarvele avelacum@.", "Sozdanie");
        }

        private void button_Paroli_Modulya_Vvod_Click(object sender, EventArgs e)
        {
            if (textBox4.Text == "1111")
            {
                textBox5.ReadOnly = false;
                textBox6.ReadOnly = false;
                textBox7.ReadOnly = false;
            }
        }

        private void Moduli_Sozdanie_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //dateTimePicker1.MaxDate = DateTime.Now;
            dateTimePicker2.MinDate = dateTimePicker1.Value;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            
        }
    }
}
