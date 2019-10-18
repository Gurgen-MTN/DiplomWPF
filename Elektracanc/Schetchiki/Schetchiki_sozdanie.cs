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

namespace Elektracanc.Schetchiki
{
    public partial class Schetchiki_sozdanie : Form
    {
        public Schetchiki_sozdanie()
        {
            InitializeComponent();
        }

        SqlConnection sqlConnection;

        private async void Schetchiki_sozdanie_Load(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;

            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";

            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();

            //
            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    textBox1.Text = ((int)sqlReader["CounterID"] + 1).ToString("d7");

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
            //----------------------------------------------------------------------------
            comboBox1.Items.Clear();
            //
            SqlCommand command1 = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);

            try
            {
                sqlReader = await command1.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    string sid = ((int)sqlReader["ShkafID"]).ToString("d6");
                    comboBox1.Items.Add(sid);
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
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            errorProvider1.Clear();
            if (comboBox1.Text == "")
            {
                errorProvider1.SetError(comboBox1, "Error set phone owner");
                return;
            }
            if (textBox2.Text.Trim() == "")
            {
                errorProvider1.SetError(textBox2, "Error set modul");
                return;
            }
            if (textBox3.Text.Trim() == "")
            {
                errorProvider1.SetError(textBox3, "Error set owner");
                return;
            }
            

            if (dateTimePicker1.Text == "")
            {
                errorProvider1.SetError(dateTimePicker1, "Error set date install");
                return;
            }
            if (dateTimePicker2.Text == "")
            {
                errorProvider1.SetError(dateTimePicker2, "Error set date proverki");
                return;
            }


            SqlCommand command = new SqlCommand("INSERT INTO [Counters] (ShkafID,CounterOwner,TelephoneOwner,InstallDate,ProverkaDate) " +
                "VALUES(@ShkafID,@CounterOwner,@TelephoneOwner,@InstallDate,@ProverkaDate)", sqlConnection);

            command.Parameters.AddWithValue("ShkafID", comboBox1.Text );
            command.Parameters.AddWithValue("CounterOwner", textBox2.Text);
            command.Parameters.AddWithValue("TelephoneOwner", textBox3.Text);
            command.Parameters.AddWithValue("InstallDate", dateTimePicker1.Text);
            command.Parameters.AddWithValue("ProverkaDate", dateTimePicker2.Text);
            //command.Parameters.AddWithValue("CountersQuantity", numericUpDown1.Text);

            await command.ExecuteNonQueryAsync();
            int g = int.Parse(textBox1.Text);
            g++;
            textBox1.Text = g.ToString("D7");
            textBox2.Text = "";
            textBox3.Text = "";
            comboBox1.Text = "";

            dateTimePicker1.Text = DateTime.Now.ToString();
            dateTimePicker2.Text = DateTime.Now.ToString();
            errorProvider1.Clear();
            MessageBox.Show("Hajoxutyamb katarvele avelacum@.", "Sozdanie");
        }

        private void Schetchiki_sozdanie_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;
        }
    }
}
