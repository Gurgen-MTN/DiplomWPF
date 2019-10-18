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
    public partial class Schetchiki_izmenenie : Form
    {
        public Schetchiki_izmenenie()
        {
            InitializeComponent();
        }

        SqlConnection sqlConnection;

        private async void Schetchiki_izmenenie_Load(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;

            errorProvider1.Clear();
            listBox1.Items.Clear();
            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True;MultipleActiveResultSets=True";

            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();

            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["CounterID"]).ToString("d7");
                    listBox1.Items.Add(s);
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
            if (textBox3.Text.Trim() == "")
            {
                errorProvider1.SetError(textBox3, "Error set owner");
                return;
            }
            if (textBox4.Text.Trim() == "")
            {
                errorProvider1.SetError(textBox4, "Error set owner phone");
                return;
            }
            if (dateTimePicker1.Text == "0")
            {
                errorProvider1.SetError(dateTimePicker1, "Error set install date");
                return;
            }

            if (dateTimePicker2.Text == "")
            {
                errorProvider1.SetError(dateTimePicker2, "Error set proverka date");
                return;
            }
            

            SqlCommand command = new SqlCommand("UPDATE [Counters] SET  [CounterOwner]=@CounterOwner, [TelephoneOwner]=@TelephoneOwner," +
                "[InstallDate]=@InstallDate ,[ProverkaDate]=@ProverkaDate WHERE [CounterID]=@CounterID", sqlConnection);
            command.Parameters.AddWithValue("CounterID", textBox1.Text);
            command.Parameters.AddWithValue("CounterOwner", textBox3.Text);
            command.Parameters.AddWithValue("TelephoneOwner", textBox4.Text);
            command.Parameters.AddWithValue("InstallDate", dateTimePicker1.Text);
            command.Parameters.AddWithValue("ProverkaDate", dateTimePicker2.Text);

            await command.ExecuteNonQueryAsync();

            textBox3.Text = "";
            textBox4.Text = "";
            dateTimePicker1.Text = DateTime.Now.ToString();
            dateTimePicker2.Text = DateTime.Now.ToString();
            errorProvider1.Clear();
            MessageBox.Show("Hajoxutyamb katarvele popoxum@.", "Izmenenie schetchikov");
        }

    


        private  void Schetchiki_izmenenie_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }

        private async void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string id = listBox1.SelectedItem.ToString();
            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);
            //SqlCommand command1 = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);
            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["CounterID"]).ToString("d7");
                    if (s == id)
                    {
                        textBox1.Text = s;
                        textBox2.Text = ((int)sqlReader["ShkafID"]).ToString("D6");
                        textBox3.Text = sqlReader["CounterOwner"].ToString();
                        textBox4.Text = sqlReader["TelephoneOwner"].ToString();
                        dateTimePicker1.Text = sqlReader["InstallDate"].ToString();
                        dateTimePicker2.Text = sqlReader["ProverkaDate"].ToString();
                    }
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;
        }
    }
}
