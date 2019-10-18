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
    public partial class Schetchiki_udalenie : Form
    {
        public Schetchiki_udalenie()
        {
            InitializeComponent();
        }
        SqlConnection sqlConnection;
        private async void Schetchiki_udalenie_Load(object sender, EventArgs e)
        {
            errorProvider1.Clear();
            listBox1.Items.Clear();
            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True;MultipleActiveResultSets=True";

            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();

            Update1();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                string id = textBox1.Text;

                SqlCommand command = new SqlCommand("DELETE FROM [Counters] WHERE [CounterID]=@CounterID ", sqlConnection);
                command.Parameters.AddWithValue("CounterID", textBox1.Text);
                await command.ExecuteNonQueryAsync();

                Update1();
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                textBox6.Clear();
                MessageBox.Show("Hajoxutyamb katarvele jnjum@ .", "Udalenie schetchika");
            }
            else
            {
                MessageBox.Show("Enter schetchik.");
            }
            
        }

        public async void Update1()
        {
            SqlDataReader sqlReader = null;
            listBox1.Items.Clear();
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
        private void Schetchiki_udalenie_FormClosing(object sender, FormClosingEventArgs e)
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
                        textBox5.Text = sqlReader["InstallDate"].ToString();
                        textBox6.Text = sqlReader["ProverkaDate"].ToString();
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
    }
}
