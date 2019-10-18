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
    public partial class Moduli_izmenenie : Form
    {
        public Moduli_izmenenie()
        {
            InitializeComponent();
        }

        SqlConnection sqlConnection;
        //DataTable dt;

        private async void Moduli_izmenenie_Load(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;

            listBox1.Items.Clear();
            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";

            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();

            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");
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



        private void Moduli_izmenenie_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }


        public async void Counter()
        {
            string id = listBox1.SelectedItem.ToString();
            SqlDataReader sqlReader = null;
            
            //SqlCommand command = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);
            SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);
            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                int k = 0;
                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                    if (s == id)
                    {
                        k++;
                    }
                    textBox4.Text = k.ToString();
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

        private void button_Paroli_Modulya_Vvod_Click(object sender, EventArgs e)
        {
           if(textBox6.Text=="1111")
           {
                textBox7.ReadOnly = false;
                textBox8.ReadOnly = false;
                textBox9.ReadOnly = false;
           }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            
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

            if (textBox7.Text == "")
            {
                errorProvider1.SetError(textBox7, "Error set password1");
                return;
            }
            if (textBox8.Text == "")
            {
                errorProvider1.SetError(textBox8, "Error set password2");
                return;
            }
            if (textBox9.Text == "")
            {
                errorProvider1.SetError(textBox9, "Error set password3");
                return;
            }

            SqlCommand command = new SqlCommand("UPDATE [Shkafs] SET [Address]=@Address, [InstallDate]=@InstallDate, " +
                "[ProverkaDate]=@ProverkaDate , [Installer]=@Installer , [CountersQuantity]=@CountersQuantity , " +
                "[Password1]=@Password1 , [Password2]=@Password2, [Password3]=@Password3 WHERE [ShkafID]=@ShkafID", sqlConnection);

            command.Parameters.AddWithValue("ShkafID", textBox1.Text);
            command.Parameters.AddWithValue("Address", textBox2.Text);
            command.Parameters.AddWithValue("InstallDate", dateTimePicker1.Text);
            command.Parameters.AddWithValue("ProverkaDate", dateTimePicker2.Text);
            command.Parameters.AddWithValue("Installer", textBox3.Text);
            command.Parameters.AddWithValue("CountersQuantity", numericUpDown1.Text);
            command.Parameters.AddWithValue("Password1", textBox7.Text);
            command.Parameters.AddWithValue("Password2", textBox8.Text);
            command.Parameters.AddWithValue("Password3", textBox9.Text);

            await command.ExecuteNonQueryAsync();

            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            numericUpDown1.Text = "";
            dateTimePicker1.Text = DateTime.Now.ToString();
            dateTimePicker2.Text = DateTime.Now.ToString();
            errorProvider1.Clear();
            MessageBox.Show("Hajoxutyamb katarvele popoxum@.", "Izmenenie");
        }

        private async void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string id = listBox1.SelectedItem.ToString();
            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);
            //SqlCommand command1 = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);
            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {


                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                    if (s == id)
                    {
                        textBox1.Text = s;
                        dateTimePicker1.Text = sqlReader["InstallDate"].ToString();
                        dateTimePicker2.Text = sqlReader["ProverkaDate"].ToString();
                        textBox2.Text = sqlReader["Address"].ToString();
                        textBox3.Text = sqlReader["Installer"].ToString();
                        numericUpDown1.Text = sqlReader["CountersQuantity"].ToString();
                        textBox7.Text = sqlReader["Password1"].ToString();
                        textBox8.Text = sqlReader["Password2"].ToString();
                        textBox9.Text = sqlReader["Password3"].ToString();
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

            Counter();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //dateTimePicker1.MaxDate = DateTime.Now;
            dateTimePicker2.MinDate = dateTimePicker1.Value;
        }
    }
}
