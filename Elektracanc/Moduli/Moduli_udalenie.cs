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


namespace Elektracanc.Moduli
{
    public partial class Moduli_udalenie : Form
    {
        public Moduli_udalenie()
        {
            InitializeComponent();
        }
        SqlConnection sqlConnection;

        private async void Moduli_udalenie_Load(object sender, EventArgs e)
        {

            
            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";

            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();


            Update1();

        }
        public async void Update1()
        {
            SqlDataReader sqlReader = null;
            listBox1.Items.Clear();
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
        private async void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                string id = textBox1.Text;
                SqlDataReader sqlReader = null;
                List<int> number = new List<int>(); ;
                //SqlCommand command = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);
                SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);
                try
                {
                    sqlReader = await command.ExecuteReaderAsync();
                    number = new List<int>();
                    while (await sqlReader.ReadAsync())
                    {
                        string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                        if (s == id)
                        {
                            number.Add((int)sqlReader["CounterID"]);

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

                for (int i = 0; i < number.Count; i++)
                {
                    SqlCommand command3 = new SqlCommand("DELETE FROM [CounterStatements] WHERE [CounterID]=@CounterID ", sqlConnection);
                    command3.Parameters.AddWithValue("CounterID", number[i].ToString());
                    await command3.ExecuteNonQueryAsync();
                }



                SqlCommand command1 = new SqlCommand("DELETE FROM [Counters] WHERE [ShkafID]=@ShkafID ", sqlConnection);
                command1.Parameters.AddWithValue("ShkafID", textBox1.Text);
                await command1.ExecuteNonQueryAsync();

                SqlCommand command2 = new SqlCommand("DELETE FROM [ShkafStatements] WHERE [ShkafID]=@ShkafID ", sqlConnection);
                command2.Parameters.AddWithValue("ShkafID", textBox1.Text);
                await command2.ExecuteNonQueryAsync();

                SqlCommand command4 = new SqlCommand("DELETE FROM [Shkafs] WHERE [ShkafID]=@ShkafID ", sqlConnection);
                command4.Parameters.AddWithValue("ShkafID", textBox1.Text);
                await command4.ExecuteNonQueryAsync();
                Update1();

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                textBox6.Clear();
                textBox8.Clear();
                MessageBox.Show("Hajoxutyamb katarvele jnjum@.", "Udalenie");
            }
            else
            {
                MessageBox.Show("Enter modul.");
            }
        
        }


        private void Moduli_udalenie_FormClosing(object sender, FormClosingEventArgs e)
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
                    textBox8.Text = k.ToString();


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
                        textBox2.Text = sqlReader["InstallDate"].ToString();
                        textBox3.Text = sqlReader["ProverkaDate"].ToString();
                        textBox4.Text = sqlReader["Address"].ToString();
                        textBox5.Text = sqlReader["Installer"].ToString();
                        textBox6.Text = sqlReader["CountersQuantity"].ToString();

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
    }
}
