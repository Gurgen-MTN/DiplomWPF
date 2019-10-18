using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office;
using System.Reflection;

namespace Elektracanc.PoiskIOtechtei
{
    public partial class SchetchikovSIstekshimSrokomProverki : Form
    {
        public SchetchikovSIstekshimSrokomProverki()
        {
            InitializeComponent();
        }

        SqlConnection sqlConnection;

        private async void Poisk1()
        {
            List<string> AiD = new List<string>();
            List<string> AAdress = new List<string>();
            //dataGridView1.Items.Clear();
            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();

            SqlCommand command1 = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await command1.ExecuteReaderAsync();
                //sqlReader1 = await command1.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                    AiD.Add(s);
                    AAdress.Add(sqlReader["Address"].ToString());
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


            //SqlDataReader sqlReader1 = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
        

            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                //sqlReader1 = await command1.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    string SAdress = "";
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                    for (int i = 0; i < AiD.Count; i++)
                    {
                        if (AiD[i] == s)
                        {
                            SAdress = AAdress[i];
                            //return;
                        }


                    }
                    string st = sqlReader["ProverkaDate"].ToString();
                    string sub = st.Substring(0, st.Length - 6);
                    sub = st[st.Length - 7].ToString() + st[st.Length - 6].ToString() + st[st.Length - 5].ToString() + st[st.Length - 4].ToString();
                    int k = int.Parse(sub);
                    
                    dataGridView1.Rows.Add(((int)sqlReader["CounterID"]).ToString("d7"), sqlReader["CounterOwner"],
                        sqlReader["TelephoneOwner"], sqlReader["InstallDate"], st, k + 5,s, SAdress);
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

        private async void Poisk2()
        {
            List<string> AiD = new List<string>();
            List<string> AAdress = new List<string>();
            //dataGridView1.Items.Clear();
            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();

            SqlCommand command1 = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await command1.ExecuteReaderAsync();
                //sqlReader1 = await command1.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                    AiD.Add(s);
                    AAdress.Add(sqlReader["Address"].ToString());
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


            //SqlDataReader sqlReader1 = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);


            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                //sqlReader1 = await command1.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    string SAdress = "";
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                    for (int i = 0; i < AiD.Count; i++)
                    {
                        if (AiD[i] == s)
                        {
                            SAdress = AAdress[i];
                            //return;
                        }


                    }
                    string st = sqlReader["ProverkaDate"].ToString();
                    string sub = st.Substring(0, st.Length - 6);
                    sub = st[st.Length - 7].ToString() + st[st.Length - 6].ToString() + st[st.Length - 5].ToString() + st[st.Length - 4].ToString();
                    int k = int.Parse(sub);
                    k += 5;
                    if (DateTime.Now.Year >= k)
                    {
                        dataGridView1.Rows.Add(((int)sqlReader["CounterID"]).ToString("d7"), sqlReader["CounterOwner"],
                            sqlReader["TelephoneOwner"], sqlReader["InstallDate"], st, k, s, SAdress);
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

        private async void Poisk()
        {
            List<string> AiD = new List<string>();
            List<string> AAdress = new List<string>();
            //dataGridView1.Items.Clear();
            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();

            SqlCommand command1 = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await command1.ExecuteReaderAsync();
                //sqlReader1 = await command1.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                    AiD.Add(s);
                    AAdress.Add(sqlReader["Address"].ToString());
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


            //SqlDataReader sqlReader1 = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
            var column1 = new DataGridViewColumn();
            column1.HeaderText = "Nomer schetchika";
            column1.ReadOnly = true;
            column1.Name = "NomerSchetchika";
            column1.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column1);
            var column2 = new DataGridViewColumn();
            column2.HeaderText = "Counter Owner";
            column2.ReadOnly = true;
            column2.Name = "Counter Owner";
            column2.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column2);
            var column3 = new DataGridViewColumn();
            column3.HeaderText = "Telephone owner";
            column3.ReadOnly = true;
            column3.Name = "TelephoneOwner";
            column3.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column3);
            var column4 = new DataGridViewColumn();
            column4.HeaderText = "Install date";
            column4.ReadOnly = true;
            column4.Name = "InstallDate";
            column4.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column4);
            var column5 = new DataGridViewColumn();
            column5.HeaderText = "Proverka date";
            column5.ReadOnly = true;
            column5.Name = "ProverkaDate";
            column5.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column5);
            var column6 = new DataGridViewColumn();
            column6.HeaderText = "Date istek sroka proverki";
            column6.ReadOnly = true;
            column6.Name = "DateIstek";
            column6.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column6);
            var column7 = new DataGridViewColumn();
            column7.HeaderText = "Nomer modulya";
            column7.ReadOnly = true;
            column7.Name = "NomerModulya";
            column7.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column7);
            var column8 = new DataGridViewColumn();
            column8.HeaderText = "Adress modulya";
            column8.ReadOnly = true;
            column8.Name = "AdressModulya";
            column8.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column8);

            //SqlCommand command1 = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                //sqlReader1 = await command1.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    string SAdress = "";
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                    for (int i = 0; i < AiD.Count; i++)
                    {
                        if (AiD[i] == s)
                        {
                            SAdress = AAdress[i];
                            //return;
                        }


                    }
                    //string s1 = ((int)sqlReader["ShkafID"]).ToString("d6");
                    //Adress(s);
                    //bool b = (bool)sqlReader["Power"];
                    //if (b == false)
                    //{

                    //}
                    //dataGridView1.Rows.Add(s, SAdress, sqlReader["Year"], sqlReader["Month"], sqlReader["AccessQuantity"], sqlReader["BadAccessQuantity"], sqlReader["Power"]);
                    string st = sqlReader["ProverkaDate"].ToString();
                    string sub = st.Substring(0, st.Length - 6);
                    sub = st[st.Length - 7].ToString() + st[st.Length - 6].ToString() + st[st.Length - 5].ToString() + st[st.Length - 4].ToString();
                    int k = int.Parse(sub);
                    //DateTime dt = new DateTime();
                    //MessageBox.Show(dt.ToString());
                    //DateTime dtIstek = new DateTime(dt.Year + 5, dt.Month, dt.Day);
                    
                    dataGridView1.Rows.Add(((int)sqlReader["CounterID"]).ToString("d7"), sqlReader["CounterOwner"], 
                        sqlReader["TelephoneOwner"], sqlReader["InstallDate"],st ,k+5, s, SAdress);
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

        private void SchetchikovSIstekshimSrokomProverki_Load(object sender, EventArgs e)
        {
            Poisk();
        }

        private void SchetchikovSIstekshimSrokomProverki_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            //int s = dateTimePicker1.Value.Month;
            //MessageBox.Show(s.ToString());
            dataGridView1.Rows.Clear();
            //if (checkBox1.Checked == true)
            //{
                List<string> AiD = new List<string>();
                List<string> AAdress = new List<string>();
                //dataGridView1.Items.Clear();
                string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";
                sqlConnection = new SqlConnection(connectionString);
                await sqlConnection.OpenAsync();

                SqlCommand command1 = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);
                SqlDataReader sqlReader = null;
                try
                {
                    sqlReader = await command1.ExecuteReaderAsync();
                    //sqlReader1 = await command1.ExecuteReaderAsync();
                    while (await sqlReader.ReadAsync())
                    {
                        string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                        AiD.Add(s);
                        AAdress.Add(sqlReader["Address"].ToString());
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


                //SqlDataReader sqlReader1 = null;
                SqlCommand command = new SqlCommand("SELECT * FROM [ShkafStatements]", sqlConnection);
                SqlDataAdapter dataAdapter = new SqlDataAdapter(command);

                //SqlCommand command1 = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);

                try
                {
                    sqlReader = await command.ExecuteReaderAsync();
                    //sqlReader1 = await command1.ExecuteReaderAsync();
                    while (await sqlReader.ReadAsync())
                    {
                        string SAdress = "";
                        string s = ((int)sqlReader["ShkafID"]).ToString("d6");
                        for (int i = 0; i < AiD.Count; i++)
                        {
                            if (AiD[i] == s)
                            {
                                SAdress = AAdress[i];
                            }
                        }
                        //string s1 = ((int)sqlReader["ShkafID"]).ToString("d6");
                        //Adress(s);
                        //bool b = (bool)sqlReader["Power"];
                        string year = sqlReader["Year"].ToString();
                        string month = sqlReader["Month"].ToString();
                        //-------------
                        bool amenshe = false;
                        //if (dateTimePicker1.Value.Year < int.Parse(year))
                        //{
                        //    amenshe = true;
                        //}
                        //else if (dateTimePicker1.Value.Year == int.Parse(year) &&
                        //    dateTimePicker1.Value.Month <= int.Parse(month))
                        //{
                        //    amenshe = true;
                        //}
                        //else
                        //{
                        //    amenshe = false;
                        //}

                        //bool abolshe = false;
                        //if (dateTimePicker2.Value.Year > int.Parse(year))
                        //{
                        //    abolshe = true;
                        //}
                        //else if (dateTimePicker2.Value.Year == int.Parse(year) &&
                        //    dateTimePicker2.Value.Month >= int.Parse(month))
                        //{
                        //    abolshe = true;
                        //}
                        //else
                        //{
                        //    abolshe = false;
                        //}
                        int k = (int)sqlReader["BadAccessQuantity"];
                        //if (k > 0 && amenshe && abolshe)
                        //{
                        //    dataGridView1.Rows.Add(s, SAdress, year, month, sqlReader["AccessQuantity"], k);
                        //}
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
            //}
            //else
            //{
            //    MessageBox.Show("Enter DateTime");
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export2.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }
        }


        private void Export_Data_To_Word(DataGridView DGV, string filename)
            {

                if (DGV.Rows.Count != 0)
                {
                    int RowCount = DGV.Rows.Count;
                    int ColumnCount = DGV.Columns.Count;
                    Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                    //add rows
                    int r = 0;
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        for (r = 0; r <= RowCount - 1; r++)
                        {
                            DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                        } //end row loop
                    } //end column loop

                    Word.Document oDoc = new Word.Document();
                    oDoc.Application.Visible = true;

                    //page orintation
                    oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                    dynamic oRange = oDoc.Content.Application.Selection.Range;
                    string oTemp = "";
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        for (int c = 0; c <= ColumnCount - 1; c++)
                        {
                            oTemp = oTemp + DataArray[r, c] + "\t";

                        }
                    }

                    //table format
                    oRange.Text = oTemp;
                    object oMissing = Missing.Value;
                    object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                    object ApplyBorders = true;
                    object AutoFit = true;
                    object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                    oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                          Type.Missing, Type.Missing, ref ApplyBorders,
                                          Type.Missing, Type.Missing, Type.Missing,
                                          Type.Missing, Type.Missing, Type.Missing,
                                          Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                    oRange.Select();

                    oDoc.Application.Selection.Tables[1].Select();
                    oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                    oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                    oDoc.Application.Selection.Tables[1].Rows[1].Select();
                    oDoc.Application.Selection.InsertRowsAbove(1);
                    oDoc.Application.Selection.Tables[1].Rows[1].Select();

                    //header row style
                    oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                    oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                    oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                    //add header row manually
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                    }

                    //table style 
                    oDoc.Application.Selection.Tables[1].Rows[1].Select();
                    oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    //header text
                    foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                    {
                        Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                        headerRange.Text = "Spisok schetchikov i dati istecheniya proverki  \n ";
                        headerRange.Font.Size = 16;
                        headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }

                    //save the file

                    oDoc.SaveAs(filename, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing);

                    //NASSIM LOUCHANI
                }
            }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                dataGridView1.Rows.Clear();
                Poisk1();
            }
            else if (radioButton2.Checked == true)
            {
                dataGridView1.Rows.Clear();
                Poisk2();
            }
        }
    }
}
