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


namespace Elektracanc.PoiskIOtechtei
{
    public partial class InformaciiOSchetchike : Form
    {
        public InformaciiOSchetchike()
        {
            InitializeComponent();
        }
        SqlConnection sqlConnection;
        List<int> t1 ;
        List<int> t2 ;
        List<int> t3 ;
        List<int> t4 ;
        private async void InformaciiOSchetchike_Load(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;


            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";

            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();

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


        private void InformaciiOSchetchike_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
            }
            else
            {
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
            }

        }




        private async void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string g = listBox1.SelectedItem.ToString();
            listBox2.Items.Clear();
            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");

                    if (s == g)
                    {
                        listBox2.Items.Add(((int)sqlReader["CounterID"]).ToString("d7"));
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
            g = listBox1.SelectedItem.ToString();
            sqlReader = null;
            SqlCommand command1 = new SqlCommand("SELECT * FROM [Shkafs]", sqlConnection);

            try
            {
                sqlReader = await command1.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["ShkafID"]).ToString("d6");

                    if (s == g)
                    {
                        //listBox2.Items.Add(((int)sqlReader["CounterID"]).ToString("d7"));
                        textBox6.Text = sqlReader["Address"].ToString();
                        textBox7.Text = sqlReader["Installer"].ToString();
                        textBox8.Text = sqlReader["InstallDate"].ToString();
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

        private async void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string g = listBox2.SelectedItem.ToString();
            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT * FROM [Counters]", sqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    string s = ((int)sqlReader["CounterID"]).ToString("d7");

                    if (s == g)
                    {
                        textBox9.Text = sqlReader["InstallDate"].ToString();
                        textBox10.Text = sqlReader["CounterOwner"].ToString();
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

        private  void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            if (checkBox1.Checked == true)
            {
                Poisk();
                int sum1 = 0;
                MessageBox.Show("Poisk osuschestvlen");
                for(int i=0; i < t1.Count; i++)
                {
                    sum1 += t1[i];
                    //MessageBox.Show(sum1.ToString());
                }
                textBox1.Text = sum1.ToString();
                int sum2 = 0;
                for(int i = 0; i < t2.Count; i++)
                {
                    sum2 += t2[i];
                }
                textBox2.Text = sum2.ToString();
                int sum3 = 0;
                for(int i = 0; i < t3.Count; i++)
                {
                    sum3 += t3[i];
                }
                textBox3.Text = sum3.ToString();
                int sum4 = 0;
                for(int i = 0; i < t4.Count; i++)
                {
                    sum4 += t4[i];
                }
                textBox4.Text = sum4.ToString();

                textBox5.Text = (sum1 + sum2 + sum3 + sum4).ToString();
            }
            else
            {
                MessageBox.Show("Enter DateTime");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export3.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }
        }

        public void Export_Data_To_Word(DataGridView DGV, string filename)
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
                    headerRange.Text = "Informaciya o schetchike \n \t  Nomer Shkafa:" + listBox1.SelectedItem.ToString() + "\n\t Adress shkafa:" + textBox6.Text
                     +"\n Nomer schetchika: "+listBox2.SelectedItem.ToString()+"\t     Owner name :"+textBox10.Text+"\n Data ustanovki:"+ textBox9.Text +"\n\n"+
                     "Itogoviy rasxod po tarifam    " +textBox1.Text+"    "+textBox2.Text+"    "+textBox3.Text+"    "+textBox4.Text+"\n"+
                     "Summarniy rasxod    "+textBox5.Text
                     ;

                    headerRange.Font.Size = 14;
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


        private async void Poisk()
        {
             t1 = new List<int>();
             t2 = new List<int>();
             t3 = new List<int>();
             t4 = new List<int>();

            string gId = listBox2.SelectedItem.ToString();

            dataGridView1.Rows.Clear();
            string connectionString = @"Data Source=DESKTOP-DMGPN0K\SQLEXPRESS;Initial Catalog=DataBaseEnergy;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();

            SqlCommand command = new SqlCommand("SELECT * FROM [CounterStatements]", sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    string cId = ((int)sqlReader["CounterID"]).ToString("d7");
                    int k1 = int.Parse(sqlReader["Tarif1"].ToString());
                    int k2 = int.Parse(sqlReader["Tarif2"].ToString());
                    int k3 = int.Parse(sqlReader["Tarif3"].ToString());
                    int k4 = int.Parse(sqlReader["Tarif4"].ToString());

                    if (gId == cId)
                    {
                        string year = sqlReader["Year"].ToString();
                        string month = sqlReader["Month"].ToString();
                        //-------------
                        bool amenshe = false;
                        if (dateTimePicker1.Value.Year < int.Parse(year))
                        {
                            amenshe = true;
                        }
                        else if (dateTimePicker1.Value.Year == int.Parse(year) &&
                            dateTimePicker1.Value.Month <= int.Parse(month))
                        {
                            amenshe = true;
                        }
                        else
                        {
                            amenshe = false;
                        }

                        bool abolshe = false;
                        if (dateTimePicker2.Value.Year > int.Parse(year))
                        {
                            abolshe = true;
                        }
                        else if (dateTimePicker2.Value.Year == int.Parse(year) &&
                            dateTimePicker2.Value.Month >= int.Parse(month))
                        {
                            abolshe = true;
                        }
                        else
                        {
                            abolshe = false;
                        }
                        if (amenshe && abolshe)
                        {
                            dataGridView1.Rows.Add(year, month, k1, k2, k3, k4);
                            t1.Add(k1);
                            t2.Add(k2);
                            t3.Add(k3);
                            t4.Add(k4);
                        }


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
            //dateTimePicker1.MaxDate = DateTime.Now;
            dateTimePicker2.MinDate = dateTimePicker1.Value;
        }
    }
}
