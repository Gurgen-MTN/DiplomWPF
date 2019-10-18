using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office;
using System.Reflection;


namespace Elektracanc.PoiskIOtechtei
{
    public partial class PoiskModuleysRazryajennimiBatareykami : Form
    {
        public PoiskModuleysRazryajennimiBatareykami()
        {
            InitializeComponent();
        }
        SqlConnection sqlConnection;
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
            SqlCommand command = new SqlCommand("SELECT * FROM [ShkafStatements]", sqlConnection);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
            var column1 = new DataGridViewColumn();
            column1.HeaderText = "Nomer Modulya";
            column1.ReadOnly = true;
            column1.Name = "NomerModulya";
            column1.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column1);
            var column2 = new DataGridViewColumn();
            column2.HeaderText = "Adress";
            column2.ReadOnly = true;
            column2.Name = "Adress";
            column2.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column2);
            var column3 = new DataGridViewColumn();
            column3.HeaderText = "Year";
            column3.ReadOnly = true;
            column3.Name = "Year";
            column3.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column3);
            var column4 = new DataGridViewColumn();
            column4.HeaderText = "Month";
            column4.ReadOnly = true;
            column4.Name = "Month";
            column4.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column4);
            var column5 = new DataGridViewColumn();
            column5.HeaderText = "Battery";
            column5.ReadOnly = true;
            column5.Name = "Battery";
            column5.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column5);
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
                    bool b = (bool)sqlReader["Power"];
                    if (b == false)
                    {
                        dataGridView1.Rows.Add(s, SAdress, sqlReader["Year"], sqlReader["Month"], b);
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
        private async void PoiskModuleysRazryajennimiBatareykami_Load(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;

            Poisk();
        }

        private void PoiskModuleysRazryajennimiBatareykami_FormClosing(object sender, FormClosingEventArgs e)
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

        private async void button1_Click(object sender, EventArgs e)
        {
            //int s = dateTimePicker1.Value.Month;
            //MessageBox.Show(s.ToString());
            dataGridView1.Rows.Clear();
            if (checkBox1.Checked==true)
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
                        bool b = (bool)sqlReader["Power"];
                        string year = sqlReader["Year"].ToString();
                        string month = sqlReader["Month"].ToString();
                        //-------------
                        bool amenshe=false;
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

                        if (b == false && amenshe && abolshe)
                        {
                            dataGridView1.Rows.Add(s, SAdress, year, month, b);
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
            else
            {
                MessageBox.Show("Enter DateTime");
            }
        }


        
        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

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
                    headerRange.Text = "Spisok modulei s razryajennimi batareykami   \n  Nachalnoe vremya"+dateTimePicker1.Text+"\n konechnoe vremya"+dateTimePicker2.Text;
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //dateTimePicker1.MaxDate = DateTime.Now;
            dateTimePicker2.MinDate = dateTimePicker1.Value;
        }
    }
}
