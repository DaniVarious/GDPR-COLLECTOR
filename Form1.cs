using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Schema;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.IO;
using System.Xml;

namespace DE.IT.GDPR_COLLECTOR
{
    public partial class Form1 : Form
    {
        private SqlConnection conn = new SqlConnection(@"Data Source = PRADE-DB-006; Initial Catalog = IT; Integrated Security = True; Encrypt=False");
      

        public Form1()
        {
            InitializeComponent();
        }


        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            conn.Open();

            string sql = "SELECT AC.AccountID,C.ContactID, C.FirstName, C.LastName, C.Z_ID FROM [PRACS-DB-006_2].DebtrakEU_Repl.dbo.tblAccount_Contact AC INNER JOIN [PRACS-DB-006_2].DebtrakEU_Repl.dbo.tblContact C on C.ContactID=AC.ContactID WHERE AC.Accountid =" + textBox4.Text;
            SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);
            var datatable = new DataTable();

            adapter.Fill(datatable);
            listView1.Items.Clear();

            foreach (DataRow row in datatable.Rows)
            {
                //MessageBox.Show(row["ContactID"].ToString());
                listView1.Items.Add(row["ContactID"].ToString());
                listView1.Items[listView1.Items.Count-1].SubItems.Add(row["FirstName"].ToString() + " " + row["LastName"].ToString());
                listView1.Items[listView1.Items.Count - 1].SubItems.Add(row["Z_ID"].ToString());

            }

            conn.Close();

            conn.Open();

            string sql2 = "SELECT Z_ID FROM [PRACS-DB-006_2].DebtrakEU_Repl.dbo.tblaccount AS TA WHERE accountID =" + textBox4.Text;
            SqlCommand command = new SqlCommand(sql2, conn);

            //if (textBox4.Text.Length > 6)
            //{
                textBox1.Text = Convert.ToString(command.ExecuteScalar());
            //}

            conn.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox4.Text = "4548232";
        }

        private void button2_Click(object sender, EventArgs e)
        {

            bool Type15 = false;
            Type15 = checkBox1.Checked;

            string allSelectedDebtors = string.Empty;
                
            foreach(ListViewItem item in this.listView1.Items)
            {
                if(item.Checked)
                {
                    if (allSelectedDebtors == string.Empty)
                    {
                        allSelectedDebtors += item.SubItems[2].Text;
                    }else
                    {
                    allSelectedDebtors = allSelectedDebtors + "," + item.SubItems[2].Text;
                    }     
                }                
            }


            conn.Open();
            string sql3 = string.Empty;
            var dt = new DataSet();

            if (textBox1.Text != string.Empty) 
            {
                if (allSelectedDebtors == string.Empty)
                {

                    if (Type15 == false)
                    {
                        sql3 = "SELECT Historyno, objecttype, objectno, subno, CONVERT(VARCHAR,TIME,120)Time, Type, Collector, Action,CONVERT(varchar(max), CONVERT(VarBinary(MAX), Data))Data FROM nova_light.dbo.nihist WHERE objecttype=1 AND objectno IN (" + textBox1.Text + ") AND subno = 0  ORDER BY time DESC";
                    }
                    else
                    {
                        sql3 = "SELECT Historyno, objecttype, objectno, subno, CONVERT(VARCHAR,TIME,120)Time, Type, Collector, Action,CONVERT(varchar(max), CONVERT(VarBinary(MAX), Data))Data FROM nova_light.dbo.nihist WHERE objecttype=1 AND objectno IN (" + textBox1.Text + ") AND subno = 0  and Type=15 ORDER BY time DESC";
                    }
                }
                else
                {
                    if (Type15 == false)
                    {
                        sql3 = "SELECT Historyno, objecttype, objectno, subno, CONVERT(VARCHAR,TIME,120)Time, Type, Collector, Action,CONVERT(varchar(max), CONVERT(VarBinary(MAX), Data))Data FROM nova_light.dbo.nihist WHERE objecttype=4 AND objectno IN (" + allSelectedDebtors + ") UNION SELECT Historyno, objecttype, objectno, subno, CONVERT(VARCHAR,TIME,120)Time, Type, Collector, Action,CONVERT(varchar(max), CONVERT(VarBinary(MAX), Data))Data FROM nova_light.dbo.nihist WHERE objecttype=1 AND objectno IN (" + textBox1.Text + ") AND subno IN (" + allSelectedDebtors + ") UNION SELECT Historyno, objecttype, objectno, subno, CONVERT(VARCHAR,TIME,120)Time, Type, Collector, Action,CONVERT(varchar(max), CONVERT(VarBinary(MAX), Data))Data FROM nova_light.dbo.nihist WHERE objecttype=1 AND objectno IN (" + textBox1.Text + ") AND subno = 0  ORDER BY time DESC";
                    }
                    else
                    {
                        sql3 = "SELECT Historyno, objecttype, objectno, subno, CONVERT(VARCHAR,TIME,120)Time, Type, Collector, Action,CONVERT(varchar(max), CONVERT(VarBinary(MAX), Data))Data FROM nova_light.dbo.nihist WHERE objecttype=4 AND objectno IN (" + allSelectedDebtors + ") and Type=15 UNION SELECT Historyno, objecttype, objectno, subno, CONVERT(VARCHAR,TIME,120)Time, Type, Collector, Action,CONVERT(varchar(max), CONVERT(VarBinary(MAX), Data))Data FROM nova_light.dbo.nihist WHERE objecttype=1 AND objectno IN (" + textBox1.Text + ") AND subno IN (" + allSelectedDebtors + ") and Type=15 UNION SELECT Historyno, objecttype, objectno, subno, CONVERT(VARCHAR,TIME,120)Time, Type, Collector, Action,CONVERT(varchar(max), CONVERT(VarBinary(MAX), Data))Data FROM nova_light.dbo.nihist WHERE objecttype=1 AND objectno IN (" + textBox1.Text + ") AND subno = 0 and Type=15 ORDER BY time DESC";
                    }
                }
            
                SqlDataAdapter adapter = new SqlDataAdapter(sql3, conn);
                

            //dataset.Tables[0].Columns.Add(new DataColumn() { ColumnName = "Checked", DataType = typeof(bool), DefaultValue = false });
                adapter.Fill(dt);
            }


            while (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Rows.Clear();
            }
            

            foreach (DataRow row in dt.Tables[0].Rows)
            {
                //  row.ItemArray.
                //dataGridView1.Rows.Add("false", row.Cells["Historyno"].Value.ToString());
                //row.Cells["Historyno"].Value.ToString()

                var historyRow = new DataGridViewRow();
                
                // Unchecked checkbox
                historyRow.Cells.Add(new DataGridViewCheckBoxCell() { Value = false });

                // Add data to the row
                historyRow.Cells.Add(new DataGridViewTextBoxCell() { Value = row["Historyno"].ToString() });
                historyRow.Cells.Add(new DataGridViewTextBoxCell() { Value = row["ObjectType"].ToString() });
                historyRow.Cells.Add(new DataGridViewTextBoxCell() { Value = row["ObjectNo"].ToString() });
                historyRow.Cells.Add(new DataGridViewTextBoxCell() { Value = row["SubNo"].ToString() });
                historyRow.Cells.Add(new DataGridViewTextBoxCell() { Value = row["Time"].ToString() });
                historyRow.Cells.Add(new DataGridViewTextBoxCell() { Value = row["Type"].ToString() });
                historyRow.Cells.Add(new DataGridViewTextBoxCell() { Value = row["Collector"].ToString() });
                historyRow.Cells.Add(new DataGridViewTextBoxCell() { Value = row["Action"].ToString() });
                historyRow.Cells.Add(new DataGridViewTextBoxCell() { Value = row["Data"].ToString() });

                dataGridView1.Rows.Add(historyRow);
            }

            dataGridView1.Columns[0].Visible = true;
            dataGridView1.Columns[0].Width = 8;
            dataGridView1.Columns[0].DefaultCellStyle.Alignment= DataGridViewContentAlignment.MiddleLeft;

            conn.Close();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataGridViewCheckBoxColumn SelectedColumn = new DataGridViewCheckBoxColumn();
            {
                SelectedColumn.HeaderText = "";
                SelectedColumn.Name = "Selected";
                SelectedColumn.Width = 8;
                SelectedColumn.TrueValue = true;
                SelectedColumn.FalseValue = false;
                SelectedColumn.IndeterminateValue = true;
                SelectedColumn.CellTemplate = new DataGridViewCheckBoxCell();
            }

            DataGridViewTextBoxColumn HistoryNoColumn = new DataGridViewTextBoxColumn();
            {
                //checkColumn.HeaderText = "";
                HistoryNoColumn.Name = "HistoryNo";
                HistoryNoColumn.Width = 30;
                HistoryNoColumn.CellTemplate = new DataGridViewTextBoxCell();
            }

            DataGridViewTextBoxColumn ObjectTypeColumn = new DataGridViewTextBoxColumn();
            {
                //checkColumn.HeaderText = "";
                ObjectTypeColumn.Name = "ObjectType";
                ObjectTypeColumn.Width = 30;
                ObjectTypeColumn.CellTemplate = new DataGridViewTextBoxCell();
            }

            DataGridViewTextBoxColumn ObjectNoColumn = new DataGridViewTextBoxColumn();
            {
                //checkColumn.HeaderText = "";
                ObjectNoColumn.Name = "ObjectNo";
                ObjectNoColumn.Width = 30;
                ObjectNoColumn.CellTemplate = new DataGridViewTextBoxCell();
            }

            DataGridViewTextBoxColumn SubNoColumn = new DataGridViewTextBoxColumn();
            {
                //checkColumn.HeaderText = "";
                SubNoColumn.Name = "SubNo";
                SubNoColumn.Width = 30;
                SubNoColumn.CellTemplate = new DataGridViewTextBoxCell();
            }

            DataGridViewTextBoxColumn TimeColumn = new DataGridViewTextBoxColumn();
            {
                //checkColumn.HeaderText = "";
                TimeColumn.Name = "Time";
                TimeColumn.Width = 30;
                TimeColumn.CellTemplate = new DataGridViewTextBoxCell();
            }

            DataGridViewTextBoxColumn TypeColumn = new DataGridViewTextBoxColumn();
            {
                //checkColumn.HeaderText = "";
                TypeColumn.Name = "Type";
                TypeColumn.Width = 30;
                TypeColumn.CellTemplate = new DataGridViewTextBoxCell();
            }

            DataGridViewTextBoxColumn CollectorColumn = new DataGridViewTextBoxColumn();
            {
                //checkColumn.HeaderText = "";
                CollectorColumn.Name = "Collector";
                CollectorColumn.Width = 30;
                CollectorColumn.CellTemplate = new DataGridViewTextBoxCell();
            }

            DataGridViewTextBoxColumn ActionColumn = new DataGridViewTextBoxColumn();
            {
                //checkColumn.HeaderText = "";
                ActionColumn.Name = "Action";
                ActionColumn.Width = 30;
                ActionColumn.CellTemplate = new DataGridViewTextBoxCell();
            }
            DataGridViewTextBoxColumn DataColumn = new DataGridViewTextBoxColumn();
            {
                //checkColumn.HeaderText = "";
                DataColumn.Name = "Data";
                DataColumn.Width = 30;
                DataColumn.CellTemplate = new DataGridViewTextBoxCell();
            }


            dataGridView1.Columns.Add(SelectedColumn);
            dataGridView1.Columns.Add(HistoryNoColumn);
            dataGridView1.Columns.Add(ObjectTypeColumn);
            dataGridView1.Columns.Add(ObjectNoColumn);
            dataGridView1.Columns.Add(SubNoColumn);
            dataGridView1.Columns.Add(TimeColumn);
            dataGridView1.Columns.Add(TypeColumn);
            dataGridView1.Columns.Add(CollectorColumn);
            dataGridView1.Columns.Add(ActionColumn);
            dataGridView1.Columns.Add(DataColumn);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //before your loop
            var csv = new StringBuilder();
            string first=String.Empty;
            string HistoryNo = String.Empty;
            string ObjectType = String.Empty;
            string ObjectNo = String.Empty;
            string SubNo = String.Empty;
            string Time = string.Empty;
            string Type = String.Empty;
            string Collector = String.Empty;
            string action=String.Empty ;
            string DataColumnRead=String.Empty;
            


            bool firstLine = false;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {

                if ((bool)row.Cells[0].Value == true)
                {
                    if (dataGridView1[0, row.Index].Value.ToString() == "True")
                    {

                        //MessageBox.Show(dataGridView1.Rows[row.Index].Cells["Historyno"].Value.ToString());
                        if (firstLine == false)
                        {
                            firstLine = true;
                            if (dataGridView1.Rows.Count > 0)
                            {
                                var newLine = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8}", "HistoryNo", "ObjectType", "ObjectNo", "SubNo", "Time", "Type", "Collector", "Action", "Data");
                                csv.AppendLine(newLine);
                            }
                        }

                        HistoryNo = dataGridView1.Rows[row.Index].Cells["HistoryNo"].Value.ToString();
                        ObjectType = dataGridView1.Rows[row.Index].Cells["ObjectType"].Value.ToString();
                        ObjectNo = dataGridView1.Rows[row.Index].Cells["ObjectNo"].Value.ToString();
                        SubNo = dataGridView1.Rows[row.Index].Cells["SubNo"].Value.ToString();
                        Time = dataGridView1.Rows[row.Index].Cells["Time"].Value.ToString();
                        Type = dataGridView1.Rows[row.Index].Cells["Type"].Value.ToString();
                        Collector = dataGridView1.Rows[row.Index].Cells["Collector"].Value.ToString();
                        action = dataGridView1.Rows[row.Index].Cells["Action"].Value.ToString();
                        DataColumnRead = dataGridView1.Rows[row.Index].Cells["Data"].Value.ToString();

                        //in your loop
                        first = DataColumnRead.Replace("@echo off", "");
                        first = first.Replace("chcp 1252", "");
                        first = first.Replace("explorer", "");
                        first = first.Replace("\"", "");
                        first = first.Replace("\r\n", "");
                        first = first.Replace("      START /D", "");
                        //var second = first + "\r\n";
                        //Suggestion made by KyleMit
                        var newLine2 = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8}", HistoryNo, ObjectType, ObjectNo, SubNo, Time, Type, Collector, action, first);
                        csv.AppendLine(newLine2);

                        dataGridView1[0, row.Index].Value = false;

                    }
                }


            }
            //after your loop
            string result;
            result = textBox2.Text ;
            result = result.Replace("\r\n", "");
            result = result + "\\" + textBox3.Text + "_" + DateTime.Now.ToString("yyyy-MM-dd_HHmm") + ".csv";
            //MessageBox.Show(result);
            //\\global.aktivkapital.com\Germany$\Common\Nova-Export\DocExporter\test.txt

            if (csv.Length>0) 
            {
                File.WriteAllText(result, csv.ToString());
                MessageBox.Show(" exported to: " + result, "Export finished");
            }
            



        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            DialogResult dresult = folderDlg.ShowDialog();
            if (dresult == DialogResult.OK)
            {
                textBox2.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
            
        }
    }
}
