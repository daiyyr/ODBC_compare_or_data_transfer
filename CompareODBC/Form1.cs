using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Web;
using System.Configuration;
using System.Collections;
using System.IO;
using System.Data.Odbc;
using System.Web.Script.Serialization;
using System.Threading;

namespace CompareODBC
{
    public partial class Form1 : Form
    {

        public OdbcConnection conn1;
        public OdbcConnection conn2;
        public OdbcTransaction tran1;
        public OdbcTransaction tran2;

        public Form1()
        {
            InitializeComponent();
            loadBC();
        }


        #region tool function
        public delegate void setLog(RichTextBox rtb, string str1);
        public delegate void setLogWithColor(RichTextBox rtb, string str1, Color color1);
        public void setLogT(RichTextBox r, string s)
        {
            if (r.InvokeRequired)
            {
                setLog sl = new setLog(delegate(RichTextBox rtb, string text)
                {
                    rtb.AppendText(text + Environment.NewLine);
                });
                r.Invoke(sl, r, s);
            }
            else
            {
                r.AppendText(s + Environment.NewLine);
            }
        }

        public void setLogtColorful(RichTextBox r, string s, Color c)//something wrong, if it's first line, no color
        {
            if (r.InvokeRequired)
            {
                setLogWithColor sl = new setLogWithColor(delegate(RichTextBox rtb, string text, Color color)
                {
                    rtb.AppendText(text + Environment.NewLine);
                    int i = rtb.Text.LastIndexOf("\n", rtb.Text.Length - 2);
                    if (i > 1)
                    {
                        rtb.Select(i, rtb.Text.Length);
                        rtb.SelectionColor = color;
                        rtb.Select(i, rtb.Text.Length);
                        rtb.SelectionFont = new Font(rtb.Font, FontStyle.Bold);
                    }
                });
                r.Invoke(sl, r, s, c);
            }
            else
            {
                r.AppendText(s + Environment.NewLine);
                int i = r.Text.LastIndexOf("\n", r.Text.Length - 2);
                if (i > 1)
                {
                    r.Select(i, r.Text.Length);
                    r.SelectionColor = c;
                    r.Select(i, r.Text.Length);
                    r.SelectionFont = new Font(r.Font, FontStyle.Bold);
                }
            }
        }

        public static string StrToDQuoteSQL(string str)
        {
            try
            {
                str = str.Replace("\\", "\\\\");
                str = str.Replace("\"", "\\\"");
                str = "\"" + str + "\"";
                return str;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string DateToSQL(object date)
        {
            try
            {
                if (date == null || date == DBNull.Value || Convert.ToString(date) == "" || Convert.ToString(date).ToLower() == "null") return "null";
                else
                {
                    DateTime dateObj = DateTime.MinValue;
                    dateObj = Convert.ToDateTime(date);
                    if (dateObj == DateTime.MinValue) return "null";
                    else return "'" + dateObj.Year + "-" + dateObj.Month + "-" + dateObj.Day + "'";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private string getNextJournalNum()
        {
            string new_jnl_num;
            OdbcConnection conn2 = new OdbcConnection();
            conn2 = new OdbcConnection(textBox1.Text);
            conn2.Open();
            OdbcTransaction tran2 = conn2.BeginTransaction();
            try
            {
                

                //get next Journal number
                string sql_num;
                sql_num = "select CONCAT_WS('', system_value, pilot) AS new_jnl_num ";
                sql_num += " from system ";
                sql_num += "      inner join (select cast(system_value as unsigned)+1 as pilot from system where system_code = 'JOURNALPILOT') as T1 ";
                sql_num += " where system_code = 'JOURNALPREFIX' ";

                OdbcCommand comm2 = new OdbcCommand(sql_num, conn2, tran2);
                OdbcDataReader reader = comm2.ExecuteReader();

                if (reader.HasRows)
                {
                    // get new num
                    reader.Read();
                    new_jnl_num = reader["new_jnl_num"].ToString();
                    reader.Close();

                    // update DB num
                    sql_num = " update system set system_value = system_value + 1 where system_code = 'JOURNALPILOT' ";
                    comm2 = new OdbcCommand(sql_num, conn2, tran2);
                    comm2.ExecuteNonQuery();
                    comm2 = new OdbcCommand("SELECT LAST_INSERT_ID()", conn2, tran2);
                    int system_id = Convert.ToInt32(comm2.ExecuteScalar());
                    comm2.Dispose();
                    tran2.Commit();
                    return new_jnl_num;
                }
                else
                {
                    comm2.Dispose();
                    reader.Close();
                    string msg = "Can not get data from system, cd: 'JOURNALPREFIX' & 'JOURNALPILOT'";
                    setLogT(logT, DateTime.Now.ToString() + " " + msg);
                    throw new Exception(msg);
                }
                tran2.Commit();

            }
            catch (Exception ex)
            {
                if (conn2.State == ConnectionState.Open)
                {
                    tran2.Rollback();
                }
                setLogT(logT, DateTime.Now.ToString() + " " + "Exception: " + ex.Message);
                MessageBox.Show("Error.");
                throw ex;
            }
            finally
            {
                if (conn2 != null)
                {
                    conn2.Close();
                }
            }
        }

        private void deleteBC(string id)
        {
            
            // Bodycorp_ID
            string bodycorp_id = id;

            // DB connection string
            try
            {
                #region delete bc
                // cinvoice
                string sql1 = " DELETE FROM cinvoice_gls WHERE cinvoice_gl_cinvoice_id IN ( SELECT cinvoice_id FROM cinvoices WHERE cinvoice_bodycorp_id = " + bodycorp_id + " )"
                            + " or cinvoice_gl_gl_id in ( select gl_transaction_id from gl_transactions where gl_transaction_bodycorp_id = " + bodycorp_id + " ) ";
                string sql2 = " DELETE FROM cinvoices WHERE cinvoice_bodycorp_id = " + bodycorp_id;

                // cpayments
                string sql3 = " DELETE FROM cpayment_gls WHERE cpayment_gl_cpayment_id IN ( SELECT cpayment_id FROM cpayments WHERE cpayment_bodycorp_id = " + bodycorp_id + " )"
                            + " or cpayment_gl_gl_id in ( select gl_transaction_id from gl_transactions where gl_transaction_bodycorp_id = " + bodycorp_id + " ) ";
                string sql4 = " DELETE FROM cpayments WHERE cpayment_bodycorp_id = " + bodycorp_id;

                // invoice_master
                string sql5 = " DELETE FROM invoice_gls WHERE invoice_gl_invoice_id IN ( SELECT invoice_master_id FROM invoice_master WHERE invoice_master_bodycorp_id = " + bodycorp_id + " )"
                            + " or invoice_gl_gl_id in ( select gl_transaction_id from gl_transactions where gl_transaction_bodycorp_id = " + bodycorp_id + " ) ";
                string sql6 = " DELETE FROM invoice_master WHERE invoice_master_bodycorp_id = " + bodycorp_id;

                // receipts
                string sql7 = " DELETE FROM receipt_gls WHERE receipt_gl_receipt_id IN ( SELECT receipt_id FROM receipts WHERE receipt_bodycorp_id = " + bodycorp_id + " )"
                            + " or receipt_gl_gl_id in ( select gl_transaction_id from gl_transactions where gl_transaction_bodycorp_id = " + bodycorp_id + " ) ";
                string sql8 = " DELETE FROM receipts WHERE receipt_bodycorp_id = " + bodycorp_id;

                // gl_transactions
                string sql9 = " DELETE FROM gl_tran_gls WHERE gl_tran_gl_offset_id IN ( SELECT gl_transaction_id FROM gl_transactions WHERE gl_transaction_bodycorp_id = " + bodycorp_id + " )"
                                    + " or gl_tran_gl_parent_id in ( SELECT gl_transaction_id FROM gl_transactions WHERE gl_transaction_bodycorp_id = " + bodycorp_id + " )";
                string sql10 = " DELETE FROM gl_transactions WHERE gl_transaction_bodycorp_id = " + bodycorp_id;

                // bodycorp_comms
                string sql11 = " DELETE FROM bodycorp_comms WHERE bodycorp_comm_bodycorp_id = " + bodycorp_id;

                // bodycorp_managers
                string sql12 = " DELETE FROM bodycorp_managers WHERE bodycorp_manager_bodycorp_id = " + bodycorp_id;

                // budget_fields
                string sql13 = " DELETE FROM budget_fields WHERE budget_field_bodycorp_id = " + bodycorp_id;

                // budget_master
                string sql14 = " DELETE FROM budget_master WHERE budget_master_bodycorp_id = " + bodycorp_id;

                // utility_master
                string sql15 = " DELETE FROM utility_master WHERE utility_master_unit_id in ( select unit_master_id FROM unit_master WHERE unit_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) )";

                // ownerships
                string sql16 = " DELETE FROM ownerships WHERE ownership_unit_id in ( select unit_master_id FROM unit_master WHERE unit_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) )";

                // unit_master
                DataTable unit = new DataTable();
                string sql_u = "SELECT `unit_master_id` FROM `unit_master` WHERE unit_master_property_id not in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) ";
                OdbcCommand cmd = new OdbcCommand(sql_u, conn2, tran2);
                OdbcDataAdapter da = new OdbcDataAdapter(cmd);
                da.Fill(unit);
                string unitID = unit.Rows[0]["unit_master_id"].ToString();
                string sql17_1 = " Update gl_transactions set gl_transaction_unit_id = " + unitID + " WHERE gl_transaction_unit_id in "
                                + "( select unit_master_id FROM unit_master WHERE unit_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " )  )";
                string sql17_2 = " Update receipts set receipt_unit_id = " + unitID + " WHERE receipt_unit_id in "
                                + "( select unit_master_id FROM unit_master WHERE unit_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " )  )";
                string sql17_3 = " Update invoice_master set invoice_master_unit_id = " + unitID + " WHERE invoice_master_unit_id in "
                                + "( select unit_master_id FROM unit_master WHERE unit_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " )  )";
                string sql17_4 = " Update cinvoices set cinvoice_unit_id = " + unitID + " WHERE cinvoice_unit_id in "
                                + "( select unit_master_id FROM unit_master WHERE unit_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " )  )";
                string sql17_5 = " Update unit_master as u1 INNER JOIN unit_master as u2 on u1.unit_master_principal_id = u2.unit_master_id "
                                    + " set u1.unit_master_principal_id = " + unitID + " WHERE u2.unit_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " )";

                string sql17 = " DELETE FROM unit_master WHERE unit_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) ";

                // pptyvt_master
                string sql18 = " DELETE FROM pptyvt_master WHERE pptyvt_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) ";

                // pptycntr_master
                string sql19 = " DELETE FROM pptycntr_master WHERE pptycntr_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) ";

                // pptyins_master
                string sql20 = " DELETE FROM pptyins_master WHERE pptyins_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) ";

                // pptymaint_master
                string sql21 = " DELETE FROM pptymaint_master WHERE pptymaint_master_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) ";

                // pptymtgs
                string sql22 = " DELETE FROM pptymtgs WHERE pptymtg_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) ";

                // pptytitles
                string sql23 = " DELETE FROM pptytitles WHERE pptytitle_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) ";

                //pptycontacts
                string sql24_01 = " DELETE FROM property_contacts WHERE property_contact_property_id in ( select property_master_id FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id + " ) ";

                // property_master
                string sql24 = " DELETE FROM property_master WHERE property_master_bodycorp_id = " + bodycorp_id;

                // purchorder_master
                string sql25 = " DELETE FROM purchorder_master WHERE purchorder_master_bodycorp_id = " + bodycorp_id;

                //debtor_master
                DataTable bc = new DataTable();
                string sql_bc = "SELECT `bodycorp_id` FROM `bodycorps` WHERE bodycorp_id <> " + bodycorp_id;
                cmd = new OdbcCommand(sql_bc, conn2, tran2);
                da = new OdbcDataAdapter(cmd);
                da.Fill(bc);
                string bcID = bc.Rows[0]["bodycorp_id"].ToString();
                string sql26 = " Update debtor_master set debtor_master_bodycorp_id = " + bcID + " WHERE debtor_master_bodycorp_id = " + bodycorp_id;

                //bodycorps
                string sql27 = "DELETE FROM `bodycorps` WHERE `bodycorp_id`=" + bodycorp_id;

                List<string> sqlL = new List<string>();
                sqlL.Add(sql1);
                sqlL.Add(sql2);
                sqlL.Add(sql3);
                sqlL.Add(sql4);
                sqlL.Add(sql5);
                sqlL.Add(sql6);
                sqlL.Add(sql7);
                sqlL.Add(sql8);
                sqlL.Add(sql9);
                sqlL.Add(sql10);
                sqlL.Add(sql11);
                sqlL.Add(sql12);
                sqlL.Add(sql13);
                sqlL.Add(sql14);
                sqlL.Add(sql15);
                sqlL.Add(sql16);
                sqlL.Add(sql17_1);
                sqlL.Add(sql17_2);
                sqlL.Add(sql17_3);
                sqlL.Add(sql17_4);
                sqlL.Add(sql17_5);
                sqlL.Add(sql17);
                sqlL.Add(sql18);
                sqlL.Add(sql19);
                sqlL.Add(sql20);
                sqlL.Add(sql21);
                sqlL.Add(sql22);
                sqlL.Add(sql23);
                sqlL.Add(sql24_01);
                sqlL.Add(sql24);
                sqlL.Add(sql25);
                sqlL.Add(sql26);
                sqlL.Add(sql27);

                OdbcCommand comm = conn2.CreateCommand();
                comm.Connection = conn2;
                comm.Transaction = tran2;
                foreach (string s in sqlL)
                {
                    comm.CommandText = s;
                    comm.ExecuteScalar();
                }

            #endregion

                #region Clear Debtors and Creditors without reference
                //creditor_comms
                string csql1 = " DELETE FROM creditor_comms WHERE creditor_comm_creditor_id in( " +
                                "SELECT creditor_master_id FROM creditor_master  " +
                                "where creditor_master_id not in (select cinvoice_creditor_id from cinvoices)  " +
                                "and creditor_master_id not in (select cpayment_creditor_id from cpayments)  " +
                                "and creditor_master_id not in (select pptycntr_master_creditor_id from pptycntr_master)  " +
                                "and creditor_master_id not in (select pptymaint_master_creditor_id from pptymaint_master) " +
                                "and creditor_master_id not in (select purchorder_master_creditor_id from purchorder_master)  )";

                //debtor_master
                string csql2 = "DELETE FROM creditor_master  " +
                                "where creditor_master_id not in (select cinvoice_creditor_id from cinvoices)  " +
                                "and creditor_master_id not in (select cpayment_creditor_id from cpayments)  " +
                                "and creditor_master_id not in (select pptycntr_master_creditor_id from pptycntr_master)  " +
                                "and creditor_master_id not in (select pptymaint_master_creditor_id from pptymaint_master) " +
                                "and creditor_master_id not in (select purchorder_master_creditor_id from purchorder_master) ";

                //debtor_comms
                string csql3 = " DELETE FROM debtor_comms WHERE debtor_comm_debtor_id in( " +
                                "SELECT debtor_master_id FROM debtor_master where debtor_master_id not in (select invoice_master_debtor_id from invoice_master)  " +
                                "and debtor_master_id not in (select ownership_debtor_id from ownerships)  " +
                                "and debtor_master_id not in (select receipt_debtor_id from receipts)  " +
                                "and debtor_master_id not in (select unit_master_debtor_id from unit_master) )";

                //debtor_master
                string csql4 = "DELETE FROM debtor_master where debtor_master_id not in (select invoice_master_debtor_id from invoice_master)  " +
                                "and debtor_master_id not in (select ownership_debtor_id from ownerships)  " +
                                "and debtor_master_id not in (select receipt_debtor_id from receipts)  " +
                                "and debtor_master_id not in (select unit_master_debtor_id from unit_master) ";

                //comm_master
                string csql5 = "DELETE FROM comm_master where comm_master_id not in (select bodycorp_comm_comm_id from bodycorp_comms)and comm_master_id not in (select creditor_comm_comm_id from creditor_comms) "
                    + " and comm_master_id not in (select debtor_comm_comm_id from debtor_comms) and comm_master_id not in ( select property_comm_comm_id from property_comms) and comm_master_id not in (select account_comm_comm_id from account_comms)"
                    + " and comm_master_id not in (select contact_comm_comm_id from contact_comms)";

                List<string> csqlL = new List<string>();
                csqlL.Add(csql1);
                csqlL.Add(csql2);
                csqlL.Add(csql3);
                csqlL.Add(csql4);
                csqlL.Add(csql5);
                comm = conn2.CreateCommand();
                comm.Connection = conn2;
                comm.Transaction = tran2;
                foreach (string s in csqlL)
                {
                    comm.CommandText = s;
                    comm.ExecuteScalar();
                }
                #endregion

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion


        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            conn1 = new OdbcConnection();
            conn2 = new OdbcConnection();
            conn1 = new OdbcConnection(textBox1.Text);
            conn2 = new OdbcConnection(textBox2.Text);
            conn1.Open();
            conn2.Open();
            tran1 = conn1.BeginTransaction();
            tran2 = conn2.BeginTransaction();
            try
            {
                setLogT(logT, DateTime.Now.ToString() + " " + "Start comparing...");
                //string connectionString = "DSN=Test;UID=Chester;Pwd=Tester;";
                //string connectionString = "SERVER=localhost;dsn=sapp_sms;DATABASE=sapp_sms;UID=root;PASSWORD=password;";

                
                DataTable Tables1 = conn1.GetSchema("Tables");
                DataTable Tables2 = conn2.GetSchema("Tables");

                #region table1
                foreach (DataRow dr1 in Tables1.Rows)
                {
                    bool tableMatch = false;
                    foreach (DataRow dr2 in Tables2.Rows)
                    {
                        if (dr1["TABLE_NAME"].ToString() == dr2["TABLE_NAME"].ToString())
                        {
                            string sql = "SELECT * from " + dr1["TABLE_NAME"].ToString() + " LIMIT 1";
                            OdbcCommand comm1 = new OdbcCommand(sql, conn1, tran1);
                            OdbcDataReader reader1 = comm1.ExecuteReader();
                            OdbcCommand comm2 = new OdbcCommand(sql, conn2, tran2);
                            OdbcDataReader reader2 = comm2.ExecuteReader();
                            int fieldCount1 = reader1.FieldCount;
                            int fieldCount2 = reader2.FieldCount;

                            bool firstUnmatchedColumn = true;
                            for (int i = 0; i < fieldCount1; i++)
                            {
                                bool columnMatch = false;
                                for (int j = 0; j < fieldCount2; j++)
                                {

                                    if (reader1.GetName(i) == reader2.GetName(j))
                                    {
                                        if (reader1.GetDataTypeName(i) == reader2.GetDataTypeName(j))
                                        {
                                            columnMatch = true;
                                            break;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                                if (!columnMatch)
                                {
                                    if (firstUnmatchedColumn)
                                    {
                                        setLogT(richTextBox1, dr1["TABLE_NAME"].ToString());
                                        firstUnmatchedColumn = false;
                                    }
                                    setLogtColorful(richTextBox1, "       " + reader1.GetName(i) + ": " + reader1.GetDataTypeName(i), Color.Red);
                                }

                            }

                            reader1.Close();
                            reader2.Close();
                            comm1.Dispose();
                            comm2.Dispose();

                            tableMatch = true;
                            break;
                        }
                    }
                    if (!tableMatch)
                    {
                        setLogtColorful(richTextBox1, dr1["TABLE_NAME"].ToString(), Color.Red);
                    }
                }
                #endregion

                #region table2
                foreach (DataRow dr2 in Tables2.Rows)
                {
                    bool tableMatch = false;
                    foreach (DataRow dr1 in Tables1.Rows)
                    {
                        if (dr1["TABLE_NAME"].ToString() == dr2["TABLE_NAME"].ToString())
                        {
                            string sql = "SELECT * from " + dr1["TABLE_NAME"].ToString() + " LIMIT 1";
                            OdbcCommand comm1 = new OdbcCommand(sql, conn1, tran1);
                            OdbcDataReader reader1 = comm1.ExecuteReader();
                            OdbcCommand comm2 = new OdbcCommand(sql, conn2, tran2);
                            OdbcDataReader reader2 = comm2.ExecuteReader();
                            int fieldCount1 = reader1.FieldCount;
                            int fieldCount2 = reader2.FieldCount;

                            bool firstUnmatchedColumn = true;
                            for (int i = 0; i < fieldCount2; i++)
                            {
                                bool columnMatch = false;
                                for (int j = 0; j < fieldCount1; j++)
                                {
                                    if (reader1.GetName(j) == reader2.GetName(i))
                                    {
                                        if (reader1.GetDataTypeName(j) == reader2.GetDataTypeName(i))
                                        {
                                            columnMatch = true;
                                            break;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                                if (!columnMatch)
                                {
                                    if (firstUnmatchedColumn)
                                    {
                                        setLogT(richTextBox2, dr2["TABLE_NAME"].ToString());
                                        firstUnmatchedColumn = false;
                                    }
                                    setLogtColorful(richTextBox2, "       " + reader2.GetName(i) + ": " + reader2.GetDataTypeName(i), Color.Red);
                                }

                            }

                            reader1.Close();
                            reader2.Close();
                            comm1.Dispose();
                            comm2.Dispose();

                            tableMatch = true;
                            break;
                        }
                    }
                    if (!tableMatch)
                    {
                        setLogtColorful(richTextBox2, dr2["TABLE_NAME"].ToString(), Color.Red);
                    }
                }
                #endregion


                //while (dr.Read())
                //{
                //    Console.WriteLine(dr.GetValue(0).ToString());
                //    Console.WriteLine(dr.GetValue(1).ToString());
                //    Console.WriteLine(dr.GetValue(2).ToString());
                //}
                //dr.Close();
                //comm.Dispose();

                tran1.Commit();
                tran2.Commit();
                setLogT(logT, DateTime.Now.ToString() + " " + "Comparison complete!");
                conn1.Close();
                conn2.Close();

                conn1.Dispose();
                conn2.Dispose();
            }
            catch (Exception ex)
            {
                if (conn1.State == ConnectionState.Open)
                {
                    tran1.Rollback();
                }
                if (conn2.State == ConnectionState.Open)
                {
                    tran2.Rollback();
                }
                setLogT(logT, DateTime.Now.ToString() + " " + "Exception: " + ex.Message);
                MessageBox.Show("Error in comparing.");
                //throw ex;
            }
            finally
            {
                if (conn1 != null)
                {
                    conn1.Close();
                }
                if (conn2 != null)
                {
                    conn2.Close();
                }
            }

        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Invoke((MethodInvoker)delegate()
            {
                button4.Enabled = false;
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                checkedListBox1.Enabled = false;
                button7.Enabled = false;
                button6.Enabled = false;
                button5.Enabled = false;
                button2.Enabled = false;
            });
            string currentBCCode = "";
            string currentBCName = "";
            string currentBCId = "";
            bool check = false;
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {

                #region for one BC
                this.Invoke((MethodInvoker)delegate()
                {
                     CheckState st = checkedListBox1.GetItemCheckState(i);
                     if (st == CheckState.Checked)
                     {
                         checkedListBox1.SelectedIndex = i;
                         Thread.Sleep(10);
                         currentBCCode = textBox3.Text;
                         currentBCName = textBox4.Text;
                         currentBCId = ((DataRowView)comboBox1.SelectedItem).Row["bodycorp_id"].ToString();
                         check = true;
                     }
                     else
                     {
                         check = false;
                     }
                });

                if (check)
                {
                    conn1 = new OdbcConnection();
                    conn2 = new OdbcConnection();
                    conn1 = new OdbcConnection(textBox1.Text);
                    conn2 = new OdbcConnection(textBox2.Text);
                    conn1.Open();
                    conn2.Open();
                    tran1 = conn1.BeginTransaction();
                    tran2 = conn2.BeginTransaction();
                    try
                    #region try body
                    {
                        setLogT(logT, "Start transfer BC: " + currentBCCode + " | " + currentBCName + "...");

                        #region check public chart account exist in ODBC2
                        string sql = "SELECT chart_master_name,chart_master_type_id,chart_master_bank_account from chart_master where chart_master_id in ( "
                                + " select gl_transaction_chart_id from gl_transactions group by gl_transaction_chart_id )";
                        OdbcCommand comm1 = new OdbcCommand(sql, conn1, tran1);
                        OdbcDataAdapter da = new OdbcDataAdapter(comm1);
                        DataTable ChartDt1 = new DataTable();
                        DataTable ChartDt2 = new DataTable();
                        da.Fill(ChartDt1);
                        da.Dispose();
                        sql = "SELECT chart_master_name,chart_master_type_id,chart_master_bank_account from chart_master";
                        OdbcCommand comm2 = new OdbcCommand(sql, conn2, tran2);
                        da = new OdbcDataAdapter(comm2);
                        da.Fill(ChartDt2);
                        da.Dispose();
                        bool chartAccountCovered = true;
                        foreach (DataRow dr1 in ChartDt1.Rows)
                        {
                            bool exist = false;
                            foreach (DataRow dr2 in ChartDt2.Rows)
                            {
                                if (dr1["chart_master_name"].ToString() == dr2["chart_master_name"].ToString()
                                    && dr1["chart_master_type_id"].ToString() == dr2["chart_master_type_id"].ToString()
                                    && dr1["chart_master_bank_account"].ToString() == dr2["chart_master_bank_account"].ToString()
                                    )
                                {
                                    exist = true;
                                    break;
                                }
                            }
                            if (!exist)
                            {
                                chartAccountCovered = false;
                                setLogT(logT, "Error! Chart Account missed in ODBC2, chart_master_name: " + dr1["chart_master_name"].ToString() + ", chart_master_type_id: " + dr1["chart_master_type_id"].ToString() + ", chart_master_bank_account: " + dr1["chart_master_bank_account"].ToString());
                            }
                        }
                        if (!chartAccountCovered)
                        {
                            return;
                        }

                        #endregion

                        #region add bc account
                        //check if exist
                        sql = "select * from chart_master where chart_master_code = " + "'BC" + currentBCCode + "' ";
                        comm2 = new OdbcCommand(sql, conn2, tran2);
                        OdbcDataReader dbrd = comm2.ExecuteReader();
                        int newBCAccountId = 0;
                        if (dbrd.HasRows && dbrd.Read())
                        {
                            newBCAccountId = int.Parse(dbrd["chart_master_id"].ToString());
                        }
                        else
                        {
                            string sql_cm = "insert chart_master (";
                            sql_cm += "chart_master_code, chart_master_name, chart_master_type_id, ";
                            sql_cm += "chart_master_notax, chart_master_bank_account, chart_master_trust_account, chart_master_levy_base, chart_master_inactive";
                            sql_cm += ") values(";
                            sql_cm += "'" + "BC" + currentBCCode + "', ";                    // chart_master_code
                            sql_cm += "'Body Corporate " + currentBCCode + "', ";   // chart_master_name
                            sql_cm += "3, 0, 1, 0, 0, 0)";
                            comm2 = new OdbcCommand(sql_cm, conn2, tran2);
                            comm2.ExecuteNonQuery();
                            comm2 = new OdbcCommand("SELECT LAST_INSERT_ID()", conn2, tran2);
                            newBCAccountId = Convert.ToInt32(comm2.ExecuteScalar());
                            setLogT(logT, "Add " + 1 + " BC Account.");
                        }
                        #endregion

                        #region add bc
                        //check if exist
                        sql = "select * from bodycorps where bodycorp_code = '" + currentBCCode + "'";
                        comm2 = new OdbcCommand(sql, conn2, tran2);
                        dbrd = comm2.ExecuteReader();
                        if (dbrd.HasRows && dbrd.Read())
                        {
                            setLogtColorful(logT, "BC " + currentBCCode + " | " + currentBCName + " exist.", Color.Red);
                            setLogtColorful(logT, "Delete BC " + currentBCCode + " | " + currentBCName, Color.Red);
                            deleteBC(dbrd["bodycorp_id"].ToString());
                        }

                        sql = "select * from bodycorps where bodycorp_id = " + currentBCId;
                        comm1 = new OdbcCommand(sql, conn1, tran1);
                        da = new OdbcDataAdapter(comm1);
                        DataTable bcTD1 = new DataTable();
                        da.Fill(bcTD1);
                        da.Dispose();

                        List<UtilityTemplateObject_BCnote> newL = new List<UtilityTemplateObject_BCnote>();
                        UtilityTemplateObject_BCnote note1 = new UtilityTemplateObject_BCnote();
                        note1.Id = Guid.NewGuid().ToString();
                        note1.Date = DateTime.Now;
                        note1.Title = "Untitled";
                        note1.Details = bcTD1.Rows[0]["bodycorp_notes"].ToString();
                        newL.Add(note1);
                        JavaScriptSerializer json_serializer = new JavaScriptSerializer();
                        string JsonNote = json_serializer.Serialize(newL);

                        sql = "insert into bodycorps set bodycorp_code='" + currentBCCode + "', bodycorp_name='" + currentBCName + "', bodycorp_nogst=0, bodycorp_discount=0, bodycorp_account_id=" + newBCAccountId +
                            ", bodycorp_close_off=" + "'0000-00-00', bodycorp_inactive=0, bodycorp_begin_date=" + DateToSQL(DateTime.Parse(bcTD1.Rows[0]["bodycorp_begin_date"].ToString()))
                            + ", bodycorp_notes=" + StrToDQuoteSQL(JsonNote);
                        comm2 = new OdbcCommand(sql, conn2, tran2);
                        comm2.ExecuteNonQuery();
                        comm2 = new OdbcCommand("SELECT LAST_INSERT_ID()", conn2, tran2);
                        int newBCId = Convert.ToInt32(comm2.ExecuteScalar());
                        setLogT(logT, "Add " + 1 + " BC.");
                        #endregion

                        #region add property
                        sql = "select * from property_master where property_master_bodycorp_id = " + currentBCId;
                        comm1 = new OdbcCommand(sql, conn1, tran1);
                        da = new OdbcDataAdapter(comm1);
                        DataTable pmDT = new DataTable();
                        da.Fill(pmDT);
                        da.Dispose();

                        string sql_pm = "insert property_master (";
                        sql_pm += " property_master_bodycorp_id, property_master_code, property_master_type_id, ";
                        sql_pm += " property_master_name, property_master_begin_date";
                        sql_pm += ") values(";
                        sql_pm += newBCId + ", ";     //property_master_bodycorp_id
                        sql_pm += "'" + pmDT.Rows[0]["property_master_code"].ToString() + "', ";         //property_master_code
                        sql_pm += "1, ";                //property_master_type_id
                        sql_pm += "'" + pmDT.Rows[0]["property_master_name"].ToString() + "', ";            //property_master_name
                        sql_pm += DateToSQL(DateTime.Parse(pmDT.Rows[0]["property_master_begin_date"].ToString())) + " )";      //property_master_begin_date
                        comm2 = new OdbcCommand(sql_pm, conn2, tran2);
                        comm2.ExecuteNonQuery();
                        comm2 = new OdbcCommand("SELECT LAST_INSERT_ID()", conn2, tran2);
                        int propertyId = Convert.ToInt32(comm2.ExecuteScalar());

                        setLogT(logT, "Add " + 1 + " property.");
                        #endregion

                        #region add bc comm
                        sql = "select * from bodycorp_comms where bodycorp_comm_bodycorp_id = " + currentBCId;
                        comm1 = new OdbcCommand(sql, conn1, tran1);
                        da = new OdbcDataAdapter(comm1);
                        DataTable BcCommDT = new DataTable();
                        da.Fill(BcCommDT);
                        da.Dispose();
                        int addBCCommCount = 0;
                        foreach (DataRow dr in BcCommDT.Rows)
                        {
                            sql = "select * from comm_master where comm_master_id = " + dr["bodycorp_comm_comm_id"].ToString();
                            comm1 = new OdbcCommand(sql, conn1, tran1);
                            dbrd = comm1.ExecuteReader();
                            if (!dbrd.HasRows || !dbrd.Read())
                            {
                                throw new Exception("Error in select comm_master");
                            }

                            sql = "insert into comm_master set "
                                + " comm_master_type_id = " + dbrd["comm_master_type_id"].ToString() + ", "
                                + " comm_master_data = " + StrToDQuoteSQL(dbrd["comm_master_data"].ToString()) + ", "
                                + " comm_master_primary = '" + dbrd["comm_master_primary"].ToString() + "', "
                                + " comm_master_order = '" + dbrd["comm_master_order"].ToString() + "'";
                            comm2 = new OdbcCommand(sql, conn2, tran2);
                            comm2.ExecuteNonQuery();
                            comm2 = new OdbcCommand("SELECT LAST_INSERT_ID()", conn2, tran2);
                            int commId = Convert.ToInt32(comm2.ExecuteScalar());

                            sql = "insert into bodycorp_comms set bodycorp_comm_bodycorp_id=" + newBCId + ", bodycorp_comm_comm_id=" + commId;
                            comm2 = new OdbcCommand(sql, conn2, tran2);
                            addBCCommCount += comm2.ExecuteNonQuery();
                        }
                        setLogT(logT, "Add " + addBCCommCount + " BC comms.");
                        #endregion

                        #region add debtors and debtor comms

                        sql = "select * from debtor_master where debtor_master_bodycorp_id = " + currentBCId;
                        comm1 = new OdbcCommand(sql, conn1, tran1);
                        da = new OdbcDataAdapter(comm1);
                        DataTable DebtorDT = new DataTable();
                        da.Fill(DebtorDT);
                        da.Dispose();
                        int addDebtorCount = 0;
                        int addDebtorCommCount = 0;
                        foreach (DataRow dr in DebtorDT.Rows)
                        {
                            string sql_debtor;
                            sql_debtor = "insert debtor_master (";
                            sql_debtor += "debtor_master_code, ";
                            sql_debtor += "debtor_master_name, ";
                            sql_debtor += "debtor_master_bodycorp_id, ";
                            sql_debtor += "debtor_master_type_id, ";
                            sql_debtor += "debtor_master_salutation, ";
                            sql_debtor += "debtor_master_print, ";
                            sql_debtor += "debtor_master_email ";
                            sql_debtor += ") values (";
                            sql_debtor += "'" + dr["debtor_master_code"].ToString() + "', ";          //debtor_master_code
                            sql_debtor += "'" + dr["debtor_master_name"].ToString() + "', ";          //debtor_master_name
                            sql_debtor += newBCId + ", ";   //debtor_master_bodycorp_id
                            sql_debtor += "1, ";                            //debtor_master_type_id
                            sql_debtor += StrToDQuoteSQL(dr["debtor_master_salutation"].ToString()) + ", ";    //debtor_master_salutation
                            sql_debtor += "1, ";                            //debtor_master_print
                            sql_debtor += "0 ";                             //debtor_master_email
                            sql_debtor += ")";

                            comm2 = new OdbcCommand(sql_debtor, conn2, tran2);
                            comm2.ExecuteNonQuery();
                            comm2 = new OdbcCommand("SELECT LAST_INSERT_ID()", conn2, tran2);
                            int debtorId = Convert.ToInt32(comm2.ExecuteScalar());
                            addDebtorCount++;

                            //add debtor comms
                            sql = "select * from debtor_comms where debtor_comm_debtor_id = " + dr["debtor_master_id"].ToString();
                            comm1 = new OdbcCommand(sql, conn1, tran1);
                            da = new OdbcDataAdapter(comm1);
                            DataTable DebtorCommDT = new DataTable();
                            da.Fill(DebtorCommDT);
                            da.Dispose();

                            foreach (DataRow drDebComm in DebtorCommDT.Rows)
                            {
                                sql = "select * from comm_master where comm_master_id = " + drDebComm["debtor_comm_comm_id"].ToString();
                                comm1 = new OdbcCommand(sql, conn1, tran1);
                                dbrd = comm1.ExecuteReader();
                                if (!dbrd.HasRows || !dbrd.Read())
                                {
                                    throw new Exception("Error in select comm_master");
                                }

                                sql = "insert into comm_master set "
                                    + " comm_master_type_id = " + dbrd["comm_master_type_id"].ToString() + ", "
                                    + " comm_master_data = " + StrToDQuoteSQL(dbrd["comm_master_data"].ToString()) + ", "
                                    + " comm_master_primary = '" + dbrd["comm_master_primary"].ToString() + "', "
                                    + " comm_master_order = '" + dbrd["comm_master_order"].ToString() + "'";
                                comm2 = new OdbcCommand(sql, conn2, tran2);
                                comm2.ExecuteNonQuery();
                                comm2 = new OdbcCommand("SELECT LAST_INSERT_ID()", conn2, tran2);
                                int commId = Convert.ToInt32(comm2.ExecuteScalar());

                                sql = "insert into debtor_comms set debtor_comm_debtor_id=" + debtorId + ", debtor_comm_comm_id=" + commId;
                                comm2 = new OdbcCommand(sql, conn2, tran2);
                                addDebtorCommCount += comm2.ExecuteNonQuery();
                            }

                        }
                        setLogT(logT, "Add " + addDebtorCount + " Debtors.");
                        setLogT(logT, "Add " + addDebtorCommCount + " Debtor Comms.");
                        #endregion

                        #region add units

                        sql = "select * from unit_master where unit_master_property_id = " + pmDT.Rows[0]["property_master_id"].ToString();
                        comm1 = new OdbcCommand(sql, conn1, tran1);
                        da = new OdbcDataAdapter(comm1);
                        DataTable UnitDT = new DataTable();
                        da.Fill(UnitDT);
                        da.Dispose();
                        int addUnitCount = 0;
                        foreach (DataRow dr in UnitDT.Rows)
                        {
                            sql = "select * from debtor_master where debtor_master_id = " + dr["unit_master_debtor_id"].ToString();
                            comm1 = new OdbcCommand(sql, conn1, tran1);
                            dbrd = comm1.ExecuteReader();
                            string debtor_code = "";
                            if (dbrd.Read())
                            {
                                debtor_code = dbrd["debtor_master_code"].ToString();
                            }

                            sql = "select * from debtor_master where debtor_master_code = " + StrToDQuoteSQL(debtor_code);
                            comm2 = new OdbcCommand(sql, conn2, tran2);
                            dbrd = comm2.ExecuteReader();
                            string debtor_id = "";
                            if (dbrd.Read())
                            {
                                debtor_id = dbrd["debtor_master_id"].ToString();
                            }

                            string sql_unit;
                            sql_unit = "insert unit_master (";
                            sql_unit += "unit_master_code, unit_master_knowas, unit_master_type_id, unit_master_property_id, ";
                            sql_unit += "unit_master_debtor_id, unit_master_ownership_interest, unit_master_notes, unit_master_begin_date";
                            sql_unit += ") values (";
                            sql_unit += StrToDQuoteSQL(dr["unit_master_code"].ToString()) + ", ";                  //unit_master_code
                            sql_unit += StrToDQuoteSQL(dr["unit_master_knowas"].ToString()) + ", ";                //unit_master_knowas
                            sql_unit += "1, ";                                  //unit_master_type_id
                            sql_unit += propertyId + ",";           //unit_master_property_id
                            sql_unit += debtor_id + ", ";             //unit_master_debtor_id
                            sql_unit += StrToDQuoteSQL(dr["unit_master_ownership_interest"].ToString()) + ", ";    //unit_master_ownership_interest (???)
                            sql_unit += StrToDQuoteSQL(dr["unit_master_notes"].ToString()) + ", ";                 //unit_master_notes
                            sql_unit += DateToSQL(DateTime.Parse(dr["unit_master_begin_date"].ToString())) + ")";	        //unit_master_begin_date

                            comm2 = new OdbcCommand(sql_unit, conn2, tran2);
                            addUnitCount += comm2.ExecuteNonQuery();
                        }
                        setLogT(logT, "Add " + addUnitCount + " Units.");
                        #endregion

                        #region add Ownership

                        sql = "select * from ownerships where ownership_unit_id in ( select unit_master_id  from unit_master where unit_master_property_id= " + pmDT.Rows[0]["property_master_id"].ToString() + " )";
                        comm1 = new OdbcCommand(sql, conn1, tran1);
                        da = new OdbcDataAdapter(comm1);
                        DataTable OwnershipDT = new DataTable();
                        da.Fill(OwnershipDT);
                        da.Dispose();
                        int addOwnershipCount = 0;
                        foreach (DataRow dr in OwnershipDT.Rows)
                        {
                            sql = "select * from unit_master where unit_master_id = " + dr["ownership_unit_id"].ToString();
                            comm1 = new OdbcCommand(sql, conn1, tran1);
                            dbrd = comm1.ExecuteReader();
                            string unit_code = "";
                            if (dbrd.Read())
                            {
                                unit_code = dbrd["unit_master_code"].ToString();
                            }
                            sql = "select * from unit_master where unit_master_code = " + StrToDQuoteSQL(unit_code);
                            comm2 = new OdbcCommand(sql, conn2, tran2);
                            dbrd = comm2.ExecuteReader();
                            string unit_id = "";
                            if (dbrd.Read())
                            {
                                unit_id = dbrd["unit_master_id"].ToString();
                            }

                            sql = "select * from debtor_master where debtor_master_id = " + dr["ownership_debtor_id"].ToString();
                            comm1 = new OdbcCommand(sql, conn1, tran1);
                            dbrd = comm1.ExecuteReader();
                            string debtor_code = "";
                            if (dbrd.Read())
                            {
                                debtor_code = dbrd["debtor_master_code"].ToString();
                            }

                            sql = "select * from debtor_master where debtor_master_code = " + StrToDQuoteSQL(debtor_code);
                            comm2 = new OdbcCommand(sql, conn2, tran2);
                            dbrd = comm2.ExecuteReader();
                            string debtor_id = "";
                            if (dbrd.Read())
                            {
                                debtor_id = dbrd["debtor_master_id"].ToString();
                            }
                            string sql_owner;
                            sql_owner = "insert ownerships (ownership_unit_id, ownership_debtor_id, ownership_start, ownership_end, ownership_notes)";
                            sql_owner += " values ";
                            sql_owner += "(" + unit_id + ", " + debtor_id + ","
                                + "'" + dr["ownership_start"].ToString() + "', "
                                + "'" + dr["ownership_end"].ToString() + "', "
                                + StrToDQuoteSQL(dr["ownership_notes"].ToString()) + ") ";

                            comm2 = new OdbcCommand(sql_owner, conn2, tran2);
                            addOwnershipCount += comm2.ExecuteNonQuery();
                        }
                        setLogT(logT, "Add " + addOwnershipCount + " Ownerships.");
                        #endregion

                        #region add journals
                        sql = "select * from gl_transactions where gl_transaction_bodycorp_id =" + currentBCId;
                        comm1 = new OdbcCommand(sql, conn1, tran1);
                        da = new OdbcDataAdapter(comm1);
                        DataTable journalDT = new DataTable();
                        da.Fill(journalDT);
                        da.Dispose();
                        int addJournalCount = 0;
                        foreach (DataRow dr in journalDT.Rows)
                        {
                            string newJouranlNum = getNextJournalNum();

                            sql = "select * from chart_master where chart_master_id = " + dr["gl_transaction_chart_id"].ToString();
                            comm1 = new OdbcCommand(sql, conn1, tran1);
                            dbrd = comm1.ExecuteReader();
                            string chart_name = "";
                            if (dbrd.Read())
                            {
                                chart_name = dbrd["chart_master_name"].ToString();
                            }
                            sql = "select * from chart_master where chart_master_name = " + StrToDQuoteSQL(chart_name) + " and chart_master_type_id=" + dbrd["chart_master_type_id"].ToString() + "  and chart_master_bank_account =" + dbrd["chart_master_bank_account"].ToString();
                            comm2 = new OdbcCommand(sql, conn2, tran2);
                            dbrd = comm2.ExecuteReader();
                            string chart_id = "";
                            if (dbrd.Read())
                            {
                                chart_id = dbrd["chart_master_id"].ToString();
                            }

                            string unit_id = "NULL";
                            if (!string.IsNullOrEmpty(dr["gl_transaction_unit_id"].ToString()))
                            {
                                sql = "select * from unit_master where unit_master_id = " + dr["gl_transaction_unit_id"].ToString();
                                comm1 = new OdbcCommand(sql, conn1, tran1);
                                dbrd = comm1.ExecuteReader();
                                string unit_code = "";
                                if (dbrd.Read())
                                {
                                    unit_code = dbrd["unit_master_code"].ToString();
                                }
                                sql = "select * from unit_master where unit_master_code = " + StrToDQuoteSQL(unit_code);
                                comm2 = new OdbcCommand(sql, conn2, tran2);
                                dbrd = comm2.ExecuteReader();

                                if (dbrd.Read())
                                {
                                    unit_id = dbrd["unit_master_id"].ToString();
                                }
                            }
                            string sql_gl;
                            sql_gl = " insert gl_transactions (";
                            sql_gl += "gl_transaction_type_id, ";
                            sql_gl += "gl_transaction_ref, ";
                            sql_gl += "gl_transaction_ref_type_id, ";
                            sql_gl += "gl_transaction_oldref, ";
                            sql_gl += "gl_trasaction_oldid, ";
                            sql_gl += "gl_transaction_chart_id, ";
                            sql_gl += "gl_transaction_bodycorp_id, ";
                            //sql_gl += "gl_transaction_creditor_id, ";
                            sql_gl += "gl_transaction_unit_id, ";
                            sql_gl += "gl_transaction_description, ";
                            //$sql_gl .= "gl_transaction_batch_id, ";
                            sql_gl += "gl_transaction_net, ";
                            sql_gl += "gl_transaction_tax, ";
                            sql_gl += "gl_transaction_gross, ";
                            sql_gl += "gl_transaction_date, ";
                            sql_gl += "gl_transaction_rev, ";
                            sql_gl += "gl_transaction_rec, ";
                            //$sql_gl .= "gl_transaction_recbatchid, ";
                            //$sql_gl .= "gl_transaction_reccutoff, ";
                            sql_gl += "gl_transaction_createdate, ";
                            sql_gl += "gl_transaction_user_id ";
                            sql_gl += ") values (";
                            sql_gl += "6, ";                        //gl_transaction_type_id
                            sql_gl += "'" + dr["gl_transaction_ref"].ToString() + "', ";      //gl_transaction_ref
                            sql_gl += "6, ";                        //gl_transaction_ref_type_id
                            sql_gl += "'" + dr["gl_transaction_oldref"].ToString() + "', ";   //gl_transaction_oldref
                            sql_gl += "'" + dr["gl_trasaction_oldid"].ToString() + "', ";     //gl_trasaction_oldid
                            sql_gl += chart_id + ", ";     //gl_transaction_chart_id
                            sql_gl += newBCId + ", ";  //gl_transaction_bodycorp_id
                            sql_gl += unit_id + ", "; //gl_transaction_unit_id
                            sql_gl += StrToDQuoteSQL(dr["gl_transaction_description"].ToString()) + ", ";  //gl_transaction_description
                            sql_gl += dr["gl_transaction_net"].ToString() + ", ";          //gl_transaction_net
                            sql_gl += "0, ";                            //gl_transaction_tax
                            sql_gl += "0, ";                            //gl_transaction_gross
                            sql_gl += DateToSQL(DateTime.Parse(dr["gl_transaction_date"].ToString())) + ", ";         //gl_transaction_date
                            sql_gl += "0, ";                            //gl_transaction_rev
                            sql_gl += "0, ";                            //gl_transaction_rec
                            sql_gl += DateToSQL(DateTime.Parse(dr["gl_transaction_createdate"].ToString())) + ", ";   //gl_transaction_createdate
                            sql_gl += "0 ";                             //gl_transaction_user_id
                            sql_gl += ")";

                            comm2 = new OdbcCommand(sql_gl, conn2, tran2);
                            addJournalCount += comm2.ExecuteNonQuery();
                        }
                        setLogT(logT, "Add " + addJournalCount + " Journals.");

                        #endregion
                        tran1.Commit();
                        tran2.Commit();

                        comm1.Dispose();
                        comm2.Dispose();

                        setLogtColorful(logT, "BC: " + currentBCCode + " | " + currentBCName + " transfer succeed!", Color.Green);
                        setLogT(logT, "");


                    }
                    #endregion
                    catch (Exception ex)
                    {
                        if (conn1.State == ConnectionState.Open)
                        {
                            tran1.Rollback();
                        }
                        if (conn2.State == ConnectionState.Open)
                        {
                            tran2.Rollback();
                        }
                        setLogT(logT, DateTime.Now.ToString() + " " + "Exception: " + ex.Message);
                        MessageBox.Show("Error.");
                        //throw ex;
                    }
                    finally
                    {
                        if (conn1 != null)
                        {
                            conn1.Close();
                        }
                        if (conn2 != null)
                        {
                            conn2.Close();
                        }

                    }
                }

                #endregion

            }

            this.Invoke((MethodInvoker)delegate()
            {
                button4.Enabled = true;
                textBox1.Enabled = true;
                textBox2.Enabled = true;
                checkedListBox1.Enabled = true;
                button7.Enabled = true;
                button6.Enabled = true;
                button5.Enabled = true;
                button2.Enabled = true;
            });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
        }
        

        private void button4_Click(object sender, EventArgs e)
        {
            loadBC();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(((DataRowView)comboBox1.SelectedItem).Row["bodycorp_id"].ToString()))
            {
                return;
            }
            textBox3.Text = ((DataRowView)comboBox1.SelectedItem).Row["bodycorp_code"].ToString();
            textBox4.Text = ((DataRowView)comboBox1.SelectedItem).Row["bodycorp_name"].ToString();
        }

        public void loadBC()
        {

            OdbcConnection conn1 = new OdbcConnection();
            try
            {
                conn1 = new OdbcConnection(textBox1.Text);
                conn1.Open();

                string sql = "SELECT bodycorp_id,bodycorp_code,bodycorp_name from bodycorps ORDER BY bodycorp_code";
                OdbcCommand comm1 = new OdbcCommand(sql, conn1);
                OdbcDataAdapter da = new OdbcDataAdapter(comm1);
                DataTable dt = new DataTable();
                da.Fill(dt);
                da.Dispose();

                dt.Columns.Add("Code");
                foreach (DataRow dr in dt.Rows)
                {
                    dr["Code"] = dr["bodycorp_code"].ToString() + " | " + dr["bodycorp_name"].ToString();
                }
                comboBox1.DataSource = dt;
                comboBox1.DisplayMember = "Code";
                comboBox1.ValueMember = "bodycorp_id";
                ((ListBox)checkedListBox1).DataSource = dt;
                ((ListBox)checkedListBox1).DisplayMember = "Code";
                ((ListBox)checkedListBox1).ValueMember = "bodycorp_id";
            }
            catch (Exception ex)
            {
                setLogT(logT, DateTime.Now.ToString() + " " + "Exception: " + ex.Message);
                MessageBox.Show("Error in comparing.");
                //throw ex;
            }
            finally
            {
                if (conn1 != null)
                {
                    conn1.Close();
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            backgroundWorker2.RunWorkerAsync();
        }



        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox2.Clear();
            logT.Clear();
        }


        private class UtilityTemplateObject_BCnote
        {
            public string Id { get; set; }
            public DateTime Date { get; set; }
            public string Title { get; set; }
            public string Details { get; set; }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, ItemCheckEventArgs e)
        {
            int count = checkedListBox1.CheckedItems.Count + ((e.NewValue == CheckState.Checked) ? 1 : -1);
            label7.Text = count + " BCs selected";
            button2.Text = "Transfer " + count + " BCs From ODBC1 to ODBC2";
            //label7.Text = checkedListBox1.CheckedItems.Count + " BCs selected";
            //button2.Text = "Transfer " + checkedListBox1.CheckedItems.Count + " BCs From ODBC1 to ODBC2";
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //read csv data
                try
                {
                    //uncheck all
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        checkedListBox1.SetItemChecked(i, false);
                    }

                    string[] Lines = File.ReadAllLines(openFileDialog1.FileName);
                    string[] Fields;

                    foreach (string line in Lines)
                    {
                        Fields = line.Split(new char[] { ',' });
                        for (int i = 0; i < checkedListBox1.Items.Count; i++)
                        {
                            foreach (string field in Fields)
                            {
                                if (field == ((DataRowView)checkedListBox1.Items[i]).Row["bodycorp_code"].ToString())
                                {
                                    checkedListBox1.SetItemChecked(i, true);
                                }
                            }
                        }
                    }
                    


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error in reading initial data file: " + ex.ToString());
                    throw;
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
        }

        private void logT_TextChanged(object sender, EventArgs e)
        {
            // set the current caret position to the end
            logT.SelectionStart = logT.Text.Length;
            // scroll it automatically
            logT.ScrollToCaret();
        }




    }
}
