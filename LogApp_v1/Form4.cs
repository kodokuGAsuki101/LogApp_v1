using LogApp_v1.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace LogApp_v1
{
    public partial class Form4 : Form
    {
        private List<PullTicketDataModel> pullTicketData = new List<PullTicketDataModel>();
        DataTable tempdateasper = new DataTable();
        DataTable pullQuanTempTable = new DataTable();
        DataTable asperTable = new DataTable();
        DataTable asperQuanTable = new DataTable();
        DataTable cloneQuanTempTable = new DataTable() ;
        bool asperMarked;
        Form1 f1 = new Form1();

        public DataTable dt = new DataTable();
        public Form4()
        {
            InitializeComponent();
            asperfilterBox.SelectedIndex = 0;
            //f1.tempdatatbl = new DataTable();
            pullQuanTempTable = f1.tempdatatbl;
        }
        private void btnasper_done_Click(object sender, EventArgs e)
        {
            //exits from the "as per advise" view of the application and resets some controls
            this.Close();

        }
        private void Form4_Load(object sender, EventArgs e)
        {
            /*Form1 f1 = new Form1();
            DataTable dt = f1.tempdatatbl;
            asperfilterdatagrid.DataSource = dt;*/
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Opacity == 1)
            {
                timer1.Stop();
            }
            Opacity += .2;
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            if (Opacity <= 0)
            {
                this.Close();
            }
            Opacity -= .2;
        }
        public void TriggerTimerTick()
        {
            timer1_Tick(this, EventArgs.Empty);
        }
        bool isSelectAllButtonAsperAdviceClicked = false;
        //The will select all the data
        private void btnasper_selall_Click(object sender, EventArgs e)
        {
            if (isSelectAllButtonAsperAdviceClicked == false)
            {
                //  asperfilterdatagrid
                foreach (DataGridViewRow row in asperfilterdatagrid.Rows)
                {
                    row.Cells[0].Value = true;
                    isSelectAllButtonAsperAdviceClicked = true;
                }
                btnasper_selall.Text = "UNSELECT ALL";
                btnasper_selall.Width = 120;
                btnasper_mark.Location = new Point(520, 28);
            }
            else 
            {
                //  asperfilterdatagrid
                foreach (DataGridViewRow row in asperfilterdatagrid.Rows)
                {

                    row.Cells[0].Value = false;
                    isSelectAllButtonAsperAdviceClicked = false;
                }
                btnasper_selall.Text = "SELECT ALL";
                btnasper_selall.Width = 92;
                btnasper_mark.Location = new Point(490, 28);
            }   
        }
        private void btnasper_mark_Click(object sender, EventArgs e)
        {
            if (asperfilterdatagrid.Rows.Count == 0) // checks if there is data on the DataGridView
            {
                MessageBox.Show("There is nothing to mark! Please search and select the item to mark it!");
                return;
            }

            for (int a = 0; a < asperfilterdatagrid.Rows.Count - 1; a++) // checks if the user selected any row in the DataGridView
            {
                if ((bool)asperfilterdatagrid.Rows[a].Cells["SELECTED"].Value == true)
                {
                    break;
                }

                if (a == asperfilterdatagrid.Rows.Count - 1)
                {
                    MessageBox.Show("There is no item selected!");
                    return;
                }
            }

            // creates a datatable for as per data if the user has selected for the first time
            if (!asperMarked)
            {
                DataTable asperTable = new DataTable();
                DataTable asperQuanTable = new DataTable();
                asperQuanTable = pullQuanTempTable.Clone();

                DataColumn asperTableCol = asperTable.Columns.Add("TRANS_DATE", typeof(string));
                asperTable.Columns.Add("PROD_DATE", typeof(string));
                asperTable.Columns.Add("PROD_TIME", typeof(string));
                asperTable.Columns.Add("DEL_DATE", typeof(DateTime));
                asperTable.Columns.Add("DEL_TIME", typeof(string));
                asperTable.Columns.Add("JOB_NO", typeof(string));
                asperTable.Columns.Add("FACILITY", typeof(string));
                asperTable.Columns.Add("PARTNUMBER", typeof(string));
                asperTable.Columns.Add("PULL_QTY", typeof(string));
                asperTable.Columns.Add("OPEN_QTY", typeof(string));
                asperTable.Columns.Add("STOCK_QTY", typeof(string));
                asperTable.Columns.Add("QUANTITY_DELIVERED", typeof(string));
                asperTable.Columns.Add("END_BALANCE", typeof(string));
                asperTable.Columns.Add("GO_NUMBER", typeof(string));
                asperTable.Columns.Add("SKU_ASSEMBLY", typeof(string));
                asperTable.Columns.Add("GO_LINE_NUMBER", typeof(string));
                asperTable.Columns.Add("CELL_NUM", typeof(string));
                asperTable.Columns.Add("REMARKS", typeof(string));
                asperTable.Columns.Add("PULL_TICKET_NUMBER", typeof(string));
                asperTable.Columns.Add("LINE", typeof(string));
                asperTable.Columns.Add("VENDOR_REMARKS", typeof(string));
                asperTable.Columns.Add("QTY_DEL", typeof(int));
                asperTable.Columns.Add("ORIGINAL_PULL", typeof(int));
                asperTable.Columns.Add("BACKLOGTYPE", typeof(string));
            }

            // creates a datatable for the partnumber and line number
            DataTable asperPullLine = new DataTable();
            DataColumn asperPullLineCol = asperPullLine.Columns.Add("PULL_TICKET_NUMBER", typeof(string));
            asperPullLine.Columns.Add("LINE", typeof(string));

            for (int cr = 0; cr < asperfilterdatagrid.Rows.Count; cr++)
            {
                if ((bool)asperfilterdatagrid.Rows[cr].Cells["SELECTED"].Value)
                {
                    // stores the selected row in a datatable
                    DataRow pq = asperTable.NewRow();
                    pq["TRANS_DATE"] = DateTime.Now.ToString("yyyy/MM/dd");
                    pq["PROD_DATE"] = asperfilterdatagrid.Rows[cr].Cells["PROD_DATE"].Value;
                    pq["PROD_TIME"] = asperfilterdatagrid.Rows[cr].Cells["PROD_TIME"].Value;
                    pq["DEL_DATE"] = asperfilterdatagrid.Rows[cr].Cells["DEL_DATE"].Value;
                    pq["DEL_TIME"] = asperfilterdatagrid.Rows[cr].Cells["DEL_TIME"].Value;
                    pq["JOB_NO"] = asperfilterdatagrid.Rows[cr].Cells["JOB_NO"].Value;
                    pq["FACILITY"] = asperfilterdatagrid.Rows[cr].Cells["FACILITY"].Value;
                    pq["PARTNUMBER"] = asperfilterdatagrid.Rows[cr].Cells["PARTNUMBER"].Value;
                    pq["PULL_QTY"] = asperfilterdatagrid.Rows[cr].Cells["PULL_QTY"].Value;
                    pq["OPEN_QTY"] = "";
                    pq["STOCK_QTY"] = "";
                    pq["QUANTITY_DELIVERED"] = "";
                    pq["END_BALANCE"] = "";
                    pq["GO_NUMBER"] = "";
                    pq["SKU_ASSEMBLY"] = asperfilterdatagrid.Rows[cr].Cells["SKU_ASSEMBLY"].Value;
                    pq["GO_LINE_NUMBER"] = "";
                    pq["CELL_NUM"] = asperfilterdatagrid.Rows[cr].Cells["CELLNUMBER"].Value;
                    pq["REMARKS"] = asperfilterdatagrid.Rows[cr].Cells["REMARKS"].Value;
                    pq["PULL_TICKET_NUMBER"] = asperfilterdatagrid.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value;
                    pq["LINE"] = asperfilterdatagrid.Rows[cr].Cells["LINE"].Value;
                    pq["VENDOR_REMARKS"] = asperfilterdatagrid.Rows[cr].Cells["VENDOR_REMARKS"].Value;
                    pq["QTY_DEL"] = asperfilterdatagrid.Rows[cr].Cells["QTY_DEL"].Value;
                    pq["ORIGINAL_PULL"] = asperfilterdatagrid.Rows[cr].Cells["ORIGINAL_PULL"].Value;
                    pq["BACKLOGTYPE"] = "AS PER ADVISE";
                    asperTable.Rows.Add(pq);
                    asperTable.AcceptChanges();

                    // stores the selected row in a datatable
                    DataRow pq2 = asperQuanTable.NewRow();
                    pq2["PROD_DATE"] = asperfilterdatagrid.Rows[cr].Cells["PROD_DATE"].Value;
                    pq2["PROD_TIME"] = asperfilterdatagrid.Rows[cr].Cells["PROD_TIME"].Value;
                    pq2["DEL_DATE"] = asperfilterdatagrid.Rows[cr].Cells["DEL_DATE"].Value;
                    pq2["DEL_TIME"] = asperfilterdatagrid.Rows[cr].Cells["DEL_TIME"].Value;
                    pq2["JOB_NO"] = asperfilterdatagrid.Rows[cr].Cells["JOB_NO"].Value;
                    pq2["FACILITY"] = asperfilterdatagrid.Rows[cr].Cells["FACILITY"].Value;
                    pq2["PARTNUMBER"] = asperfilterdatagrid.Rows[cr].Cells["PARTNUMBER"].Value;
                    pq2["PULL_QTY"] = asperfilterdatagrid.Rows[cr].Cells["PULL_QTY"].Value;
                    pq2["VENDOR_NAME"] = asperfilterdatagrid.Rows[cr].Cells["VENDOR_NAME"].Value;
                    pq2["SKU_ASSEMBLY"] = asperfilterdatagrid.Rows[cr].Cells["SKU_ASSEMBLY"].Value;
                    pq2["CELLNUMBER"] = asperfilterdatagrid.Rows[cr].Cells["CELLNUMBER"].Value;
                    pq2["REMARKS"] = asperfilterdatagrid.Rows[cr].Cells["REMARKS"].Value;
                    pq2["PULL_TICKET_NUMBER"] = asperfilterdatagrid.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value;
                    pq2["LINE"] = asperfilterdatagrid.Rows[cr].Cells["LINE"].Value;
                    pq2["FILEUPLOADDATE"] = asperfilterdatagrid.Rows[cr].Cells["FILEUPLOADDATE"].Value;
                    pq2["VENDOR_REMARKS"] = asperfilterdatagrid.Rows[cr].Cells["VENDOR_REMARKS"].Value;
                    pq2["ACKNOWLEDGMENT_DATE"] = asperfilterdatagrid.Rows[cr].Cells["ACKNOWLEDGMENT_DATE"].Value;
                    pq2["ACKNOWLEDGMENT_REMARKS"] = asperfilterdatagrid.Rows[cr].Cells["ACKNOWLEDGMENT_REMARKS"].Value;
                    pq2["COMMIT_QTY"] = asperfilterdatagrid.Rows[cr].Cells["COMMIT_QTY"].Value;
                    pq2["COMMIT_DATE"] = asperfilterdatagrid.Rows[cr].Cells["COMMIT_DATE"].Value;
                    pq2["BUYER_REMARKS_FOR_VENDOR"] = asperfilterdatagrid.Rows[cr].Cells["BUYER_REMARKS_FOR_VENDOR"].Value;
                    pq2["QTY_DELIVERED"] = asperfilterdatagrid.Rows[cr].Cells["QTY_DELIVERED"].Value;
                    pq2["DL_VARIENCE"] = asperfilterdatagrid.Rows[cr].Cells["DL_VARIENCE"].Value;
                    pq2["HITMISS"] = asperfilterdatagrid.Rows[cr].Cells["HITMISS"].Value;
                    pq2["STATUS"] = asperfilterdatagrid.Rows[cr].Cells["STATUS"].Value;
                    pq2["PULLTYPE"] = "NEW";
                    pq2["QTY_DEL"] = asperfilterdatagrid.Rows[cr].Cells["QTY_DEL"].Value;
                    pq2["ORIGINAL_PULL"] = asperfilterdatagrid.Rows[cr].Cells["ORIGINAL_PULL"].Value;
                    asperQuanTable.Rows.Add(pq2);
                    asperQuanTable.AcceptChanges();

                    // stores the selected row in a datatable
                    DataRow pq3 = asperPullLine.NewRow();
                    pq3["PULL_TICKET_NUMBER"] = asperfilterdatagrid.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value;
                    pq3["LINE"] = asperfilterdatagrid.Rows[cr].Cells["LINE"].Value;
                    asperPullLine.Rows.Add(pq3);
                    asperPullLine.AcceptChanges();
                }
            }

            //deletes the selected rows in the pullticket datatable and in the backup pullticket datatable
            for (int cr = 0; cr < asperPullLine.Rows.Count; cr++)
            {
                string backexp = "PULL_TICKET_NUMBER = '" + asperPullLine.Rows[cr]["PULL_TICKET_NUMBER"] + "' AND LINE = '" + asperPullLine.Rows[cr]["LINE"] + "'";
                DataRow[] backrow;
                backrow = pullQuanTempTable.Select(backexp);

                foreach (DataRow backrow2 in backrow)
                {
                    backrow2.Delete();
                }

                string backexp2 = "PULL_TICKET_NUMBER = '" + asperPullLine.Rows[cr]["PULL_TICKET_NUMBER"] + "' AND LINE = '" + asperPullLine.Rows[cr]["LINE"] + "'";
                DataRow[] backrow3;
                backrow3 = cloneQuanTempTable.Select(backexp2);

                foreach (DataRow backrow4 in backrow3)
                {
                    backrow4.Delete();
                }

                cloneQuanTempTable.AcceptChanges();
                pullQuanTempTable.AcceptChanges();
            }

            asperfilterdatagrid.DataSource = null; // clears the datagridview
            asperfilterdatagrid.DataSource = pullQuanTempTable; // sets the datagridview data source
            // hides unnecessary columns in the datagridview
            asperfilterdatagrid.Columns["PROD_DATE"].Visible = false;
            asperfilterdatagrid.Columns["PROD_TIME"].Visible = false;
            asperfilterdatagrid.Columns["JOB_NO"].Visible = false;
            asperfilterdatagrid.Columns["VENDOR_NAME"].Visible = false;
            asperfilterdatagrid.Columns["SKU_ASSEMBLY"].Visible = false;
            asperfilterdatagrid.Columns["CELLNUMBER"].Visible = false;

            asperMarked = true;


        }

        //This will hold the value of the row that current selected
        private bool isRowChecked = false;

        private void asperfilterdatagrid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (asperfilterdatagrid.SelectedCells.Count > 0) {

                int selectedrowindex = asperfilterdatagrid.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = asperfilterdatagrid.Rows[selectedrowindex];
                string cellValueSelected = Convert.ToString(selectedRow.Cells["Selected"].Value);

                if (cellValueSelected.Equals("True"))
                {
                    isRowChecked = true;
                }
                else
                {
                    isRowChecked = false;
                }
                String prodDate = selectedRow.Cells["PROD_DATE"].Value.ToString();
                String prodTime = selectedRow.Cells["PROD_TIME"].Value.ToString();
                String delDate = selectedRow.Cells["DEL_DATE"].Value.ToString();
                String delTime = selectedRow.Cells["DEL_TIME"].Value.ToString();
                String jobNo = selectedRow.Cells["JOB_NO"].Value.ToString();
                String facility = selectedRow.Cells["FACILITY"].Value.ToString();
                String partNo = selectedRow.Cells["PARTNUMBER"].Value.ToString();
                String pull_qty = selectedRow.Cells["PULL_QTY"].Value.ToString();
                String vendor_name = selectedRow.Cells["VENDOR_NAME"].Value.ToString();
                String sku_assembly = selectedRow.Cells["SKU_ASSEMBLY"].Value.ToString();
                String cell_no = selectedRow.Cells["CELLNUMBER"].Value.ToString();
                String remarks = selectedRow.Cells["REMARKS"].Value.ToString();
                String pull_ticket_no = selectedRow.Cells["PULL_TICKET_NUMBER"].Value.ToString();
                String line = selectedRow.Cells["LINE"].Value.ToString();
                String fileUploadDate = selectedRow.Cells["FILEUPLOADDATE"].Value.ToString();
                String vendor_acknowledgment = selectedRow.Cells["VENDOR_REMARKS"].Value.ToString();
                String acknowledgement_date = selectedRow.Cells["ACKNOWLEDGMENT_DATE"].Value.ToString();
                String acknowledgement_remarks_for_vendor = selectedRow.Cells["ACKNOWLEDGMENT_REMARKS"].Value.ToString();
                String buyer_remarks_for_vendor = selectedRow.Cells["BUYER_REMARKS_FOR_VENDOR"].Value.ToString();
                String qty_delivered = selectedRow.Cells["QTY_DELIVERED"].Value.ToString();
                String dl_varience = selectedRow.Cells["DL_VARIENCE"].Value.ToString();
                String hitmiss = selectedRow.Cells["HITMISS"].Value.ToString();
                String status = selectedRow.Cells["STATUS"].Value.ToString();
                String pulltype = selectedRow.Cells["PULLTYPE"].Value.ToString();
                String qty_del = selectedRow.Cells["QTY_DEL"].Value.ToString();
                String original_pull = selectedRow.Cells["ORIGINAL_PULL"].Value.ToString();

                if (isRowChecked == false) //if rowCheck is false
                {
                    //checkedBacklogData.RemoveAll(r => r.isChecked == false);
                    //This will check if the backload datamodel if it was unchecked
                    //This will check if the current selected row have
                    //the same del_date, partnumber, balance and line if 
                    //they are have the same value it will not be added.
                    int checkedCountValue = 0;
                    foreach (PullTicketDataModel data in pullTicketData)
                    {
                        var matches = pullTicketData.Where(p => p.delDate == delDate &&
                          p.jobNo == jobNo && p.facility == facility && p.line == line &&
                          p.hitmiss == hitmiss && p.cellNo == cell_no);

                        checkedCountValue = matches.Count();
                    }
                    if (checkedCountValue == 0)
                    {
                        pullTicketData.Add(new PullTicketDataModel(
                            prodDate,
                            prodTime,
                            delDate,
                            delTime,
                            jobNo,
                            facility,
                            partNo,
                            pull_qty,
                            vendor_name,
                            sku_assembly,
                            cell_no,
                            remarks,
                            pull_ticket_no,
                            line,
                            fileUploadDate,
                            vendor_acknowledgment,
                            acknowledgement_date,
                            acknowledgement_remarks_for_vendor,
                            qty_delivered,
                            dl_varience,
                            hitmiss,
                            status
                            ));
                        selectedRow.Cells["SELECTED"].Value = true;
                    }
                }
                else 
                {
                    int checkedCountValue = 0;
                    foreach (PullTicketDataModel data in pullTicketData)
                    {
                        var matches = pullTicketData.Where(p => p.delDate == delDate &&
                      p.jobNo == jobNo && p.facility == facility && p.line == line &&
                      p.hitmiss == hitmiss && p.cellNo == cell_no);

                        checkedCountValue = matches.Count();
                    }

                    if (checkedCountValue > 0)
                    {
                        pullTicketData.Remove(pullTicketData.Find(p => p.delDate == delDate &&
                      p.jobNo == jobNo && p.facility == facility && p.line == line &&
                      p.hitmiss == hitmiss && p.cellNo == cell_no));

                        selectedRow.Cells["SELECTED"].Value = false;
                    }
                }
            }
        }
        private void asperfilterTxt_TextChanged(object sender, EventArgs e)
        {
            pullQuanTempTable = f1.tempdatatbl;
            if (asperfilterTxt.Text == "")
            {
                // Executes if the textbox is empty
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = pullQuanTempTable;

                asperfilterdatagrid.Columns["PROD_DATE"].Visible = false;
                asperfilterdatagrid.Columns["PROD_TIME"].Visible = false;
                asperfilterdatagrid.Columns["JOB_NO"].Visible = false;
                asperfilterdatagrid.Columns["VENDOR_NAME"].Visible = false;
                asperfilterdatagrid.Columns["SKU_ASSEMBLY"].Visible = false;
                asperfilterdatagrid.Columns["CELLNUMBER"].Visible = false;
                return;
            }

            if (asperfilterBox.SelectedIndex == -1)
            {
                // Executes if the combobox's text is empty
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = pullQuanTempTable;

                asperfilterdatagrid.Columns["PROD_DATE"].Visible = false;
                asperfilterdatagrid.Columns["PROD_TIME"].Visible = false;
                asperfilterdatagrid.Columns["JOB_NO"].Visible = false;
                asperfilterdatagrid.Columns["VENDOR_NAME"].Visible = false;
                asperfilterdatagrid.Columns["SKU_ASSEMBLY"].Visible = false;
                asperfilterdatagrid.Columns["CELLNUMBER"].Visible = false;
                return;
            }

            DataTable filterAsPer = new DataTable();
            pullQuanTempTable.CaseSensitive = false;
            DataView negaView = new DataView(pullQuanTempTable);

            // Filters the datatable according to the selected value in combobox and stores them in a datatable. Displays the result in the datagridview
            if (asperfilterBox.Text == "PARTNUMBER")
            {
                negaView.RowFilter = "PARTNUMBER LIKE '%" + asperfilterTxt.Text + "%'";
                filterAsPer = negaView.ToTable();
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "FACILITY")
            {
                negaView.RowFilter = "FACILITY LIKE '%" + asperfilterTxt.Text + "%'";
                filterAsPer = negaView.ToTable();
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "PULLTICKET")
            {
                negaView.RowFilter = "PULL_TICKET_NUMBER LIKE '%" + asperfilterTxt.Text + "%'";
                filterAsPer = negaView.ToTable();
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = filterAsPer;
            }
            else
            {
                negaView.RowFilter = "REMARKS LIKE '%" + asperfilterTxt.Text + "%'";
                filterAsPer = negaView.ToTable();
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = filterAsPer;
            }

            // Hides unnecessary columns
            asperfilterdatagrid.Columns["PROD_DATE"].Visible = false;
            asperfilterdatagrid.Columns["PROD_TIME"].Visible = false;
            asperfilterdatagrid.Columns["JOB_NO"].Visible = false;
            asperfilterdatagrid.Columns["VENDOR_NAME"].Visible = false;
            asperfilterdatagrid.Columns["SKU_ASSEMBLY"].Visible = false;
            asperfilterdatagrid.Columns["CELLNUMBER"].Visible = false;
        }

        private void asperfilterBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            pullQuanTempTable = f1.tempdatatbl;
            if (asperfilterBox.SelectedIndex == -1)
            {
                return; // if the combobox's text is empty
            }
            else
            {
                asperfilterBox.Enabled = true; // enables the textbox if a value is selected
            }

            if (string.IsNullOrEmpty(asperfilterTxt.Text))
            {
                // binds the datatable to a binding source and sets it as a data source for the datagridview
                BindingSource consoBinding = new BindingSource();
                consoBinding.DataSource = pullQuanTempTable;
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = consoBinding;
                // hides unnecessary columns
                asperfilterdatagrid.Columns["PROD_DATE"].Visible = false;
                asperfilterdatagrid.Columns["PROD_TIME"].Visible = false;
                asperfilterdatagrid.Columns["JOB_NO"].Visible = false;
                asperfilterdatagrid.Columns["VENDOR_NAME"].Visible = false;
                asperfilterdatagrid.Columns["SKU_ASSEMBLY"].Visible = false;
                asperfilterdatagrid.Columns["CELLNUMBER"].Visible = false;
                return;
            }

            // creates a datatable for filtering
            DataTable filterAsPer = new DataTable();
            pullQuanTempTable.CaseSensitive = false;
            DataView negaView = new DataView(pullQuanTempTable);

            BindingSource consoBinding2 = new BindingSource();
            consoBinding2.DataSource = negaView;

            // filters the datatable view based on the selected value in the combobox, stores it in a datatable and displays them in the datagridview
            negaView.RowFilter = "COLUMN_NAME = '" + asperfilterBox.SelectedItem + "'";
            asperfilterdatagrid.DataSource = consoBinding2;

            if (asperfilterBox.Text == "PARTNUMBER")
            {
                negaView.RowFilter = $"PARTNUMBER LIKE '%{asperfilterTxt.Text}%'";
                filterAsPer = negaView.ToTable();
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "FACILITY")
            {
                negaView.RowFilter = $"FACILITY LIKE '%{asperfilterTxt.Text}%'";
                filterAsPer = negaView.ToTable();
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "PULLTICKET")
            {
                negaView.RowFilter = $"PULL_TICKET_NUMBER LIKE '%{asperfilterTxt.Text}%'";
                filterAsPer = negaView.ToTable();
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = filterAsPer;
            }
            else
            {
                negaView.RowFilter = $"REMARKS LIKE '%{asperfilterTxt.Text}%'";
                filterAsPer = negaView.ToTable();
                asperfilterdatagrid.DataSource = null;
                asperfilterdatagrid.DataSource = filterAsPer;
            }

            // hides unnecessary columns
            asperfilterdatagrid.Columns["PROD_DATE"].Visible = false;
            asperfilterdatagrid.Columns["PROD_TIME"].Visible = false;
            asperfilterdatagrid.Columns["JOB_NO"].Visible = false;
            asperfilterdatagrid.Columns["VENDOR_NAME"].Visible = false;
            asperfilterdatagrid.Columns["SKU_ASSEMBLY"].Visible = false;
            asperfilterdatagrid.Columns["CELLNUMBER"].Visible = false;


        }

    }
}

