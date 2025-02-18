using System;
using System.Data;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using System.Globalization;
using DateTime = System.DateTime;
using System.Linq;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using LogApp_v1.Models;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Syncfusion.XlsIO;
using System.Data.Common;
namespace LogApp_v1
{
    public partial class Form1 : Form
    {
        private List<PullTicketDataModel> pullTicketData = new List<PullTicketDataModel>();
        public DataTable tempdatatbl = new DataTable(),
            cxmrdtbl = new DataTable(), aimrdtbl = new DataTable(),
            axmrtbl = new DataTable(), asperTable = new DataTable(),
            pullQuantityTable = new DataTable(), forFilter = new DataTable(),
            pullQuanTempTable = new DataTable(), excelData2 = new DataTable(),
            pullrecordTable = new DataTable(), pullbackTable = new DataTable(),
            POTbl = new DataTable(), inventoryTable = new DataTable(),
            poTable = new DataTable(), consolidatedTable = new DataTable(),
            poTableCopy = new DataTable(), totalPullSM = new DataTable(),
            colorPOsm = new DataTable(), excelData = new DataTable(),
            unpostedTbl = new DataTable(), kanbanTbl = new DataTable(),
            invSMorNonTbl = new DataTable(), excelData3 = new DataTable(),
            cloneQuanTempTable = new DataTable(), pullDeleteTable = new DataTable(),
            conso_delTable = new DataTable(), pullRevTable = new DataTable(),
            manualDrTable = new DataTable(), advBacklogTbl = new DataTable(),
            amarokTbl = new DataTable(), backNewDelDate = new DataTable(),
            historyBack = new DataTable(), backTempotbl = new DataTable(),
            invTblCopy = new DataTable(), conso_final_formatTable = new DataTable(),
            selAdvBackTbl = new DataTable(), asperQuanTable = new DataTable(),
            backlogTempoTable = new DataTable(), fgSM = new DataTable();
        public string UndelFileName,
            Filename, pullFileName;
        // Excel application instance for interacting with Excel files
        public Excel.Application xlApp = new Excel.Application();

        // Workbook object to represent the currently opened Excel workbook
        public Excel.Workbook xlWorkbook;

        // Worksheet object to represent the currently active worksheet in the workbook
        public Excel.Worksheet xlWorksheet;

        // Another instance of Excel application for a different purpose or operation
        public Excel.Application oExcel2 = new Excel.Application();

        // Index to track the current row in Excel operations
        public int rowIndex;
        // String to hold the current date and time formatted for logging or display
        public string date_time = DateTime.Now.ToString("yyyy/MM/dd HH:mm:00");

        // String to hold the current date and time formatted for use in file names
        public string date_time_file = DateTime.Now.ToString("yyyy_MM_dd (hhmmtt)");

        // Binding source for data binding operations, typically used in data-driven applications
        BindingSource consoBinding = new BindingSource();
        DateTime datePick1, datePick2,
            datePick3, datePick4;
        //bool variables for importing files
        public bool importPull, importInv,
            importCxmr, importPO,
            calculateBtn, uploadOne,
            uploadTwo, errorDelDate,
            pullticketfilter, zoomClick;
        //bool for checking datas
        public bool backlogCheck, DelChanged, backlogDel,
            asperclick, genClick, facFil,
            date2sel, date1sel, asperMarked;
        public object pull_request_trans = 0;
        // string that holds oracle connection query
        private string connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.50.40)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User ID=mec;Password=mec2024;"; //oracle host source
        // variable for oracle connection used to control the database
        private OracleConnection con = new OracleConnection();
        public Form1()
        {
            InitializeComponent();

            pulltktgrid.Visible = false;// datagrid hide by default
            btnfilter.Visible = false; //filter button hide by default
            backlogDel = false; //backlog delete is false by default
            deliverydatedocker1.MinDate = DateTime.Now; // minimum date for deliverydatedocker1
            DelChanged = true; //delivery date changed is true by default
            btnaxmr340.Visible = false; //axmr340 button hide by default
            refesherOrb.Visible = false; //refresh button hide by default
            ptkFileName.Visible = false; //pull ticket file name label hide by default
            importPull = false; importCxmr = false; importInv = false; importPO = false; calculateBtn = false; //bool for import files are false by default
            connectdata(); //calling this metohd to open the sql connection
            genClick = false;
            asperfilterBox.SelectedIndex = 0; //index for asperfilterbox is 0 by default
            btncmpl.Visible = false;
            dataGridView4.Visible = false; //datagridview4 is hidden by default
            //oExcel2 = xlApp;
        }
        //method to connect sql data
        private void connectdata()
        {
            con.ConnectionString = connectionString; //declaring connection from host
            con.Open(); //to open database connection
        }
        //--> start --> button custom style onhover
        private void btnpull_tkt_MouseEnter(object sender, EventArgs e)
        {
            btnpull_tkt.ForeColor = Color.White;
        }
        private void btnpull_tkt_MouseLeave(object sender, EventArgs e)
        {
            btnpull_tkt.ForeColor = Color.White;
        }
        private void cxmr620_MouseEnter(object sender, EventArgs e)
        {
            btncxmr620.ForeColor = Color.White;
        }
        private void cxmr620_MouseLeave(object sender, EventArgs e)
        {
            btncxmr620.ForeColor = Color.White;
        }
        private void aimr407_MouseEnter(object sender, EventArgs e)
        {
            btnaimr407.ForeColor = Color.White;
        }
        private void aimr407_MouseLeave(object sender, EventArgs e)
        {
            btnaimr407.ForeColor = Color.White;
        }
        private void axmr432_MouseEnter(object sender, EventArgs e)
        {
            btnaxmr432.ForeColor = Color.White;
        }
        private void axmr432_MouseLeave(object sender, EventArgs e)
        {
            btnaxmr432.ForeColor = Color.White;
        }
        private void axmr340_MouseEnter(object sender, EventArgs e)
        {
            btnaxmr340.ForeColor = Color.White;
        }
        private void axmr340_MouseLeave(object sender, EventArgs e)
        {
            btnaxmr340.ForeColor = Color.White;
        }
        private void cmpl_MouseEnter(object sender, EventArgs e)
        {
            btncmpl.ForeColor = Color.White;
        }
        private void cmpl_MouseLeave(object sender, EventArgs e)
        {
            btncmpl.ForeColor = Color.White;
        }
        private void download_MouseEnter_1(object sender, EventArgs e)
        {
            btndownload.ForeColor = Color.White;
        }
        private void download_MouseLeave(object sender, EventArgs e)
        {
            btndownload.ForeColor = Color.White;
        }
        //button custom style onhover -->end<---///
        private int imageNumber = 1; //index for carousel
        //carousel slideshow
        private void LoadNextImage()
        {
            if (imageNumber == 5)
            {
                imageNumber = 1;
            }
            slider.Image = Properties.Resources.ResourceManager.GetObject($"mecslide{imageNumber}") as Image;
            imageNumber++;
        }
        //timer for carousel
        private void timer1_Tick(object sender, EventArgs e)
        {
            LoadNextImage();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            DoubleBuffered = true;
            EnableDoubleBuffering();
            ExtendedMethods.DoubleBuffered(pulltktgrid, true);
            ExtendedMethods.DoubleBuffered(DataGridView3, true);
            ExtendedMethods.DoubleBuffered(dataGridView4, true);
        }
        public void EnableDoubleBuffering()
        {
            // Set the value of the double-buffering style bits to true.
            this.SetStyle(ControlStyles.DoubleBuffer |
                           ControlStyles.UserPaint |
                           ControlStyles.AllPaintingInWmPaint,
                           true);
            this.UpdateStyles();
        }
        private void asperfilterBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //pullQuanTempTable = tempdatatbl;
            if (asperfilterBox.SelectedIndex == 0)
            {
                return; // if the combobox's text is empty
            }
            else
            {
                asperfilterBox.Enabled = true; // enables the textbox if a value is selected
            }
            if (asperfilterBox.Text == "ALL")
            {
                asperfilterTxt.Enabled = false;
                asperfilterTxt.Text = null;
                BindingSource consoBinding = new BindingSource();
                consoBinding.DataSource = pullQuanTempTable;
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = consoBinding;
                // hides unnecessary columns
                DataGridView3.Columns["PROD_DATE"].Visible = false;
                DataGridView3.Columns["PROD_TIME"].Visible = false;
                DataGridView3.Columns["JOB_NO"].Visible = false;
                DataGridView3.Columns["VENDOR_NAME"].Visible = false;
                DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
                DataGridView3.Columns["CELLNUMBER"].Visible = false;
                return;
            }
            if (string.IsNullOrEmpty(asperfilterTxt.Text))
            {
                // binds the datatable to a binding source and sets it as a data source for the datagridview
                BindingSource consoBinding = new BindingSource();
                consoBinding.DataSource = pullQuanTempTable;
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = consoBinding;
                // hides unnecessary columns
                DataGridView3.Columns["PROD_DATE"].Visible = false;
                DataGridView3.Columns["PROD_TIME"].Visible = false;
                DataGridView3.Columns["JOB_NO"].Visible = false;
                DataGridView3.Columns["VENDOR_NAME"].Visible = false;
                DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
                DataGridView3.Columns["CELLNUMBER"].Visible = false;
                return;
            }
            // creates a datatable for filtering
            DataTable filterAsPer = new DataTable();
            pullQuanTempTable.CaseSensitive = false;
            DataView negaView = new DataView(pullQuanTempTable);
            BindingSource consoBinding2 = new BindingSource();
            consoBinding2.DataSource = negaView;
            // filters the datatable view based on the selected value in the combobox, stores it in a datatable and displays them in the datagridview
            if (asperfilterBox.Text == "PARTNUMBER")
            {
                asperfilterTxt.Enabled = true;
                negaView.RowFilter = $"PARTNUMBER LIKE '%{asperfilterTxt.Text}%'";
                filterAsPer = negaView.ToTable();
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "FACILITY")
            {
                asperfilterTxt.Enabled = true;
                negaView.RowFilter = $"FACILITY LIKE '%{asperfilterTxt.Text}%'";
                filterAsPer = negaView.ToTable();
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "PULLTICKET")
            {
                asperfilterTxt.Enabled = true;
                negaView.RowFilter = $"PULL_TICKET_NUMBER LIKE '%{asperfilterTxt.Text}%'";
                filterAsPer = negaView.ToTable();
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "REMARKS")
            {
                asperfilterTxt.Enabled = true;
                negaView.RowFilter = $"REMARKS LIKE '%{asperfilterTxt.Text}%'";
                filterAsPer = negaView.ToTable();
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = filterAsPer;
            }
            else
            {
                asperfilterTxt.Enabled = false;
                asperfilterTxt.Text = null;
                BindingSource consoBinding = new BindingSource();
                consoBinding.DataSource = pullQuanTempTable;
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = consoBinding;
                // hides unnecessary columns
                DataGridView3.Columns["PROD_DATE"].Visible = false;
                DataGridView3.Columns["PROD_TIME"].Visible = false;
                DataGridView3.Columns["JOB_NO"].Visible = false;
                DataGridView3.Columns["VENDOR_NAME"].Visible = false;
                DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
                DataGridView3.Columns["CELLNUMBER"].Visible = false;
                return;
            }
            // hides unnecessary columns
            DataGridView3.Columns["PROD_DATE"].Visible = false;
            DataGridView3.Columns["PROD_TIME"].Visible = false;
            DataGridView3.Columns["JOB_NO"].Visible = false;
            DataGridView3.Columns["VENDOR_NAME"].Visible = false;
            DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
            DataGridView3.Columns["CELLNUMBER"].Visible = false;
        }
        private void asperfilterTxt_TextChanged(object sender, EventArgs e)
        {
            //pullQuanTempTable = tempdatatbl;
            if (asperfilterTxt.Text == "")
            {
                // Executes if the textbox is empty
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = pullQuanTempTable;
                DataGridView3.Columns["PROD_DATE"].Visible = false;
                DataGridView3.Columns["PROD_TIME"].Visible = false;
                DataGridView3.Columns["JOB_NO"].Visible = false;
                DataGridView3.Columns["VENDOR_NAME"].Visible = false;
                DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
                DataGridView3.Columns["CELLNUMBER"].Visible = false;
                return;
            }

            if (asperfilterBox.SelectedIndex == 0)
            {
                // Executes if the combobox's text is empty
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = pullQuanTempTable;
                DataGridView3.Columns["PROD_DATE"].Visible = false;
                DataGridView3.Columns["PROD_TIME"].Visible = false;
                DataGridView3.Columns["JOB_NO"].Visible = false;
                DataGridView3.Columns["VENDOR_NAME"].Visible = false;
                DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
                DataGridView3.Columns["CELLNUMBER"].Visible = false;
                asperfilterTxt.Enabled = false;
                return;
            }
            DataTable filterAsPer = new DataTable();
            pullQuanTempTable.CaseSensitive = false;
            DataView negaView = new DataView(pullQuanTempTable);
            // Filters the datatable according to the selected value in combobox and stores them in a datatable. Displays the result in the datagridview
            if (asperfilterBox.Text == "PARTNUMBER")
            {
                asperfilterTxt.Enabled = true;
                negaView.RowFilter = "PARTNUMBER LIKE '%" + asperfilterTxt.Text + "%'";
                filterAsPer = negaView.ToTable();
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "FACILITY")
            {
                asperfilterTxt.Enabled = true;
                negaView.RowFilter = "FACILITY LIKE '%" + asperfilterTxt.Text + "%'";
                filterAsPer = negaView.ToTable();
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "PULLTICKET")
            {
                asperfilterTxt.Enabled = true;
                negaView.RowFilter = "PULL_TICKET_NUMBER LIKE '%" + asperfilterTxt.Text + "%'";
                filterAsPer = negaView.ToTable();
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = filterAsPer;
            }
            else if (asperfilterBox.Text == "REMARKS")
            {
                asperfilterTxt.Enabled = true;
                negaView.RowFilter = "REMARKS LIKE '%" + asperfilterTxt.Text + "%'";
                filterAsPer = negaView.ToTable();
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = filterAsPer;
            }
            else
            {
                asperfilterTxt.Enabled = false;
                asperfilterTxt.Text = null;
                BindingSource consoBinding = new BindingSource();
                consoBinding.DataSource = pullQuanTempTable;
                DataGridView3.DataSource = null;
                DataGridView3.DataSource = consoBinding;
                // hides unnecessary columns
                DataGridView3.Columns["PROD_DATE"].Visible = false;
                DataGridView3.Columns["PROD_TIME"].Visible = false;
                DataGridView3.Columns["JOB_NO"].Visible = false;
                DataGridView3.Columns["VENDOR_NAME"].Visible = false;
                DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
                DataGridView3.Columns["CELLNUMBER"].Visible = false;
                return;
            }
            // Hides unnecessary columns
            DataGridView3.Columns["PROD_DATE"].Visible = false;
            DataGridView3.Columns["PROD_TIME"].Visible = false;
            DataGridView3.Columns["JOB_NO"].Visible = false;
            DataGridView3.Columns["VENDOR_NAME"].Visible = false;
            DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
            DataGridView3.Columns["CELLNUMBER"].Visible = false;
        }
        bool isSelectAllButtonAsperAdviceClicked = false;
        private void btnasper_selall_Click(object sender, EventArgs e)
        {
            if (isSelectAllButtonAsperAdviceClicked == false)
            {
                //  asperfilterdatagrid
                foreach (DataGridViewRow row in DataGridView3.Rows)
                {
                    row.Cells[0].Value = true;
                    isSelectAllButtonAsperAdviceClicked = true;
                }
                btnasper_selall.Text = "UNSELECT ALL";
            }
            else
            {
                //  asperfilterdatagrid
                foreach (DataGridViewRow row in DataGridView3.Rows)
                {
                    row.Cells[0].Value = false;
                    isSelectAllButtonAsperAdviceClicked = false;
                }
                btnasper_selall.Text = "SELECT ALL";
                // btnasper_selall.Width = 92;
                // btnasper_mark.Location = new Point(490, 28);
            }
        }
        private void btnasper_done_Click(object sender, EventArgs e)
        {
            //exits from the "as per advise" view of the application and resets some controls
            AsPerPanel.Hide();
            pnl_deladvise.Visible = false;
            DataGridView3.DataSource = null;
            DataGridView3.Columns.Remove("cancelled");
            asperfilterTxt.Text = null;
            asperfilterBox.SelectedIndex = 0;
            btnasper_selall.Text = "SELECT ALL";
        }
        //This will hold the value of the row that current selected
        private bool isRowChecked = false;
        private void clicker()
        {
            // for checking datagridview /////////////////////////////////
            if (genClick)
            {
                if (!(bool)DataGridView3.Rows[rowIndex].Cells[0].Value)
                {
                    for (int a = 0; a < DataGridView3.Rows.Count; a++)
                    {
                        if (DataGridView3.Rows[rowIndex].Cells["PULLNO"].Value.Equals(DataGridView3.Rows[a].Cells["PULLNO"].Value))
                        {
                            DataGridView3.Rows[a].Cells[0].Value = true;
                            DataGridView3.Rows[a].Cells[0].ReadOnly = true;
                        }
                    }
                }
                else if ((bool)DataGridView3.Rows[rowIndex].Cells[0].Value)
                {
                    for (int a = 0; a < DataGridView3.Rows.Count; a++)
                    {
                        if (DataGridView3.Rows[rowIndex].Cells["PULLNO"].Value.Equals(DataGridView3.Rows[a].Cells["PULLNO"].Value))
                        {
                            DataGridView3.Rows[a].Cells[0].Value = false;
                            DataGridView3.Rows[a].Cells[0].ReadOnly = true;
                        }
                    }
                }
                return;
            }
            if ((bool)DataGridView3.Rows[rowIndex].Cells[0].Value == false)
            {
                DataGridView3.Rows[rowIndex].Cells[0].Value = true;
                DataGridView3.Rows[rowIndex].Cells[0].ReadOnly = true;
            }
            else if ((bool)DataGridView3.Rows[rowIndex].Cells[0].Value == false)
            {
                DataGridView3.Rows[rowIndex].Cells[0].Value = false;
                DataGridView3.Rows[rowIndex].Cells[0].ReadOnly = true;
            }
        }
        private void DataGridView3_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (DataGridView3.SelectedCells.Count > 0)
            {
                int selectedrowindex = DataGridView3.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = DataGridView3.Rows[selectedrowindex];
                string cellValueSelected = Convert.ToString(selectedRow.Cells["cancelled"].Value);
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
                        selectedRow.Cells["cancelled"].Value = true;
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

                        selectedRow.Cells["cancelled"].Value = false;
                    }
                }
            }
        }
        private void deliverydatedstart_KeyDown(object sender, KeyEventArgs e)
        {
            e.SuppressKeyPress = true;
        }


        private void deliverydatedocker1_KeyDown(object sender, KeyEventArgs e)
        {
            e.SuppressKeyPress = true;
        }
        private void chkALL_CheckedChanged(object sender, EventArgs e)
        {
            chkCAV1.Checked = true;
            chkCAV2.Checked = true;
            chkCAV3.Checked = true;
            chkCAV5.Checked = true;
            chkIPAI1.Checked = true;
            chkIPAI2.Checked = true;
            chkIPAI3.Checked = true;
            chkDANAM.Checked = true;
            chkMACRO.Checked = true;

        }
        private void dropDownFacilityFilter_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (dropDownFacilityFilter.SelectedIndex == -1)
            {
                return;
            }

            facFil = true;
        }
        private void btnasper_mark_Click(object sender, EventArgs e)
        {
            if (DataGridView3.Rows.Count == 0) // checks if there is data on the DataGridView
            {
                MessageBox.Show("There is nothing to mark! Please search and select the item to mark it!");
                return;
            }
            bool itemChecker = false;
            for (int a = 0; a < DataGridView3.Rows.Count; a++) // checks if the user selected any row in the DataGridView
            {
                if (DataGridView3.Columns.Contains("cancelled"))
                {
                    // Now it's safe to access the "cancelled" cell
                    if (DataGridView3.Rows[a].Cells["cancelled"].Value != null && DataGridView3.Rows[a].Cells["cancelled"].Value.Equals(true))
                    {
                        break;
                    }
                }
                if (a == DataGridView3.Rows.Count - 1)
                {
                    MessageBox.Show("There is no item selected!");
                    return;
                }
            }
            // creates a datatable for as per data if the user has selected for the first time
            if (!asperMarked)
            {
                asperTable = new DataTable();
                asperQuanTable = new DataTable();
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
            for (int cr = 0; cr < DataGridView3.Rows.Count; cr++)
            {
                if (DataGridView3.Rows[cr].Cells["cancelled"].Value != null && DataGridView3.Rows[cr].Cells["cancelled"].Value.Equals(true))
                {
                    // stores the selected row in a datatable
                    DataRow pq = asperTable.NewRow();
                    pq["TRANS_DATE"] = DateTime.Now.ToString("yyyy/MM/dd");
                    pq["PROD_DATE"] = DataGridView3.Rows[cr].Cells["PROD_DATE"].Value;
                    pq["PROD_TIME"] = DataGridView3.Rows[cr].Cells["PROD_TIME"].Value;
                    pq["DEL_DATE"] = DataGridView3.Rows[cr].Cells["DEL_DATE"].Value;
                    pq["DEL_TIME"] = DataGridView3.Rows[cr].Cells["DEL_TIME"].Value;
                    pq["JOB_NO"] = DataGridView3.Rows[cr].Cells["JOB_NO"].Value;
                    pq["FACILITY"] = DataGridView3.Rows[cr].Cells["FACILITY"].Value;
                    pq["PARTNUMBER"] = DataGridView3.Rows[cr].Cells["PARTNUMBER"].Value;
                    pq["PULL_QTY"] = DataGridView3.Rows[cr].Cells["PULL_QTY"].Value;
                    pq["OPEN_QTY"] = "";
                    pq["STOCK_QTY"] = "";
                    pq["QUANTITY_DELIVERED"] = "";
                    pq["END_BALANCE"] = "";
                    pq["GO_NUMBER"] = "";
                    pq["SKU_ASSEMBLY"] = DataGridView3.Rows[cr].Cells["SKU_ASSEMBLY"].Value;
                    pq["GO_LINE_NUMBER"] = "";
                    pq["CELL_NUM"] = DataGridView3.Rows[cr].Cells["CELLNUMBER"].Value;
                    pq["REMARKS"] = DataGridView3.Rows[cr].Cells["REMARKS"].Value;
                    pq["PULL_TICKET_NUMBER"] = DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value;
                    pq["LINE"] = DataGridView3.Rows[cr].Cells["LINE"].Value;
                    pq["VENDOR_REMARKS"] = DataGridView3.Rows[cr].Cells["VENDOR_REMARKS"].Value;
                    pq["QTY_DEL"] = DataGridView3.Rows[cr].Cells["QTY_DEL"].Value;
                    pq["ORIGINAL_PULL"] = DataGridView3.Rows[cr].Cells["ORIGINAL_PULL"].Value;
                    pq["BACKLOGTYPE"] = "AS PER ADVISE";
                    asperTable.Rows.Add(pq);
                    asperTable.AcceptChanges();
                    // stores the selected row in a datatable
                    DataRow pq2 = asperQuanTable.NewRow();
                    pq2["PROD_DATE"] = DataGridView3.Rows[cr].Cells["PROD_DATE"].Value;
                    pq2["PROD_TIME"] = DataGridView3.Rows[cr].Cells["PROD_TIME"].Value;
                    pq2["DEL_DATE"] = DataGridView3.Rows[cr].Cells["DEL_DATE"].Value;
                    pq2["DEL_TIME"] = DataGridView3.Rows[cr].Cells["DEL_TIME"].Value;
                    pq2["JOB_NO"] = DataGridView3.Rows[cr].Cells["JOB_NO"].Value;
                    pq2["FACILITY"] = DataGridView3.Rows[cr].Cells["FACILITY"].Value;
                    pq2["PARTNUMBER"] = DataGridView3.Rows[cr].Cells["PARTNUMBER"].Value;
                    pq2["PULL_QTY"] = DataGridView3.Rows[cr].Cells["PULL_QTY"].Value;
                    pq2["VENDOR_NAME"] = DataGridView3.Rows[cr].Cells["VENDOR_NAME"].Value;
                    pq2["SKU_ASSEMBLY"] = DataGridView3.Rows[cr].Cells["SKU_ASSEMBLY"].Value;
                    pq2["CELLNUMBER"] = DataGridView3.Rows[cr].Cells["CELLNUMBER"].Value;
                    pq2["REMARKS"] = DataGridView3.Rows[cr].Cells["REMARKS"].Value;
                    pq2["PULL_TICKET_NUMBER"] = DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value;
                    pq2["LINE"] = DataGridView3.Rows[cr].Cells["LINE"].Value;
                    pq2["FILEUPLOADDATE"] = DataGridView3.Rows[cr].Cells["FILEUPLOADDATE"].Value;
                    pq2["VENDOR_REMARKS"] = DataGridView3.Rows[cr].Cells["VENDOR_REMARKS"].Value;
                    pq2["ACKNOWLEDGMENT_DATE"] = DataGridView3.Rows[cr].Cells["ACKNOWLEDGMENT_DATE"].Value;
                    pq2["ACKNOWLEDGMENT_REMARKS"] = DataGridView3.Rows[cr].Cells["ACKNOWLEDGMENT_REMARKS"].Value;
                    //pq2["COMMIT_QTY"] = DataGridView3.Rows[cr].Cells["COMMIT_QTY"].Value;
                    //pq2["COMMIT_DATE"] = DataGridView3.Rows[cr].Cells["COMMIT_DATE"].Value;
                    pq2["BUYER_REMARKS_FOR_VENDOR"] = DataGridView3.Rows[cr].Cells["BUYER_REMARKS_FOR_VENDOR"].Value;
                    pq2["QTY_DELIVERED"] = DataGridView3.Rows[cr].Cells["QTY_DELIVERED"].Value;
                    pq2["DL_VARIENCE"] = DataGridView3.Rows[cr].Cells["DL_VARIENCE"].Value;
                    pq2["HITMISS"] = DataGridView3.Rows[cr].Cells["HITMISS"].Value;
                    pq2["STATUS"] = DataGridView3.Rows[cr].Cells["STATUS"].Value;
                    pq2["PULLTYPE"] = "NEW";
                    pq2["QTY_DEL"] = DataGridView3.Rows[cr].Cells["QTY_DEL"].Value;
                    pq2["ORIGINAL_PULL"] = DataGridView3.Rows[cr].Cells["ORIGINAL_PULL"].Value;
                    asperQuanTable.Rows.Add(pq2);
                    asperQuanTable.AcceptChanges();
                    // stores the selected row in a datatable
                    DataRow pq3 = asperPullLine.NewRow();
                    pq3["PULL_TICKET_NUMBER"] = DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value;
                    pq3["LINE"] = DataGridView3.Rows[cr].Cells["LINE"].Value;
                    asperPullLine.Rows.Add(pq3);
                    asperPullLine.AcceptChanges();
                }
            }
            // deletes the selected rows in the pullticket datatable and in the backup pullticket datatable
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
            DataGridView3.DataSource = null; // clears the datagridview
            DataGridView3.DataSource = pullQuanTempTable; // sets the datagridview data source
            // hides unnecessary columns in the datagridview
            DataGridView3.Columns["PROD_DATE"].Visible = false;
            DataGridView3.Columns["PROD_TIME"].Visible = false;
            DataGridView3.Columns["JOB_NO"].Visible = false;
            DataGridView3.Columns["VENDOR_NAME"].Visible = false;
            DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
            DataGridView3.Columns["CELLNUMBER"].Visible = false;
            asperMarked = true;
        }
        private bool ProcessExists(int id)
        {
            return Process.GetProcesses().Any(x => x.Id == id);
        }
        //cancel button in deliverydate(filter)
        private void btncancel_Click(object sender, EventArgs e)
        {
            btndeldate.Visible = true;
            btnasper.Visible = true;
            btncancel1.Visible = true;
            orlabel.Visible = true;
            btncancel2.Visible = true;
            pnldeliverydate.Visible = false;
            chkCAV1.Checked = false;
            chkCAV2.Checked = false;
            chkCAV3.Checked = false;
            chkCAV5.Checked = false;
            chkIPAI1.Checked = false;
            chkIPAI2.Checked = false;
            chkIPAI3.Checked = false;
            chkDANAM.Checked = false;
            chkMACRO.Checked = false;
            chkALL.Checked = false;
        }
        //minimized button
        private void buttonmin_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }
        //maximized/normal windowstate button
        private void buttonmax_Click(object sender, EventArgs e)
        {
            if (WindowState.Equals(FormWindowState.Maximized))
            {
                WindowState = FormWindowState.Normal;
            }
            else
            {
                WindowState = FormWindowState.Maximized;
            }
        }
        //exit button
        private void buttonx_Click(object sender, EventArgs e)
        {
            Timer4.Start();
        }
        //realtime datetime
        private void timer2_Tick(object sender, EventArgs e)
        {
            labelDate.Text = DateTime.Now.ToString("dddd, MMMM d, yyy ");
        }
        private void btnfiltercancel_Click(object sender, EventArgs e)
        {
            pnl_deladvise.Visible = false;
        }
        private void lblRestore_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            string resFile;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = @"\\192.168.50.40\jmctest\";
                openFileDialog.Title = "Select the Backup File";
                // Add more settings or event handlers as needed
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    resFile = openFileDialog.FileName;
                    try
                    {
                        // Drops the following tables before importing the data
                        //drop backlog table in the database
                        string backDel = "drop table BACKLOG";
                        OracleCommand backDelCom = new OracleCommand(backDel, con);
                        backDelCom.ExecuteNonQuery();
                    }
                    catch (Exception) { }
                    //Drop manualDr table in the database.
                    try
                    {
                        string backdel1 = "drop table MANUAL_DR";
                        OracleCommand backDelCom1 = new OracleCommand(backdel1, con);
                        backDelCom1.ExecuteNonQuery();
                    }
                    catch (Exception) { }
                    //Drop pullquantityrev table in the database
                    try
                    {
                        string backdel1 = "drop table PULL_QUANTITY_REV";
                        OracleCommand backDelCom1 = new OracleCommand(backdel1, con);
                        backDelCom1.ExecuteNonQuery();
                    }
                    catch (Exception) { }
                    //Drop pullticketrecord table in the database
                    try
                    {
                        string backdel1 = "drop table PULL_TICKET_RECORD";
                        OracleCommand backDelCom1 = new OracleCommand(backdel1, con);
                        backDelCom1.ExecuteNonQuery();
                    }
                    catch (Exception) { }
                    try
                    {
                        // Opens the command prompt and executes the following command to import the file
                        Process myprocess = new Process();
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.FileName = "cmd";
                        startInfo.RedirectStandardInput = true;
                        startInfo.RedirectStandardOutput = true;
                        startInfo.UseShellExecute = false;
                        startInfo.CreateNoWindow = true;
                        myprocess.StartInfo = startInfo;
                        myprocess.Start();
                        StreamReader SR = myprocess.StandardOutput;
                        StreamWriter SW = myprocess.StandardInput;
                        SW.WriteLine(@"cd C:\oraclexe\app\oracle\product\10.2.0\server\BIN");
                        SW.WriteLine($"imp mec/mec2024@192.168.50.40 buffer=4096 grants=Y file={resFile} tables=(BACKLOG, MANUAL_DR, PULL_QUANTITY_REV, PULL_TICKET_RECORD)");
                        SW.WriteLine("exit");
                        //exits command prompt window
                        //Checks if cmd is still running and shows a form with progress bar to notify the user that the process of restoring data is not yet finished
                        int procId = myprocess.Id;
                        f2.progressBar1.Style = ProgressBarStyle.Marquee;
                        f2.Show();
                        while (ProcessExists(procId))
                        {
                            f2.label1.Text = "RESTORING DATA...";
                            f2.Refresh();
                        }
                        f2.Close();
                        SW.Close();
                        SR.Close();
                        // Displays a form telling the user that restoration was successful
                        MessageBox.Show("Restored Succesfully");
                        f2.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    /* Helper method to check if a process is still running
                    bool ProcessExists(int processId)
                    {
                        return Process.GetProcessById(processId) != null;
                    }*/
                }
            }
        }
        private void deliverydatedocker1_ValueChanged(object sender, EventArgs e)
        {
            selAdvBackTbl = new DataTable(); // creates a new instance of the datatable
            selAdvBackTbl = backNewDelDate.Clone(); // clones the columns of the datatable
            if (deliverydatedocker1.Value > DateTime.Now) // if the end date is greater than the current date
            {
                // enables all the checkboxes
                chkCAV1.Enabled = true;
                chkCAV2.Enabled = true;
                chkCAV3.Enabled = true;
                chkCAV5.Enabled = true;
                chkDANAM.Enabled = true;
                chkMACRO.Enabled = true;
                chkIPAI1.Enabled = true;
                chkIPAI2.Enabled = true;
                chkIPAI3.Enabled = true;
                chkALL.Enabled = true;
                if (backNewDelDate.Rows.Count > 0)
                {
                    DateTime datePick3 = DateTime.Now.AddDays(1); // sets the start date for the loop
                    DateTime datePick4 = deliverydatedocker1.Value; // gets and sets the end date for the loop
                    while (datePick3 <= datePick4)
                    {
                        // selects the BACKLOG data based on the date in the loop and BACKLOG is not "as per advise"
                        DataView negaView = new DataView(backNewDelDate);
                        negaView.RowFilter = "DEL_DATE = '" + datePick3 + "' AND BACKLOGTYPE IS NULL";
                        forFilter = new DataTable();
                        forFilter = backNewDelDate.Clone();
                        forFilter = negaView.ToTable();
                        foreach (DataRow row in forFilter.Rows)
                        {
                            selAdvBackTbl.ImportRow(row);
                            selAdvBackTbl.AcceptChanges();
                        }
                        datePick3 = datePick3.AddDays(1); // adds one day to the current date in the loop
                    }
                    if (selAdvBackTbl.Rows.Count > 0)
                    {
                        DataTable samepartTbl = new DataTable();
                        samepartTbl = selAdvBackTbl.Clone();
                        samepartTbl = selAdvBackTbl.DefaultView.ToTable(true, "FACILITY");
                        //selects the checkboxes that has the same label as the facility in the loop
                        for (int a = 0; a < samepartTbl.Rows.Count; a++)
                        {
                            string facility = samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper();
                            if(facility == "CAV1")
                            {
                                chkCAV1.Checked = true;
                                chkCAV1.Enabled = false;
                            }
                            else if(facility == "CAV2")
                            {
                                chkCAV2.Checked = true;
                                chkCAV2.Enabled = false;
                            }
                            else if(facility == "CAV3" || facility == "CAV3/JAC")
                            {
                                chkCAV3.Checked = true;
                                chkCAV3.Enabled = false;
                            }
                            else if (facility == "CAV5")
                            {
                                chkCAV5.Checked = true;
                                chkCAV5.Enabled = false;
                            }
                            else if (facility == "DANAM T" || facility == "DANAMT")
                            {
                                chkIPAI1.Checked = true;
                                chkIPAI1.Enabled = false;
                            }
                            else if (facility == "DKP")
                            {
                                chkIPAI2.Checked = true;
                                chkIPAI2.Enabled = false;
                            }
                            else if (facility == "CLP" || facility == "CLP/CAV5")
                            {
                                chkIPAI3.Checked = true;
                                chkIPAI3.Enabled = false;
                            }
                            else if (facility == "MACRO")
                            {
                                chkMACRO.Checked = true;
                                chkMACRO.Enabled = false;
                            }
                            else if (facility == "DANAM")
                            {
                                chkDANAM.Checked = true;
                                chkDANAM.Enabled = false;
                            }
                        }
                        facFil = true;
                    }
                    // selects the "ALL" checkbox if all the facility checkboxes were selected
                    if (chkCAV1.Checked == true && chkCAV2.Checked == true && chkCAV3.Checked == true && chkCAV5.Checked == true && chkIPAI1.Checked == true && chkIPAI2.Checked == true && chkIPAI3.Checked == true && chkDANAM.Checked == true && chkMACRO.Checked == true)
                    {
                        chkALL.Checked = true;
                        chkALL.Enabled = false;
                    }
                    date2sel = true; //end date or "TO" date was selected
                }
            }
            else
            {
                //BunifuCustomLabel22.Visible = true;
                dropDownFacilityFilter.Visible = false;
                chkCAV1.Enabled = false;
                chkCAV2.Enabled = false;
                chkCAV3.Enabled = false;
                chkCAV5.Enabled = false;
                chkDANAM.Enabled = false;
                chkMACRO.Enabled = false;
                chkIPAI1.Enabled = false;
                chkIPAI2.Enabled = false;
                chkIPAI3.Enabled = false;
                chkALL.Enabled = false;
                chkCAV1.Checked = false;
                chkCAV2.Checked = false;
                chkCAV3.Checked = false;
                chkCAV5.Checked = false;
                chkDANAM.Checked = false;
                chkMACRO.Checked = false;
                chkIPAI1.Checked = false;
                chkIPAI2.Checked = false;
                chkIPAI3.Checked = false;
                chkALL.Checked = false;
                facFil = false;
                date2sel = true; // end date or "TO" date was selected
                selAdvBackTbl = new DataTable();
                backNewDelDate = advBacklogTbl.Copy(); // copy the data of the datatable holding the BACKLOG backup
            }
        }
        private void deliverydatestart_ValueChanged(object sender, EventArgs e)
        {
            date1sel = true; // start date or "FROM" date was selected
        }
        private void btncancel2_Click(object sender, EventArgs e)
        {
            pnl_deladvise.Visible = false;
            btnasper.Visible = false;
            btndeldate.Visible = false;
            orlabel.Visible = false;
        }
        //button to show data
        private void btndata_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            if (this.WindowState == FormWindowState.Maximized)
            {
                form3.WindowState = FormWindowState.Maximized;
                form3.ShowDialog();
            }
            else
            {
                form3.WindowState = FormWindowState.Normal;
                form3.ShowDialog();
            }

            genClick = false;
            form3.recorddatagrid.ClearSelection();
        }
        //-->start button custom style onhover
        private void btndata_MouseEnter(object sender, EventArgs e)
        {
            btndata.ForeColor = System.Drawing.Color.White;
        }
        private void btndata_MouseLeave(object sender, EventArgs e)
        {
            btndata.ForeColor = System.Drawing.Color.White;
        }
        private void buttonx_MouseEnter(object sender, EventArgs e)
        {
            buttonx.BackColor = System.Drawing.Color.Red;
        }
        private void buttonx_MouseLeave(object sender, EventArgs e)
        {
            buttonx.BackColor = System.Drawing.Color.Transparent;
        }
        private void buttonmax_MouseEnter(object sender, EventArgs e)
        {
            buttonmax.BackColor = System.Drawing.Color.Orange;
        }
        private void buttonmax_MouseLeave(object sender, EventArgs e)
        {
            buttonmax.BackColor = System.Drawing.Color.Transparent;
        }
        private void buttonmin_MouseEnter(object sender, EventArgs e)
        {
            buttonmin.BackColor = System.Drawing.Color.Green;
        }
        private void buttonmin_MouseLeave(object sender, EventArgs e)
        {
            buttonmin.BackColor = System.Drawing.Color.Transparent;
        }
        private void btnrestore_MouseEnter(object sender, EventArgs e)
        {
            btnrestore.BackColor = System.Drawing.Color.LightSteelBlue;
        }
        private void btnrestore_MouseLeave(object sender, EventArgs e)
        {
            btnrestore.BackColor = System.Drawing.Color.Transparent;
        }
        //button custom style onhover -->end<--//
        //timer for transition onOpen
        private void Timer3_Tick(object sender, EventArgs e)
        {
            if (Opacity == 1)
            {
                Timer3.Stop();
            }
            Opacity += .2;
        }
        //timer for transition onclose
        private void Timer4_Tick(object sender, EventArgs e)
        {
            if (Opacity <= 0)
            {
                this.Close();
            }
            Opacity -= 0.2;
        }
        //import pull ticket button
        private void btnpull_tkt_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();

            //Checks if the file has already been imported
            if (importPull == true)
            {
                return;
            }
            //FileDialog for selecting the file to be imported
            string strFileName;
            openFD.InitialDirectory = "'C:";
            openFD.Filter = "All Files (*.*)|*.*";
            openFD.FilterIndex = 2;
            openFD.Title = "Choose a File";
            openFD.RestoreDirectory = true;
            if (openFD.ShowDialog() == DialogResult.OK)
            {
                strFileName = openFD.FileName;
                if (!string.IsNullOrEmpty(FileSystem.Dir(openFD.FileName)))
                {
                    if (!openFD.SafeFileName.ToString().ToUpper().Contains("PULL"))
                    {
                        MessageBox.Show("The file you are trying to import is named " + openFD.SafeFileName + "\r\n\r\nMake sure you are importing the correct file!");
                        return;
                    }
                    backlogDel = false; //Tells the program that the data is from a file and not from a BACKLOG data
                    f2.label2.Text = "Reading Data from " + openFD.SafeFileName;
                    f2.Refresh();
                    pullFileName = openFD.FileName;
                    try
                    {
                        excelCleaner(strFileName); //Excel File checker 
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Please close the excel file before importing then try again.");
                        f2.Close();
                        return;
                    }
                    //Create new instances of datatables
                    forFilter = new DataTable();
                    pullrecordTable = new DataTable();
                    excelData2 = new DataTable();
                    pullQuanTempTable = new DataTable();
                    pullQuantityTable = new DataTable();
                    pullbackTable = new DataTable();
                    dataGridView4.DataSource = null;
                    backLogChecking();
                    refreshData(strFileName);
                    if (errorDelDate == true)
                    {
                        errorDelDate = false;
                        importPull = false;
                        return;
                    }
                    f2.label2.Text = "Importing Complete";
                    f2.Refresh();
                    f2.Close();
                    dataGridView4.Visible = true;
                    pulltktgrid.Visible = false;
                    btnpull_tkt.BackColor = Color.Navy;
                    btnpull_tkt.Enabled = false;
                    btnfilter.Visible = true;
                    pnl_import2.Location = new System.Drawing.Point(35, 105);
                    slider.Visible = false;
                    refesherOrb.Visible = true;
                    MessageBox.Show(openFD.SafeFileName + " imported successfully!", "SUCCESS!!!");
                    importPull = true;
                    try
                    {
                        xlApp.Quit();
                        oExcel2.Quit();
                    }
                    catch (Exception ex) { }
                }
                else
                {
                    MessageBox.Show("File Not Found!");
                }
            }
        }
        private void excelCleaner(string fileName)
        {
            //Create an instance of Excel.Application
            Excel.Application oExcel2 = new Excel.Application();
            //Open the specified file
            xlWorkbook = oExcel2.Workbooks.Open(pullFileName);
            xlWorksheet = xlWorkbook.Worksheets[1];
            ReplaceInValuesAndFormulas(); //Replaces some characters in the excel file
            oExcel2.DisplayAlerts = false;
            xlWorksheet.Name = "PullTicket";
            xlWorkbook.SaveAs(pullFileName, XlFileFormat.xlXMLSpreadsheet);
            xlWorkbook.Close();
            try
            {
                xlApp.Quit();
                oExcel2.Quit();
            }
            catch (Exception ex) { }
            //Releases or disposes the object used
            ReleaseObject(xlWorksheet);
            ReleaseObject(xlWorkbook);
            ReleaseObject(oExcel2);
        }
        public void ReplaceInValuesAndFormulas()
        {
            //Replaces some characters in the excel file
            Excel.Application oExcel2 = new Excel.Application();
            oExcel2.DisplayAlerts = false;
            //xlWorksheet.Cells.Replace("'", "", XlLookAt.xlPart, XlSearchOrder.xlByRows, false, false, false);
            xlWorksheet.Cells.Replace("'", "", XlLookAt.xlPart, XlSearchOrder.xlByRows, false, false, false);
        }
        public void ReleaseObject(object obj)
        {
            try
            {
                int intRel = 0;
                do
                {
                    intRel = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                } while (intRel > 0);
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
        private void backLogChecking()
        {
            //Check the database for backlog records
            DataTable backlogChecktbl = new DataTable();
            string backCheck = $"select * from BACKLOG";
            OracleDataAdapter adapter2 = new OracleDataAdapter(backCheck, con);
            adapter2.Fill(backlogChecktbl);
            int a = backlogChecktbl.Rows.Count;
            if (a == 0)
            {
                backlogCheck = false;
            }
            else
            {
                backlogCheck = true;
            }
        }
        private void refreshData(string myfiledirect)
        {
            pullQuanTempTable = new DataTable(); // This is the table that will store the final pullticket data that the application will use
            pullrecordTable = new DataTable();
            pullQuantityTable = new DataTable(); // Table for saving the imported pullticket records from the excel file
            cloneQuanTempTable = new DataTable(); // Copy of pullticket records from excel file and BACKLOG for editing
            pullDeleteTable = new DataTable(); // Table for deleted records in the excel file that was saved in the database
            pullRevTable = new DataTable(); // Table for pullticket revisions (cancelled, closed, changed qty)
            manualDrTable = new DataTable(); // Table for manual dr
            advBacklogTbl = new DataTable(); // Table for advance BACKLOG (adjusted del date, as per advise)
            amarokTbl = new DataTable(); // Table for amarok partnumber
            kanbanTbl = new DataTable(); // Table for kanban partnumber
                                         // ////////////////////////////////// nov 9 for adjusted BACKLOG del date ////////////////////////
            backNewDelDate = new DataTable();
            historyBack = new DataTable();
            Form2 f2 = new Form2();
            // Dim backNewDelDateCol As DataColumn = backNewDelDate.Columns.Add("DEL_DATE", Type.GetType("System.DateTime"));
            // backNewDelDate.Columns.Add("PULL_TICKET_NUMBER", Type.GetType("System.String"));
            // backNewDelDate.Columns.Add("LINE", Type.GetType("System.String"));

            // creates a table for adjusted delivery dates of BACKLOG or as per advise pullticket
            // 為調整後的 BACKLOG 交貨日期或依照拉單中的建議建立一個表格。
            DataColumn backNewDelDateCol = backNewDelDate.Columns.Add("TRANS_DATE", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("PROD_DATE", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("PROD_TIME", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("DEL_DATE", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("DEL_TIME", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("JOB_NO", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("FACILITY", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("PARTNUMBER", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("PULL_QTY", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("OPEN_QTY", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("STOCK_QTY", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("QUANTITY_DELIVERED", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("END_BALANCE", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("GO_NUMBER", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("SKU_ASSEMBLY", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("GO_LINE_NUMBER", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("CELL_NUM", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("REMARKS", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("PULL_TICKET_NUMBER", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("LINE", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("VENDOR_REMARKS", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("QTY_DEL", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("ORIGINAL_PULL", Type.GetType("System.String"));
            backNewDelDate.Columns.Add("BACKLOGTYPE", Type.GetType("System.String"));

            // creates a table for storing the final pullticket data
            DataColumn pullQTcol = pullQuanTempTable.Columns.Add("PROD_DATE", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("PROD_TIME", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("DEL_DATE", Type.GetType("System.DateTime"));
            pullQuanTempTable.Columns.Add("DEL_TIME", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("JOB_NO", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("FACILITY", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("PARTNUMBER", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("PULL_QTY", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("VENDOR_NAME", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("SKU_ASSEMBLY", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("CELLNUMBER", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("REMARKS", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("PULL_TICKET_NUMBER", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("LINE", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("FILEUPLOADDATE", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("VENDOR_REMARKS", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("ACKNOWLEDGMENT_DATE", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("ACKNOWLEDGMENT_REMARKS", Type.GetType("System.String"));
            // //// Added Commit QTY and Commit_Date in the temporary table, so the log app can accept excel file /////
            // pullQuanTempTable.Columns.Add("COMMIT_QTY", Type.GetType("System.String"));
            // pullQuanTempTable.Columns.Add("COMMIT_DATE", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("BUYER_REMARKS_FOR_VENDOR", Type.GetType("System.String"));
            // pullQuanTempTable.Columns.Add("REASON_CODE", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("QTY_DELIVERED", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("DL_VARIENCE", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("HITMISS", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("STATUS", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("PULLTYPE", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("QTY_DEL", Type.GetType("System.String"));
            pullQuanTempTable.Columns.Add("ORIGINAL_PULL", Type.GetType("System.String"));

            // creates a table for pullticket revisions
            DataColumn pullRevcol = pullRevTable.Columns.Add("DEL_DATE", Type.GetType("System.DateTime"));
            pullRevTable.Columns.Add("FACILITY", Type.GetType("System.String"));
            pullRevTable.Columns.Add("PARTNUMBER", Type.GetType("System.String"));
            pullRevTable.Columns.Add("PREVIOUS_PULL", Type.GetType("System.String"));
            pullRevTable.Columns.Add("NEW_PULL", Type.GetType("System.String"));
            pullRevTable.Columns.Add("REMARKS", Type.GetType("System.String"));
            pullRevTable.Columns.Add("PULL_TICKET_NUMBER", Type.GetType("System.String"));
            pullRevTable.Columns.Add("LINE", Type.GetType("System.String"));
            pullRevTable.Columns.Add("DATE_REVISED", Type.GetType("System.Int64")); // //// added dec 12 /////////

            // creates a table for deleted records
            DataColumn pullDelCol = pullDeleteTable.Columns.Add("PULL_TICKET_NUMBER", Type.GetType("System.String"));
            pullDeleteTable.Columns.Add("LINE", Type.GetType("System.String"));

            // creates a table for manual dr
            DataColumn manualDrCol = manualDrTable.Columns.Add("PROD_DATE", Type.GetType("System.String"));
            manualDrTable.Columns.Add("PROD_TIME", Type.GetType("System.String"));
            manualDrTable.Columns.Add("DEL_DATE", Type.GetType("System.String"));
            manualDrTable.Columns.Add("DEL_TIME", Type.GetType("System.String"));
            manualDrTable.Columns.Add("JOB_NO", Type.GetType("System.String"));
            manualDrTable.Columns.Add("FACILITY", Type.GetType("System.String"));
            manualDrTable.Columns.Add("PARTNUMBER", Type.GetType("System.String"));
            manualDrTable.Columns.Add("PULL_QTY", Type.GetType("System.Int32"));
            manualDrTable.Columns.Add("VENDOR_NAME", Type.GetType("System.String"));
            manualDrTable.Columns.Add("SKU_ASSEMBLY", Type.GetType("System.String"));
            manualDrTable.Columns.Add("CELLNUMBER", Type.GetType("System.String"));
            manualDrTable.Columns.Add("REMARKS", Type.GetType("System.String"));
            manualDrTable.Columns.Add("PULL_TICKET_NUMBER", Type.GetType("System.String"));
            manualDrTable.Columns.Add("LINE", Type.GetType("System.String"));
            manualDrTable.Columns.Add("FILEUPLOADDATE", Type.GetType("System.String"));
            manualDrTable.Columns.Add("VENDOR_REMARKS", Type.GetType("System.String"));
            manualDrTable.Columns.Add("ACKNOWLEDGMENT_REMARKS", Type.GetType("System.String"));
            manualDrTable.Columns.Add("BUYER_REMARKS_FOR_VENDOR", Type.GetType("System.String"));
            manualDrTable.Columns.Add("QTY_DELIVERED", Type.GetType("System.String"));
            manualDrTable.Columns.Add("DL_VARIENCE", Type.GetType("System.String"));
            manualDrTable.Columns.Add("HITMISS", Type.GetType("System.String"));
            manualDrTable.Columns.Add("STATUS", Type.GetType("System.String"));

            // creates a table for BACKLOG
            DataTable backTempotbl = new DataTable();
            DataColumn backTempotblcol = backTempotbl.Columns.Add("Prod_date", Type.GetType("System.String"));
            backTempotbl.Columns.Add("Prod_time", Type.GetType("System.String"));
            backTempotbl.Columns.Add("Del_date", Type.GetType("System.String"));
            backTempotbl.Columns.Add("Del_time", Type.GetType("System.String"));
            backTempotbl.Columns.Add("Job_no", Type.GetType("System.String"));
            backTempotbl.Columns.Add("Facility", Type.GetType("System.String"));
            backTempotbl.Columns.Add("Partnumber", Type.GetType("System.String"));
            backTempotbl.Columns.Add("Pull_qty", Type.GetType("System.String"));
            backTempotbl.Columns.Add("SKU_ASSEMBLY", Type.GetType("System.String"));
            backTempotbl.Columns.Add("CELL_NUM", Type.GetType("System.String"));
            backTempotbl.Columns.Add("REMARKS", Type.GetType("System.String"));
            backTempotbl.Columns.Add("PULL_TICKET_NUMBER", Type.GetType("System.String"));
            backTempotbl.Columns.Add("LINE", Type.GetType("System.String"));
            backTempotbl.Columns.Add("VENDOR_REMARKS", Type.GetType("System.String"));
            backTempotbl.Columns.Add("Quantity_delivered", Type.GetType("System.String"));
            backTempotbl.Columns.Add("End_Balance", Type.GetType("System.String"));
            backTempotbl.Columns.Add("QTY_DEL", Type.GetType("System.Int32"));
            backTempotbl.Columns.Add("ORIGINAL_PULL", Type.GetType("System.Int32"));
            // ///////////////////////////////////////////////////////////////////////////////

            // opens and read the data in the pullticket excel file
            oExcel2 = new Excel.Application();
            oExcel2.DisplayAlerts = false;
            xlWorkbook = oExcel2.Workbooks.Open(pullFileName, Excel.XlFileAccess.xlReadOnly);
            xlWorksheet = xlWorkbook.Worksheets[1];
            int lRow = xlWorksheet.Range["A" + xlWorksheet.Rows.Count].End[Excel.XlDirection.xlUp].Row;
            Excel.Range range = xlWorksheet.Range["A1:Z" + lRow]; // gets and sets the range of data in the excel file
            object[,] data = (object[,])range.Value; // stores the data in an array object

            // Create new Columns in pullquantity table and adds the data from the excel file
            for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                f2.progressBar1.Maximum = range.Columns.Count;
                f2.progressBar1.Value = cCnt;
                f2.label2.Text = "Importing Pull Ticket . . .";
                f2.Refresh();

                DataColumn Column = new DataColumn();
                if (cCnt == 3)
                {
                    Column.DataType = System.Type.GetType("System.DateTime");
                }
                else
                {
                    Column.DataType = System.Type.GetType("System.String");
                }
                // Column.ColumnName = cCnt.ToString();
                Column.ColumnName = data[1, cCnt].ToString();
                pullQuantityTable.Columns.Add(Column);

                // Create rows for pullquantity table
                for (int rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {
                    string CellVal = string.Empty;
                    CellVal = data[rCnt, cCnt]?.ToString();

                    // checks if delivery date or time in the excel file is null
                    if (cCnt == 3 && string.IsNullOrEmpty(CellVal))
                    {
                        // Set delivery date to now if null
                        CellVal = DateTime.Now.ToString("MM/dd/yyyy");
                    }
                    /*if (cCnt == 4)
                    {
                        cellVal = DateTime.Now.ToString("hh:00:00 tt");
                    }*/

                    if (cCnt == 2 || cCnt == 4)
                    {
                        if (double.TryParse(CellVal, out double numericValue))
                        {
                            // If it's a number, handle it accordingly
                            if (cCnt == 4)
                            {
                                TimeSpan ts = TimeSpan.FromDays(numericValue);
                                if (string.IsNullOrEmpty(CellVal))
                                {
                                    CellVal = DateTime.Now.ToString("hh:00:00 tt");
                                }
                                else
                                {
                                    CellVal = $"{ts.Hours:00}:{ts.Minutes:00}:{ts.Seconds:00}";
                                }
                            }
                        }
                    }
                    DataRow row;
                    if (cCnt == 1)
                    {
                        row = pullQuantityTable.NewRow();
                        row[Column.ColumnName] = CellVal;
                        pullQuantityTable.Rows.Add(row);
                    }
                    else
                    {
                        row = pullQuantityTable.Rows[rCnt - 2];
                        row[Column.ColumnName] = CellVal;
                    }
                }
            }

            pullQuantityTable.AcceptChanges();
            amarokTbl = pullQuantityTable.Copy(); // creates a copy of the pullticket records

            xlWorkbook.Close();
            try
            {
                xlApp.Quit();
            }
            catch (Exception) { }

            try
            {
                ReleaseObject(lRow);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkbook);
                ReleaseObject(oExcel2);
            }
            catch (Exception) { }

            // selects pullticket data from the database
            string selectPullRec = "select * from pull_ticket_record";
            OracleDataAdapter pullRecAd = new OracleDataAdapter(selectPullRec, con);
            DataTable pullRecTbl = new DataTable();
            pullRecAd.Fill(pullRecTbl); // stores the data in a datatable
            int ticketCounter = pullRecTbl.Rows.Count;

            // selects BACKLOG data from the database
            string Backsel = "SELECT Prod_date, Prod_time, Del_date, Del_time, Job_no,Facility, Partnumber,Pull_qty, SKU_ASSEMBLY, CELL_NUM, REMARKS, PULL_TICKET_NUMBER, LINE,VENDOR_REMARKS, Quantity_delivered, End_Balance, QTY_DEL, ORIGINAL_PULL, BACKLOGTYPE, HISTORY, BALANCE from BACKLOG";
            OracleDataAdapter backSelAd = new OracleDataAdapter(Backsel, con);
            backSelAd.Fill(backTempotbl); // stores the data in a datatable
            historyBack = backTempotbl.Copy(); // creates a copy of the BACKLOG

            int r = 0;

            bool pullExisting = false; // checks if the data is still in the pullticket record
                                       // checking of pullTicket record
            for (int a = 0; a < pullRecTbl.Rows.Count; a++)
            {
                f2.progressBar1.Maximum = pullRecTbl.Rows.Count;
                r++;
                f2.progressBar1.Value = r;
                f2.label2.Text = "Checking for duplicate pull ticket . . .";
                f2.Refresh();
                // checks if the imported pullticket data contains the data of the current row in the loop
                DataRow[] z;
                string expr3 = $"Pull_Ticket_No = '{pullRecTbl.Rows[a]["PULL_TICKET_NUMBER"]}' AND Line = '{pullRecTbl.Rows[a]["LINE"]}'";
                z = pullQuantityTable.Select(expr3);

                if (z.Length > 0) // if the current row from the database has the same data with the imported excel file
                {
                    if (!pullRecTbl.Rows[a]["PULL_QTY"].Equals(z[0]["Pull_Qty"])) // if the pull quantity from the database and excel file has different pull quantity
                    {
                        // ////////////////////////// added for cancelled pull////////////////////////////
                        // added nov 13 for cancelled pull but the quantity isn't updated or changed
                        if (!z[0]["Remarks"].ToString().ToLower().Contains("cancel") && !pullRecTbl.Rows[a]["REMARKS"].ToString().ToLower().Contains("cancel")) // if the pull isn't cancelled
                        {
                            DataRow pq = pullRevTable.NewRow(); // adds new datarow in datatable
                            pq["DEL_DATE"] = z[0]["DEL_DATE"];
                            pq["FACILITY"] = z[0]["FACILITY"];
                            pq["PARTNUMBER"] = z[0]["PART_NO"];
                            pq["PREVIOUS_PULL"] = pullRecTbl.Rows[a]["PULL_QTY"];
                            pq["NEW_PULL"] = z[0]["Pull_Qty"];
                            pq["REMARKS"] = z[0]["Pull_Qty"].Equals(0) ? "CANCEL" : z[0]["Remarks"];
                            pq["PULL_TICKET_NUMBER"] = z[0]["Pull_Ticket_No"];
                            pq["LINE"] = z[0]["Line"];
                            pq["DATE_REVISED"] = DateTime.Now.ToString("yyyyMMddhhmmss");
                            pullRevTable.Rows.Add(pq);
                            pullRevTable.AcceptChanges();

                            // checks the BACKLOG datatable if it has the same data of the current row in the loop
                            DataRow[] z2;
                            string expr4 = $"PULL_TICKET_NUMBER = '{pullRecTbl.Rows[a]["PULL_TICKET_NUMBER"]}' AND LINE = '{pullRecTbl.Rows[a]["LINE"]}'";
                            z2 = backTempotbl.Select(expr4);

                            if (z2.Length > 0)
                            {
                                if (Convert.ToInt32(z2[0]["QTY_DEL"]) >= Convert.ToInt32(z[0]["Pull_Qty"])) // checks if the quantity delivered record in the BACKLOG is greater than or equal to the pull quantity of the imported file
                                {
                                    string backexp = $"PULL_TICKET_NUMBER = '{pullRecTbl.Rows[a]["PULL_TICKET_NUMBER"]}' AND LINE = '{pullRecTbl.Rows[a]["LINE"]}'";
                                    DataRow[] backrow = backTempotbl.Select(backexp);
                                    // deletes the rows in the BACKLOG datatable
                                    foreach (DataRow backrow2 in backrow)
                                    {
                                        backrow2.Delete();
                                    }
                                    backTempotbl.AcceptChanges();
                                }
                                else
                                {
                                    if (!z[0]["DEL_DATE"].ToString().ToLower().Contains(z2[0]["DEL_DATE"].ToString().ToLower().Replace("  12:00:00 AM", "")))
                                    {
                                        z2[0]["DEL_DATE"] = z[0]["DEL_DATE"]; // replaces the delivery date of the BACKLOG if it was adjusted
                                    }
                                    z2[0]["ORIGINAL_PULL"] = Convert.ToInt32(z[0]["Pull_Qty"]); // replaces the original pull with the current pull quantity in the imported excel file
                                    z2[0]["REMARKS"] = "-";
                                    backTempotbl.AcceptChanges(); // apply changes to the datatable
                                }
                            }

                            // checks if the pull in the imported excel file was cancelled
                        }
                        else if (z[0]["Remarks"].ToString().ToLower().Contains("cancel") && !pullRecTbl.Rows[a]["REMARKS"].ToString().ToLower().Contains("cancel"))
                        {
                            DataRow pq = pullRevTable.NewRow();
                            // Additional logic for handling cancelled pulls would go here
                        }
                    }
                }
            }
            pullQuanTempTable.DefaultView.Sort = "DEL_DATE ASC, DEL_TIME ASC";
            pullQuanTempTable = pullQuanTempTable.DefaultView.ToTable();
            pullQuanTempTable.AcceptChanges();
            pulltktgrid.DataSource = pullQuanTempTable;
            pulltktgrid.AllowUserToAddRows = true;
            pulltktgrid.Columns[0].DefaultCellStyle.Format = "MM/dd/yyyy";
            pulltktgrid.Columns[2].DefaultCellStyle.Format = "MM/dd/yyyy";
            pulltktgrid.Columns[14].DefaultCellStyle.Format = "MM/dd/yyyy";
        }
        //import cxmr620 button
        private void btncxmr620_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            if (btnpull_tkt.BackColor == Color.Navy)
            {
                deliverydatechange();
                if (importCxmr)
                {
                    return;
                }
                string strFileName;
                openFD.InitialDirectory = "'C:\'";
                openFD.Filter = "Excel Office | *.xlsx; *.xls";
                openFD.Title = "Choose a File";
                openFD.FilterIndex = 2;
                openFD.RestoreDirectory = true;
                if (openFD.ShowDialog().Equals(DialogResult.OK))
                {
                    strFileName = openFD.FileName;
                    ptkFileName.Text = openFD.FileName;
                    if (!string.IsNullOrEmpty(FileSystem.Dir(openFD.FileName)))
                    {
                        if (!openFD.SafeFileName.ToString().ToUpper().Contains("CXMR620"))
                        {
                            Interaction.MsgBox("The file you are trying to import is named " + openFD.SafeFileName + Microsoft.VisualBasic.Constants.vbCrLf + Microsoft.VisualBasic.Constants.vbCrLf + "Make sure you are importing the correct file!");
                            return;
                        }
                        else
                        {
                            if (strFileName != "")
                            {
                                f2.label2.Text = "Reading Data . . .";
                                f2.Refresh();
                                try
                                {
                                    unpostedsave(strFileName, openFD.SafeFileName);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Please close the excel application then try importing again!");
                                    f2.Close();
                                    btncxmr620.BackColor = Color.SteelBlue;
                                }
                                f2.label2.Text = "IMPORTING COMPLETE!";
                                f2.Close();
                                btncxmr620.BackColor = Color.Navy;
                                dataGridView4.Visible = true;
                                pulltktgrid.Visible = false;
                                slider.Visible = false;
                                refesherOrb.Visible = true;
                                importCxmr = true;
                                btncxmr620.Enabled = false;
                                btnfilter.Visible = false;
                                pnl_import2.Location = new System.Drawing.Point(35, 60);
                            }
                            else
                            {
                                MessageBox.Show("FILE NOT FOUND!");
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Import pull_ticket first!", "Warning!", MessageBoxButtons.OK);
            }
        }
        //method for importing CXMR620
        private void unpostedsave(string myfiledirect, string filename)
        {
            unpostedTbl = new DataTable();
            //pullRevTable = new DataTable();
            Form2 f2 = new Form2();
            f2.Show();
            oExcel2 = new Excel.Application();
            oExcel2.DisplayAlerts = false;
            xlWorkbook = oExcel2.Workbooks.Open(myfiledirect, XlFileAccess.xlReadOnly);
            xlWorksheet = xlWorkbook.Worksheets[1];
            int lRow = xlWorksheet.Range["A" + xlWorksheet.Rows.Count.ToString()].End[XlDirection.xlUp].Row;
            //int lRow = xlWorksheet.Cells[xlWorksheet.Rows.Count, "A"].End(XlDirection.xlUp).Row;
            Range range = xlWorksheet.Range["L7:T" + lRow];
            object[,] data = range.Value;
            for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                f2.progressBar1.Maximum = range.Columns.Count;
                f2.progressBar1.Value = cCnt;
                f2.label2.Text = "Fetching data from excel . . .";
                f2.Refresh();
                DataColumn Column = new DataColumn();
                Column.DataType = typeof(string);
                Column.ColumnName = cCnt.ToString();
                unpostedTbl.Columns.Add(Column);
                for (int rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {
                    string rCnt1 = Convert.ToString(rCnt);
                    string cCnt1 = Convert.ToString(cCnt);
                    string CellVal = string.Empty;
                    CellVal = Convert.ToString(data[rCnt, cCnt]);
                    DataRow Row;
                    // Adds the row to the DataTable           
                    if (cCnt == 1)
                    {
                        Row = unpostedTbl.NewRow();
                        Row[Column.ColumnName.ToString()] = CellVal;
                        unpostedTbl.Rows.Add(Row);
                    }
                    else
                    {
                        Row = unpostedTbl.Rows[rCnt - 2];
                        Row[Column.ColumnName.ToString()] = CellVal;
                    }
                }
            }
            unpostedTbl = unpostedTbl.DefaultView.ToTable(false, "1", "3", "4", "9", "2"); // selects only the columns needed
            DataRow[] prt = unpostedTbl.Select("[1] is null or [1] LIKE '%ORDER%'"); // selects the rows to be deleted
                                                                                     // deletes the rows
            foreach (DataRow prt2 in prt)
            {
                prt2.Delete();
            }
            unpostedTbl.AcceptChanges(); // applies the changes made to the datatable
            MessageBox.Show("Import " + filename + " successfully!!", "Success!!");
            f2.Close();
            xlWorkbook.Close();
            try
            {
                xlApp.Quit(); // quits the excel application
                oExcel2.Quit();
            }
            catch (Exception ex)
            {
                return;
            }
            // release or dispose the objects
            ReleaseObject(lRow);
            ReleaseObject(xlWorksheet);
            ReleaseObject(xlWorkbook);
            ReleaseObject(oExcel2);
        }
        // import Aimr407 button
        private void btnaimr407_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            if (btncxmr620.BackColor == Color.Navy)
            {
                if (importInv == true)
                {
                    return;
                }
                string strFileName;
                openFD.InitialDirectory = "'C:\'";
                openFD.Filter = "Excel Office | *.xlsx; *.xls";
                openFD.Title = "Choose a File";
                openFD.FilterIndex = 2;
                openFD.RestoreDirectory = true;
                if (openFD.ShowDialog().Equals(DialogResult.OK))
                {
                    strFileName = openFD.FileName;
                    ptkFileName.Text = openFD.FileName;
                    if (!string.IsNullOrEmpty(FileSystem.Dir(openFD.FileName)))
                    {
                        if (!openFD.SafeFileName.ToString().ToUpper().Contains("AIMR407"))
                        {
                            Interaction.MsgBox("The file you are trying to import is named " + openFD.SafeFileName + Microsoft.VisualBasic.Constants.vbCrLf + Microsoft.VisualBasic.Constants.vbCrLf + "Make sure you are importing the correct file!");
                            return;
                        }
                        else
                        {
                            if (strFileName != "")
                            {
                                f2.label2.Text = "Reading Data. . .";
                                f2.Refresh();
                                try
                                {
                                    refeshData(strFileName, openFD.SafeFileName);
                                    Inventory_saveRev();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                    f2.Close();
                                }
                                btnaimr407.BackColor = Color.Navy;
                                dataGridView4.Visible = true;
                                pulltktgrid.Visible = false;
                                slider.Visible = false;
                                refesherOrb.Visible = true;
                                f2.label2.Text = "Importing Complete";
                                f2.Refresh();
                                f2.Close();
                                importInv = true;
                                btnaimr407.Enabled = false;
                            }
                        }
                    }
                }
                try
                {
                    xlWorkbook.Close();
                    xlApp.Quit();
                    oExcel2.Quit();
                    ReleaseObject(xlWorksheet);
                    ReleaseObject(xlWorkbook);
                    ReleaseObject(oExcel2);
                }catch(Exception ex) { }
            }
            else
            {
                MessageBox.Show("Import CXMR620 first!", "Warning!", MessageBoxButtons.OK);
            }
        }
        //Method for importing AIMR407
        private void refeshData(string myfiledirect, string filename)
        {
            excelData = new DataTable();
            inventoryTable = new DataTable();
            invTblCopy = new DataTable();
            kanbanTbl = pullQuanTempTable.Copy();
            Form2 f2 = new Form2();
            f2.Show();
            oExcel2 = new Excel.Application();
            oExcel2.DisplayAlerts = false;
            xlWorkbook = oExcel2.Workbooks.Open(myfiledirect, XlFileAccess.xlReadOnly);
            xlWorksheet = xlWorkbook.Worksheets[1];
            //int lRow = xlWorksheet.Cells[xlWorksheet.Rows.Count, "A"].End(XlDirection.xlUp).Row;
            int lRow = xlWorksheet.Range["A" + xlWorksheet.Rows.Count.ToString()].End[XlDirection.xlUp].Row;
            Range range = xlWorksheet.Range["A9:P" + lRow.ToString()];
            object[,] data = range.Value;
            for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                f2.progressBar1.Maximum = range.Columns.Count;
                f2.progressBar1.Value = cCnt;
                f2.label2.Text = "Fetching data from excel . . .";
                f2.Refresh();
                DataColumn Column = new DataColumn();
                Column.DataType = typeof(string);
                Column.ColumnName = cCnt.ToString();
                excelData.Columns.Add(Column);
                for (int rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {
                    string rCnt1 = Convert.ToString(rCnt);
                    string cCnt1 = Convert.ToString(cCnt);
                    string CellVal = string.Empty;
                    CellVal = Convert.ToString(data[rCnt, cCnt]);
                    DataRow Row;
                    // Adds the row to the DataTable           
                    if (cCnt == 1)
                    {
                        Row = excelData.NewRow();
                        Row[Column.ColumnName.ToString()] = CellVal;
                        excelData.Rows.Add(Row);
                    }
                    else
                    {
                        Row = excelData.Rows[rCnt - 1];
                        Row[Column.ColumnName.ToString()] = CellVal;
                    }
                }
            }
            excelData = excelData.DefaultView.ToTable(false, "1", "4", "16"); // selects only the columns needed
            DataRow[] prt = excelData.Select("[1] is null or [1] = ''"); // selects the rows to be deleted
                                                                         // deletes the rows
            foreach (DataRow prt2 in prt)
            {
                prt2.Delete();
            }
            excelData.AcceptChanges(); // applies the changes made to the datatable
            invTblCopy = excelData.Copy();
            f2.Close();
            xlWorkbook.Close();
            MessageBox.Show("Import " + filename + " successfully!!", "Success!!");
            try
            {
                xlApp.Quit();
                oExcel2.Quit(); // quits the excel application
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            // release or dispose the objects
            ReleaseObject(lRow);
            ReleaseObject(xlWorksheet);
            ReleaseObject(xlWorkbook);
            ReleaseObject(oExcel2);
        }
        private void Inventory_saveRev()
        {
            Form2 f2 = new Form2();
            inventoryTable = new DataTable(); // creates a new instance of the datatable
            inventoryTable = excelData.Clone(); // clones the column of the datatable
            invSMorNonTbl = new DataTable(); // creates a new instance of the datatable
            DataTable distinctDT = new DataTable(); // creates a datatable for filtering partnumber
            distinctDT = pullQuanTempTable.DefaultView.ToTable(true, "PARTNUMBER"); // selects distinct partnumbers
            for (int a = 0; a < distinctDT.Rows.Count; a++)
            {
                f2.progressBar1.Maximum = distinctDT.Rows.Count;
                f2.progressBar1.Value = a;
                f2.label2.Text = "Computing stocks . . .";
                f2.Refresh();
                DataView negaView = new DataView(excelData);
                negaView.RowFilter = "[4] LIKE '%" + distinctDT.Rows[a][0] + "%' or [4] = '" + distinctDT.Rows[a][0] + "'"; // selects rows base on the partnumber
                forFilter = new DataTable(); // creates a new instance of the datatable
                forFilter = excelData.Clone(); // clones the columns of the datatable
                forFilter = negaView.ToTable(); // stores the selected view or rows to the datatable
                foreach (DataRow row in forFilter.Rows)
                {
                    inventoryTable.ImportRow(row); // stores each row to the datatable
                    inventoryTable.AcceptChanges(); // applies changes to the datatable
                }
                
            }
            // checking of bin card
            invTblCopy = new DataTable(); // creates a copy of the datatable
            invTblCopy = inventoryTable.Copy();
            string delInv = "[1] LIKE '%-1' or [1] LIKE '%-2' or [16] = '0'"; // selects rows to be removed
            DataRow[] delInvRow;
            delInvRow = invTblCopy.Select(delInv);
            //deletes the selected rows
            foreach (DataRow delInvRow2 in delInvRow)
            {
                delInvRow2.Delete();
            }
            invTblCopy.AcceptChanges(); // applies the changes to the datatable
            DataTable invCopyTwo = new DataTable(); // clones the columns of the datatable
            invCopyTwo = invTblCopy.Clone();
            //groups the rows that have the same value and gets the sum
            var GroupedSum1 = invTblCopy.AsEnumerable()
                 .Where(row => row.RowState != DataRowState.Deleted)
                 .GroupBy(userEntry => new { key = userEntry.Field<string>("1"), key2 = userEntry.Field<string>("4") })
                 .Select(grp => new
                 {
                     IdKey = grp.Key.key,
                     IdKey2 = grp.Key.key2,
                     Summary = grp.Sum(p => p["16"] != DBNull.Value ? Convert.ToInt32(p["16"]) : 0)
                 });
            // stores the grouped rows into a datatable
            foreach (var Col in GroupedSum1.ToList())
            {
                invCopyTwo.Rows.Add(Col.IdKey, Col.IdKey2, Col.Summary);
            }
            invTblCopy = new DataTable();
            invTblCopy = invCopyTwo.Copy();
            invTblCopy.DefaultView.Sort = "[1] DESC"; // sorts the datatable to place the rows with partnumbers containing "SM" above
            invTblCopy = invTblCopy.DefaultView.ToTable(); // saves the new view of the datatable
            invTblCopy.AcceptChanges(); // applies the changes to the datatable
            int r = 0;
            int lessStock;
            DataRow[] z2;
            r = 0;
            int unpostBal = 0;
            for (int up = 0; up < unpostedTbl.Rows.Count; up++)
            {
                f2.progressBar1.Maximum = unpostedTbl.Rows.Count;
                f2.label2.Text = "Computing stocks . . .";
                r++;
                f2.progressBar1.Value = r;
                f2.Refresh();
                string expr2 = $"[1] = '" + unpostedTbl.Rows[up][1] + "'";
                z2 = inventoryTable.Select(expr2); // selects the rows in the inventory datatable
                unpostBal = 0;
                // added for A119 and 211
                // computes the total remaining stocks
                if (z2.Length > 0)
                {
                    if (z2.Length > 1) // if the inventory has more than one partnumber (SM and non-SM)
                    {
                        lessStock = 0;
                        unpostBal = 0;
                        for (int a = 0; a < z2.Length; a++)
                        {
                            if (a == 0)
                            {
                                lessStock = int.Parse(z2[a][2].ToString()) - int.Parse(unpostedTbl.Rows[up][3].ToString()); // subtracts the unposted quantity to the remaining stocks of the first partnumber
                            }
                            else
                            {
                                unpostBal = int.Parse(unpostedTbl.Rows[up][3].ToString()); // sets the unposted quantity to the next partnumber
                                lessStock = int.Parse(z2[a][2].ToString()) - unpostBal; // subtracts the unposted quantity to the remaining stocks of the next partnumber
                            }
                            if (lessStock < 0)
                            {
                                unpostBal = Math.Abs(lessStock); // gets the remaining unposted summary to be subtracted
                                lessStock = 0;
                                z2[a].SetField(2, lessStock); // saves the value of the stock in the inventory datatable
                            }
                            else if (lessStock == 0)
                            {
                                z2[a].SetField(2, lessStock); // saves the value of the stock in the inventory datatable
                                break; // exits the loop since the stocks are sufficient
                            }
                            else
                            {
                                z2[a].SetField(2, lessStock); // saves the value of the stock in the inventory datatable
                                break; // exits the loop since the stocks are sufficient
                            }
                        }
                    }
                    else
                    {
                        lessStock = (z2[0][2] != DBNull.Value ? Convert.ToInt32(z2[0][2]) : 0) -
            (unpostedTbl.Rows[up][3] != DBNull.Value ? Convert.ToInt32(unpostedTbl.Rows[up][3]) : 0); // subtracts the unposted quantity to the remaining stocks in the inventory datatable
                        if (lessStock < 0)
                        {
                            lessStock = 0;
                        }
                        z2[0].SetField(2, lessStock); // saves the value of the stock in the inventory datatable
                    }
                }
            }
            string backexp = "[1] LIKE '%-1' or [1] LIKE '%-2' or [16] = '0'"; // selects the rows to be removed
            DataRow[] backrow;
            // selects the rows to be removed
            backrow = inventoryTable.Select(backexp);
            // deletes the rows selected
            foreach (DataRow backrow2 in backrow)
            {
                backrow2.Delete();
            }
            inventoryTable.AcceptChanges();
            DataTable invUniqueTbl = new DataTable(); // creates a new datatable to hold some data
            invUniqueTbl = inventoryTable.Clone(); // clones the columns of the datatable
            // groups the rows with the same DRAWING NUMBER and gets the sum
            var GroupedSum = from userEntry in inventoryTable.AsEnumerable()
                             group userEntry by new { key = userEntry.Field<string>("1"), key2 = userEntry.Field<string>("4") } into groupEntry
                             select new
                             {
                                 IdKey = groupEntry.Key.key,
                                 IdKey2 = groupEntry.Key.key2,
                                 Summary = groupEntry.Sum(p => p["16"] != DBNull.Value ? Convert.ToInt32(p["16"]) : 0)
                             };
            // stores the grouped rows to a datatable
            foreach (var col in GroupedSum.ToList())
            {
                invUniqueTbl.Rows.Add(col.IdKey, col.IdKey2, col.Summary);
            }
            // creates a new instance of the datatable
            inventoryTable = new DataTable();
            inventoryTable = invUniqueTbl.Copy(); // copies the datatable
            // sorts the datatable to place the partnumber with "SM" at the top
            inventoryTable.DefaultView.Sort = "[1] DESC";
            inventoryTable = inventoryTable.DefaultView.ToTable();
            inventoryTable.AcceptChanges();
            invSMorNonTbl = inventoryTable.Copy(); // copies the inventory datatable
        }
        //import axm432 button
        private void btnaxmr432_Click(object sender, EventArgs e)
        {
            if (btnaimr407.BackColor == Color.Navy)
            {
                string strFileName;
                openFD.InitialDirectory = "'C:\'";
                openFD.Filter = "Excel Office | *.xlsx; *.xls";
                openFD.Title = "Choose a File";
                openFD.FilterIndex = 2;
                openFD.RestoreDirectory = true;
                if (openFD.ShowDialog().Equals(DialogResult.OK))
                {
                    strFileName = openFD.FileName;
                    ptkFileName.Text = openFD.FileName;
                    if (!string.IsNullOrEmpty(FileSystem.Dir(openFD.FileName)))
                    {
                        if (!openFD.SafeFileName.ToString().ToUpper().Contains("AXMR432"))
                        {
                            Interaction.MsgBox("The file you are trying to import is named " + openFD.SafeFileName + Microsoft.VisualBasic.Constants.vbCrLf + Microsoft.VisualBasic.Constants.vbCrLf + "Make sure you are importing the correct file!");
                            return;
                        }
                        else
                        {
                            if (strFileName != "")
                            {
                                excelData3 = new DataTable();
                                poTable = new DataTable();
                                Axmr432ONE(strFileName, openFD.SafeFileName);
                                Axmr432TWO();
                                btnaxmr432.BackColor = System.Drawing.Color.Navy;
                                dataGridView4.Visible = true;
                                pulltktgrid.Visible = false;
                                slider.Visible = false;
                                refesherOrb.Visible = true;
                                importPO = true;
                                btnaxmr432.Enabled = false;
                                btncmpl.Visible = true;
                            }
                        }
                    }

                }
            }
            else
            {
                MessageBox.Show("Import AIMR407 first!", "Warning!", MessageBoxButtons.OK);
            }
        }
        //Method for importing AXMR432
        private void Axmr432ONE(string myfiledirect, string filename)
        {
            excelData3 = new DataTable();
            Form2 f2 = new Form2();
            f2.Show();
            oExcel2 = new Excel.Application();
            oExcel2.DisplayAlerts = false;
            xlWorkbook = oExcel2.Workbooks.Open(myfiledirect, XlFileAccess.xlReadOnly);
            xlWorksheet = xlWorkbook.Worksheets[1];
            int lRow = xlWorksheet.Cells[xlWorksheet.Rows.Count, "A"].End(XlDirection.xlUp).Row;
            //int lRow = xlWorksheet.Range["A" + xlWorksheet.Rows.Count.ToString()].End[XlDirection.xlUp].Row;
            Range range = xlWorksheet.Range["A7:K" + lRow.ToString()];
            object[,] data = range.Value;
            int r = 0;
            for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                f2.progressBar1.Maximum = range.Columns.Count;
                f2.progressBar1.Value = cCnt;
                f2.label2.Text = "Reading data from " + filename + " . . .";
                f2.Refresh();
                DataColumn Column = new DataColumn();
                Column.DataType = typeof(string);
                Column.ColumnName = cCnt.ToString();
                excelData3.Columns.Add(Column);
                for (int rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {
                    string rCnt1 = Convert.ToString(rCnt);
                    string cCnt1 = Convert.ToString(cCnt);
                    string CellVal = string.Empty;
                    CellVal = Convert.ToString(data[rCnt, cCnt]);
                    DataRow Row;
                    // Adds the row to the DataTable           
                    if (cCnt == 1)
                    {
                        Row = excelData3.NewRow();
                        Row[Column.ColumnName.ToString()] = CellVal;
                        excelData3.Rows.Add(Row);
                    }
                    else
                    {
                        Row = excelData3.Rows[rCnt - 2];
                        Row[Column.ColumnName.ToString()] = CellVal;
                    }
                }
            }
            excelData3 = excelData3.DefaultView.ToTable(false, "1", "2", "8", "9", "11", "7"); // selects only the columns needed
            DataRow[] prt = excelData3.Select("[8] is null or [8] LIKE '%SO%'"); // selects the rows to be deleted
                                                                                 // deletes the rows
            foreach (DataRow prt2 in prt)
            {
                prt2.Delete();
            }
            excelData3.AcceptChanges(); // applies the changes made to the datatable
            MessageBox.Show("Import " + filename + " successfully!!", "Success!!");
            xlWorkbook.Close();
            try
            {
                f2.label2.Text = "Process complete!";
                f2.Refresh();
                f2.Close();
                xlApp.Quit(); // quits the excel application
                oExcel2.Quit();
            }
            catch (Exception ex)
            {
                return;
            }
            // release or dispose the objects
            ReleaseObject(lRow);
            ReleaseObject(xlWorksheet);
            ReleaseObject(xlWorkbook);
            ReleaseObject(oExcel2);
        }
        private void Axmr432TWO()
        {
            Form2 f2 = new Form2();
            poTable = new DataTable();
            this.colorPOsm = new DataTable();
            DataColumn pullQTcol = poTable.Columns.Add("PRODUCT_NO", typeof(string));
            poTable.Columns.Add("NAME", typeof(string));
            poTable.Columns.Add("SO_NO", typeof(string));
            poTable.Columns.Add("SO_QTY", typeof(string));
            poTable.Columns.Add("OPEN_QTY", typeof(Int32));
            poTable.Columns.Add("CUST_NO", typeof(string));
            poTable.Columns.Add("CUST_NOLAST", typeof(Int32)); // added nov 23 
            DataTable int1 = new DataTable(); // creates a new datatable to store distinct partnumbers from the pullticket
            int1 = pullQuanTempTable.DefaultView.ToTable(true, "PARTNUMBER"); // select distinct partnumbers from pullticket
            int r = 0;
            DataRow[] z2;
            for (int i = 0; i < int1.Rows.Count; i++) // loops through the partnumbers
            {
                f2.progressBar1.Maximum = int1.Rows.Count;
                f2.label2.Text = "Calculating P.O....";
                f2.Refresh();
                f2.progressBar1.Value = i;
                // selects rows from the imported excel file datatable
                string expr = "[2] LIKE '%" + int1.Rows[i]["PARTNUMBER"].ToString() + " %' or [2] = '" + int1.Rows[i]["PARTNUMBER"].ToString() + "'";
                DataRow[] z;
                z = excelData3.Select(expr);
                for (int v = 0; v < z.Length; v++) // loops through the rows
                {
                    DataRow pq = poTable.NewRow(); // creates a new row for the PO datatable
                    pq["PRODUCT_NO"] = z[v][0];
                    pq["NAME"] = z[v][1];
                    pq["SO_NO"] = z[v][2];
                    pq["SO_QTY"] = z[v][3];
                    pq["OPEN_QTY"] = z[v][4];
                    /*int openQtyValue;
                    if (int.TryParse(z[v][4].ToString(), out openQtyValue))
                    {
                        pq["OPEN_QTY"] = openQtyValue;
                    }
                    else
                    {
                        pq["OPEN_QTY"] = 0; // or some default value
                    }*/
                    string custNo = z[v][5].ToString();
                    int custNoLength = custNo.Length;
                    pq["CUST_NO"] = custNo.Substring(0, Math.Min(7, custNoLength));
                    if (custNoLength > 7)
                    {
                        pq["CUST_NOLAST"] = custNo.Substring(7, custNoLength - 7);
                    }
                    else
                    {
                        pq["CUST_NOLAST"] = 0;
                    }
                    poTable.Rows.Add(pq);
                }
                poTable.AcceptChanges();
            }
            r = 0;
            // computes the total remaining PO
            DataRow[] z3;
            if (poTable.Rows.Count < unpostedTbl.Rows.Count) // checks which datatable has lower number of rows for faster looping
            {
                for (int a = 0; a < poTable.Rows.Count; a++) // loops through the PO datatable
                {
                    f2.progressBar1.Maximum = poTable.Rows.Count;
                    f2.label2.Text = "Calculating P.O....";
                    f2.Refresh();
                    r++;
                    f2.progressBar1.Value = r;
                    // checks the unposted datatable for same SO number
                    string expr2 = "[1] = '" + poTable.Rows[a]["SO_NO"] + "' and [2] ='" + poTable.Rows[a]["SO_QTY"] + "'";
                    z3 = unpostedTbl.Select(expr2);
                    for (int b = 0; b < z3.Length; b++) // loops through rows and subtracts the unposted qty to the QTY of PO
                    {
                        poTable.Rows[a]["OPEN_QTY"] = Convert.ToInt32(poTable.Rows[a]["OPEN_QTY"]) - Convert.ToInt32(z3[b][3]);
                    }
                }
            }
            else
            {
                for (int a = 0; a < unpostedTbl.Rows.Count; a++) // loops through the unposted datatable
                {
                    f2.progressBar1.Maximum = unpostedTbl.Rows.Count;
                    f2.label2.Text = "Calculating P.O....";
                    r++;
                    f2.progressBar1.Value = r;
                    f2.Refresh();
                    // checks the SO datatable for same SO number
                    string expr2 = $"SO_NO = '{unpostedTbl.Rows[a][0]}' AND SO_QTY = '{unpostedTbl.Rows[a][4]}'";
                    z3 = poTable.Select(expr2);
                    if (z3.Length > 0) // subtracts the unposted qty to the QTY of PO
                    {
                        string openQtyStr = z3[0]["OPEN_QTY"].ToString();
                        string unpostedQtyStr = unpostedTbl.Rows[a][3].ToString();
                        if (!string.IsNullOrEmpty(openQtyStr) && !string.IsNullOrEmpty(unpostedQtyStr))
                        {
                            int openQtyValue, unpostedQtyValue;
                            if (int.TryParse(openQtyStr, out openQtyValue) && int.TryParse(unpostedQtyStr, out unpostedQtyValue))
                            {
                                z3[0]["OPEN_QTY"] = openQtyValue - unpostedQtyValue;
                                poTable.AcceptChanges();
                            }
                            else
                            {
                                // Handle the situation where the parsing fails
                                // You can log an error, throw an exception, or set a default value
                            }
                        }
                        else
                        {
                            // Handle the situation where the values are null or empty
                            // You can log an error, throw an exception, or set a default value
                        }
                    }
                }
            }
            // selects rows that has 0 open qty in PO datatable and deletes it
            string expr3 = "OPEN_QTY <= 0";
            DataRow[] prt;
            prt = poTable.Select(expr3);
            foreach (DataRow prt2 in prt)
            {
                prt2.Delete();
            }
            //deletes unnecessary product number
            string backexp = "PRODUCT_NO LIKE '%-1' or PRODUCT_NO LIKE '%-2'";
            DataRow[] backrow;
            backrow = poTable.Select(backexp);
            foreach (DataRow backrow2 in backrow)
            {
                backrow2.Delete();
            }
            poTable.AcceptChanges(); // apply changes to the datatable
            poTable.DefaultView.Sort = "PRODUCT_NO DESC, CUST_NO ASC, CUST_NOLAST ASC"; // sorts the datatable for product number that contains "SM" and places it on top
            poTable = poTable.DefaultView.ToTable();
            poTable.AcceptChanges();
            colorPOsm = poTable.Copy(); // creates a copy of PO datatable
        }
        //compile button
        private void cmpl_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            if (importPO != true)
            {
                MessageBox.Show("Import AXMR432 first!", "Warning");
            }
            if (calculateBtn == true)
            {
                return;
            }
            Filename = "Generated_Pull";
            f2.Show();
            f2.label2.Text = "Generating pull request details. . .";
            f2.Refresh();
            POFGDelete();
            GenerateStep1Final();
            GenerateStep2Final();
            consoBinding = new BindingSource();
            consoBinding.DataSource = conso_final_formatTable;
            DataGridView3.DataSource = consoBinding;
            //dataGridView4.DataSource = consoBinding;
            pull_request_trans = 1;
            if (genClick == false)
            {
                DataGridViewCheckBoxColumn addcolumn = new DataGridViewCheckBoxColumn();
                addcolumn.HeaderText = "CANCEL";
                addcolumn.Name = "cancelled";
                addcolumn.Width = 80;
                DataGridView3.Columns.Insert(0, addcolumn);
                DataGridView3.AllowUserToAddRows = false;
                for (int z = 0; z < DataGridView3.Rows.Count; z++) // loops through the DataGridView and selects the rows based on the given conditions
                {
                    if (int.Parse(DataGridView3.Rows[z].Cells["STOCK_QTY"].Value.ToString()) == 0 || int.Parse(DataGridView3.Rows[z].Cells["QUANTITY_DELIVERED"].Value.ToString()) == 0)
                    {
                        DataGridView3.Rows[z].Cells[0].Value = true;
                        DataGridView3.Rows[z].Cells[0].ReadOnly = false;
                    }
                    // checks the PO datatable to check if the SO number contains "SM"
                    string expr2 = "SO_NO = '" + DataGridView3.Rows[z].Cells["GO_NUMBER"].Value + "'";
                    DataRow[] z2 = colorPOsm.Select(expr2);
                    if (z2.Length > 0)
                    {
                        if (z2[0][0].ToString().Contains("SM"))
                        {
                            DataGridView3.Rows[z].DefaultCellStyle.BackColor = Color.FromArgb(255, 204, 153); // change the color of the entire row of the DataGridView
                            DataGridView3.Refresh();
                            dataGridView4.DataSource = null;
                            dataGridView4.DataSource = DataGridView3.DataSource;
                        }
                    }
                }
                genClick = true;
                f2.Close();
                btncmpl.BackColor = Color.Crimson;
                calculateBtn = true;
                btndownload.Visible = true;
            }
        }
        private void POFGDelete()
        {
            DataTable invDistinct = new DataTable(); // creates a new datatable for FG
            invDistinct = new DataTable();
            invDistinct = inventoryTable.Copy(); // copies the inventory datatable
            invDistinct = invDistinct.DefaultView.ToTable(true, "1", "4"); // selects distinct values in the columns
            DataTable poTableDistinct = new DataTable(); // creates a new datatable for PO
            poTableDistinct = poTable.Copy(); // copies the PO datatable
            poTableDistinct = poTableDistinct.DefaultView.ToTable(true, "PRODUCT_NO", "NAME"); // selects distinct values in the columns
            bool poExist = false;
            bool invExist = false;
            // ///////////////////////// deleting of FG /////////////////////////////////
            for (int a = 0; a < invDistinct.Rows.Count; a++) // loops through the copy of inventory datatable
            {
                invExist = false;
                for (int b = 0; b < poTableDistinct.Rows.Count; b++) // loops through the copy of PO datatable
                {
                    if (invDistinct.Rows[a][0].Equals(poTableDistinct.Rows[b]["PRODUCT_NO"])) // checks if they have the same product number
                    {
                        invExist = true;
                        break;
                    }
                }
                if (invExist == false) // if there is no data in PO datatable
                {
                    // selects the rows in inventory datatable and deletes it
                    string delPO = "[1] = '" + invDistinct.Rows[a][0] + "'";
                    DataRow[] delPOrow = inventoryTable.Select(delPO);

                    foreach (DataRow delPOrow2 in delPOrow)
                    {
                        delPOrow2.Delete();
                    }
                    inventoryTable.AcceptChanges();
                }
            }
            invDistinct = new DataTable(); // creates a new instance of the datatable
            invDistinct = inventoryTable.Copy(); // copies the updated inventory datatable
            invDistinct = invDistinct.DefaultView.ToTable(true, "1", "4"); // selects distinct values in the columns
            // ///////////////////////// deleting of PO /////////////////////////////////
            for (int a = 0; a < poTableDistinct.Rows.Count; a++) // loops through the copy of PO datatable
            {
                poExist = false;
                for (int b = 0; b < invDistinct.Rows.Count; b++) // loops through the copy of inventory datatable
                {
                    if (invDistinct.Rows[b][0].Equals(poTableDistinct.Rows[a]["PRODUCT_NO"])) // checks if they have the same product number
                    {
                        poExist = true;
                    }
                }
                if (poExist == false) // if there is no data in inventory datatable
                {
                    // selects the rows in PO datatable and deletes it
                    string delPO = "PRODUCT_NO = '" + poTableDistinct.Rows[a]["PRODUCT_NO"] + "'";
                    DataRow[] delPOrow = poTable.Select(delPO);

                    foreach (DataRow delPOrow2 in delPOrow)
                    {
                        delPOrow2.Delete();
                    }
                    poTable.AcceptChanges();
                }
            }
        }
        private void GenerateStep1Final()
        {
            Form2 f2 = new Form2();
            Filename = "Pull_Request"; // sets a new value for the string
            pullRevTable = new DataTable();
            consolidatedTable = new DataTable(); // creates a new instance of the datatable, this table will store the final data for printing
            poTableCopy = poTable.Copy(); // copies the PO datatable
            date_time = DateTime.Now.ToString("yyyy/MM/dd HH:mm:00");
            // adds columns to the datatable
            DataColumn pullQTcol = consolidatedTable.Columns.Add("TRANS_DATE", typeof(string));
            consolidatedTable.Columns.Add("PROD_DATE", typeof(string));
            consolidatedTable.Columns.Add("PROD_TIME", typeof(string));
            consolidatedTable.Columns.Add("DEL_DATE", typeof(string));
            consolidatedTable.Columns.Add("DEL_TIME", typeof(string));
            consolidatedTable.Columns.Add("JOB_NO", typeof(string));
            consolidatedTable.Columns.Add("FACILITY", typeof(string));
            consolidatedTable.Columns.Add("PARTNUMBER", typeof(string));
            consolidatedTable.Columns.Add("PULL_QTY", typeof(string));
            consolidatedTable.Columns.Add("STOCK_QUANTITY", typeof(string));
            consolidatedTable.Columns.Add("END_BALANCE", typeof(string));
            consolidatedTable.Columns.Add("QUANTITY_DELIVERED", typeof(string));
            consolidatedTable.Columns.Add("GO_NUMBER", typeof(string));
            consolidatedTable.Columns.Add("SKU_ASSEMBLY", typeof(string));
            consolidatedTable.Columns.Add("GO_LINE_NUMBER", typeof(string));
            consolidatedTable.Columns.Add("CELL_NUM", typeof(string));
            consolidatedTable.Columns.Add("REMARKS", typeof(string));
            consolidatedTable.Columns.Add("PULL_TICKET_NUMBER", typeof(string));
            consolidatedTable.Columns.Add("LINE", typeof(string));
            consolidatedTable.Columns.Add("VENDOR_REMARKS", typeof(string));
            consolidatedTable.Columns.Add("PULLTYPE", typeof(string));
            consolidatedTable.Columns.Add("PULLNO", typeof(int));
            consolidatedTable.Columns.Add("QTY_DEL", typeof(int));
            consolidatedTable.Columns.Add("ORIGINAL_PULL", typeof(int));
            // adds columns to the datatable
            DataColumn fgSMCol = fgSM.Columns.Add("PROD_NUM", typeof(string));
            fgSM.Columns.Add("STOCK", typeof(int));
            fgSM.Columns.Add("PULLNO", typeof(int));
            fgSM.Columns.Add("FGCOUNT", typeof(int));
            // adds columns to the datatable
            DataColumn totalpullSMcol = totalPullSM.Columns.Add("PROD_NUM", typeof(string));
            totalPullSM.Columns.Add("PULL", typeof(int));
            // /////////////////////////////////////////////////////////////////////////////
            int r = 0;
            int consoCounter = 0;
            DataRow[] z2;
            foreach (DataGridViewRow rw in pulltktgrid.Rows) // loops through the datagridview that contains the pull ticket data
            {
                // Try
                int EndBalance; // this is the total remaining stocks or FG
                int Quantity; // pull qty to use base on the value of endbalance
                string Item_No; // item number of FG
                int totalStock = 0; // //////////////// addev nov 20 for SM PO /////////////////////
                int SMquan = 0; // //////////////// addev nov 20 for SM PO /////////////////////
                int NonSMquan = 0; // //////////////// addev nov 20 for SM PO /////////////////////
                int totalSmPO = 0; // //////////////// addev nov 20 for SM PO /////////////////////
                int totalSMdel = 0; // //////////////// addev nov 20 for SM PO /////////////////////
                bool noSM, allSM, withSM; // //////////////// addev nov 20 for SM PO /////////////////////
                noSM = false;
                allSM = false;
                withSM = false;
                consoCounter++;
                if (rw.Cells[6].Value != null && rw.Cells[6].Value.ToString() != "")
                {
                    f2.progressBar1.Maximum = pulltktgrid.Rows.Count;
                    r++;
                    f2.progressBar1.Value = r;
                    f2.label2.Text = "Generating Pull Request Details . . .";
                    f2.Refresh();
                    // selects rows from inventory datatable based on partnumber
                    string expr2 = "[4] LIKE '%" + rw.Cells[6].Value + "%' or [4] = '" + rw.Cells[6].Value + "'";
                    z2 = inventoryTable.Select(expr2, "[1] DESC");
                    if (z2.Length == 0) //0 means there is no stocks or finished goods
                    {
                        EndBalance = 0 - Convert.ToInt32(rw.Cells[7].Value); // substracts the pull qty and save it as endbalance
                        // creates a new row in datatable
                        DataRow pq = consolidatedTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = rw.Cells[0].Value;
                        pq["PROD_TIME"] = rw.Cells[1].Value;
                        pq["DEL_DATE"] = rw.Cells[2].Value;
                        pq["DEL_TIME"] = rw.Cells[3].Value;
                        pq["JOB_NO"] = rw.Cells[4].Value;
                        pq["FACILITY"] = rw.Cells[5].Value;
                        pq["PARTNUMBER"] = rw.Cells[6].Value;
                        pq["PULL_QTY"] = rw.Cells[7].Value;
                        pq["STOCK_QUANTITY"] = 0;
                        pq["END_BALANCE"] = EndBalance;
                        pq["QUANTITY_DELIVERED"] = 0;
                        pq["GO_NUMBER"] = "";
                        pq["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                        pq["GO_LINE_NUMBER"] = "";
                        pq["CELL_NUM"] = rw.Cells[10].Value;
                        pq["REMARKS"] = "NO FG";
                        pq["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                        pq["LINE"] = rw.Cells[13].Value;
                        pq["VENDOR_REMARKS"] = rw.Cells[15].Value;
                        pq["PULLTYPE"] = rw.Cells[23].Value;
                        pq["PULLNO"] = consoCounter;
                        pq["QTY_DEL"] = rw.Cells[24].Value;
                        pq["ORIGINAL_PULL"] = rw.Cells[25].Value;
                        consolidatedTable.Rows.Add(pq);
                        consolidatedTable.AcceptChanges();
                        continue; // continue to the next datagridview row since there is no stocks or finished goods (FG)
                    }
                    if (z2.Length == 1) //1 means there is only one partnumber in inventory
                    {
                        if (z2[0][0].ToString().Contains("SM")) //checks if the item number contains "SM"
                        {
                            if (rw.Cells[12].Value.ToString().ToUpper().Contains("KANBAN"))
                            {
                                // checks if there is sufficient PO
                                string poCheck = "PRODUCT_NO = '" + z2[0][0] + "'"; // selects row in the PO datatable with the same product number
                                DataRow[] poCheckRow;
                                poCheckRow = poTable.Select(poCheck);
                                int totSMpo = 0;
                                for (int a = 0; a < poCheckRow.Length; a++)
                                {
                                    totSMpo = totSMpo + Convert.ToInt32(poCheckRow[a]["OPEN_QTY"]); // computes the total PO qty
                                }                                                 // ///////////////////////////////////////////////////////////////////////////////
                                if (Convert.ToInt32(z2[0][2]) < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty

                                    // adds a new row to the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING FG";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since FG is not sufficient
                                }
                                else if (totSMpo < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); ; // FG - pull qty
                                    // adds a new row to the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING PO";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                            }
                            // creates a new row in the datatable to store data with "SM"
                            DataRow pq = fgSM.NewRow();
                            pq["PROD_NUM"] = z2[0][0];
                            pq["STOCK"] = z2[0][2];
                            pq["PULLNO"] = consoCounter;
                            pq["FGCOUNT"] = 1;
                            fgSM.Rows.Add(pq);
                            fgSM.AcceptChanges();
                            // selects and deletes rows in PO datatable that are non-SM77777777777777
                            string delSMpo = "PRODUCT_NO = '" + z2[0][0].ToString().Replace("-SM", "") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-SM", "-3") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-SM", "-REWORK") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-SM", "-KST") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-SM", "-MOLEX") + "'";
                            DataRow[] delSMporow;
                            delSMporow = poTable.Select(delSMpo);
                            foreach (DataRow prt2 in delSMporow)
                            {
                                prt2.Delete();
                            }
                            poTable.AcceptChanges();
                        }
                        else if (z2[0][0].ToString().Contains("-3"))
                        {
                            if (rw.Cells[12].Value.ToString().ToUpper().Contains("KANBAN"))
                            {
                                // checks if there is sufficient PO
                                string poCheck = "PRODUCT_NO = '" + z2[0][0] + "'";
                                DataRow[] poCheckRow;
                                poCheckRow = poTable.Select(poCheck); // selects rows in PO datatable with the same product number
                                int totSMpo = 0;
                                for (int a = 0; a < poCheckRow.Length; a++)
                                {
                                    totSMpo = totSMpo + Convert.ToInt32(poCheckRow[a]["OPEN_QTY"]); // computes the total PO qty
                                } // computes the total PO qty
                                  // ///////////////////////////////////////////////////////////////////////////////

                                if (Convert.ToInt32(z2[0][2]) < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING FG";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                                else if (totSMpo < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING PO";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                            }
                            // selects and deletes the rows in PO datatable that don't contain "-3"
                            string delSMpo = "PRODUCT_NO = '" + z2[0][0].ToString().Replace("-3", "") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-3", "-SM") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-3", "-REWORK") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-3", "-KST") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-3", "-MOLEX") + "'";
                            DataRow[] delSMporow;
                            delSMporow = poTable.Select(delSMpo);
                            foreach (DataRow prt2 in delSMporow)
                            {
                                prt2.Delete();
                            }
                            poTable.AcceptChanges();
                        }
                        else if (z2[0][0].ToString().Contains("-REWORK"))
                        {
                            if (rw.Cells[12].Value.ToString().ToUpper().Contains("KANBAN"))
                            {
                                // checks if there is sufficient PO
                                string poCheck = "PRODUCT_NO = '" + z2[0][0] + "'";
                                DataRow[] poCheckRow;
                                poCheckRow = poTable.Select(poCheck); // selects rows in PO datatable with the same product number
                                int totSMpo = 0;
                                for (int a = 0; a < poCheckRow.Length; a++)
                                {
                                    totSMpo += Convert.ToInt32(poCheckRow[a]["OPEN_QTY"]); // computes the total PO qty
                                }                                                 // ///////////////////////////////////////////////////////////////////////////////
                                if (Convert.ToInt32(z2[0][2]) < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(rw.Cells[7].Value) - Convert.ToInt32(z2[0][2]); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING FG";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                                else if (totSMpo < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING PO";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                            }
                            // selects and deletes the rows in PO datatable that don't contain "-REWORK"
                            string delSMpo = "PRODUCT_NO = '" + z2[0][0].ToString().Replace("-REWORK", "") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-REWORK", "-SM") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-REWORK", "-3") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-REWORK", "-KST") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-REWORK", "-MOLEX") + "'";
                            DataRow[] delSMporow;
                            delSMporow = poTable.Select(delSMpo);
                            foreach (DataRow prt2 in delSMporow)
                            {
                                prt2.Delete();
                            }
                            poTable.AcceptChanges();
                        }
                        else if (z2[0][0].ToString().Contains("-KST"))
                        {
                            if (rw.Cells[12].Value.ToString().ToUpper().Contains("KANBAN"))
                            {
                                // checks if there is sufficient PO
                                string poCheck = "PRODUCT_NO = '" + z2[0][0] + "'";
                                DataRow[] poCheckRow;
                                poCheckRow = poTable.Select(poCheck); // selects rows in PO datatable with the same product number
                                int totSMpo = 0;
                                for (int a = 0; a < poCheckRow.Length; a++)
                                {
                                    totSMpo = totSMpo + Convert.ToInt32(poCheckRow[a]["OPEN_QTY"]); // computes the total PO qty
                                }                                              // ///////////////////////////////////////////////////////////////////////////////
                                if (Convert.ToInt32(z2[0][2]) < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING FG";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                                else if (totSMpo < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING PO";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                            }
                            // selects and deletes the rows in PO datatable that don't contain "-KST"
                            string delSMpo = "PRODUCT_NO = '" + z2[0][0].ToString().Replace("-KST", "") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-KST", "-SM") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-KST", "-3") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-KST", "-REWORK") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-KST", "-MOLEX") + "'";
                            DataRow[] delSMporow;
                            delSMporow = poTable.Select(delSMpo);
                            foreach (DataRow prt2 in delSMporow)
                            {
                                prt2.Delete();
                            }
                            poTable.AcceptChanges();
                        }
                        else if (z2[0][0].ToString().Contains("-MOLEX"))
                        {
                            if (rw.Cells[12].Value.ToString().ToUpper().Contains("KANBAN"))
                            {
                                // checks if there is sufficient PO
                                string poCheck = "PRODUCT_NO = '" + z2[0][0] + "'";
                                DataRow[] poCheckRow;
                                poCheckRow = poTable.Select(poCheck); // selects rows in PO datatable with the same product number
                                int totSMpo = 0;
                                for (int a = 0; a < poCheckRow.Length; a++)
                                {
                                    totSMpo = totSMpo + Convert.ToInt32(poCheckRow[a]["OPEN_QTY"]); // computes the total PO qty
                                }                                                // ///////////////////////////////////////////////////////////////////////////////
                                if (Convert.ToInt32(z2[0][2]) < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING FG";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                                else if (totSMpo < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING PO";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                            }
                            // selects and deletes the rows in PO datatable that don't contain "-MOLEX"
                            string delSMpo = "PRODUCT_NO = '" + z2[0][0].ToString().Replace("-MOLEX", "") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-MOLEX", "-SM") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-MOLEX", "-3") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-MOLEX", "-REWORK") + "' or PRODUCT_NO = '" + z2[0][0].ToString().Replace("-MOLEX", "-KST") + "'";
                            DataRow[] delSMporow;
                            delSMporow = poTable.Select(delSMpo);
                            foreach (DataRow prt2 in delSMporow)
                            {
                                prt2.Delete();
                            }
                            poTable.AcceptChanges();
                        }
                        else
                        {
                            if (rw.Cells[12].Value.ToString().ToUpper().Contains("KANBAN"))
                            {
                                // checks if there is sufficient PO
                                string poCheck = "PRODUCT_NO = '" + z2[0][0] + "'";
                                DataRow[] poCheckRow;
                                poCheckRow = poTable.Select(poCheck); // selects rows in PO datatable with the same product number
                                int totSMpo = 0;
                                for (int a = 0; a < poCheckRow.Length; a++)
                                {
                                    totSMpo = totSMpo + Convert.ToInt32(poCheckRow[a]["OPEN_QTY"]); // computes the total PO qty
                                }                                                 // ///////////////////////////////////////////////////////////////////////////////
                                if (Convert.ToInt32(z2[0][2]) < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING FG";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                                else if (totSMpo < Convert.ToInt32(rw.Cells[7].Value))
                                {
                                    EndBalance = Convert.ToInt32(z2[0][2]) - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                    // creates a new row in the datatable
                                    DataRow pq2 = consolidatedTable.NewRow();
                                    pq2["TRANS_DATE"] = date_time;
                                    pq2["PROD_DATE"] = rw.Cells[0].Value;
                                    pq2["PROD_TIME"] = rw.Cells[1].Value;
                                    pq2["DEL_DATE"] = rw.Cells[2].Value;
                                    pq2["DEL_TIME"] = rw.Cells[3].Value;
                                    pq2["JOB_NO"] = rw.Cells[4].Value;
                                    pq2["FACILITY"] = rw.Cells[5].Value;
                                    pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                    pq2["PULL_QTY"] = rw.Cells[7].Value;
                                    pq2["STOCK_QUANTITY"] = z2[0][2];
                                    pq2["END_BALANCE"] = EndBalance;
                                    pq2["QUANTITY_DELIVERED"] = 0;
                                    pq2["GO_NUMBER"] = "";
                                    pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                    pq2["GO_LINE_NUMBER"] = "";
                                    pq2["CELL_NUM"] = rw.Cells[10].Value;
                                    pq2["REMARKS"] = "LACKING PO";
                                    pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                    pq2["LINE"] = rw.Cells[13].Value;
                                    pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                    pq2["PULLTYPE"] = rw.Cells[23].Value;
                                    pq2["PULLNO"] = consoCounter;
                                    pq2["QTY_DEL"] = rw.Cells[24].Value;
                                    pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                    consolidatedTable.Rows.Add(pq2);
                                    consolidatedTable.AcceptChanges();
                                    continue; // continue to the next datagridview row since PO is not sufficient
                                }
                            }
                            // selects and deletes the rows in PO datatable that contains an extension name"
                            string delSMpo = "PRODUCT_NO = '" + z2[0][0] + "-SM' or PRODUCT_NO = '" + z2[0][0] + "-3' or PRODUCT_NO = '" + z2[0][0] + "-REWORK' or PRODUCT_NO = '" + z2[0][0] + "-MOLEX' or PRODUCT_NO = '" + z2[0][0] + "-KST'";
                            DataRow[] delSMporow;
                            delSMporow = poTable.Select(delSMpo);
                            foreach (DataRow prt2 in delSMporow)
                            {
                                prt2.Delete();
                            }
                            poTable.AcceptChanges();
                        }
                        totalStock = Convert.ToInt32(z2[0][2]);
                        EndBalance = totalStock - Convert.ToInt32(rw.Cells[7].Value); // remaing stocks after subtracting the pull qty
                        Item_No = z2[0][0].ToString(); // item number of the pull
                        if (EndBalance < 0)
                        {
                            Quantity = totalStock; // pull qty will be equal to the total stocks since end balance is less than 0
                        }
                        else
                        {
                            Quantity = Convert.ToInt32(rw.Cells[7].Value);
                            // pull qty will be the same since FG is sufficient
                        }
                        string newRemarks; // remarks for the pull ticket data
                        if (totalStock == 0)
                        {
                            newRemarks = "NO FG"; // remarks used to indicate there is no FG or stocks available
                        }
                        else
                        {
                            if (rw.Cells[11].Value == DBNull.Value)
                            {
                                newRemarks = "-";
                            }
                            else
                            {
                                newRemarks = rw.Cells[11].Value.ToString();
                            }
                        }

                        // creates a new row in the datatable
                        DataRow pq1 = consolidatedTable.NewRow();
                        pq1["TRANS_DATE"] = date_time;
                        pq1["PROD_DATE"] = rw.Cells[0].Value;
                        pq1["PROD_TIME"] = rw.Cells[1].Value;
                        pq1["DEL_DATE"] = rw.Cells[2].Value;
                        pq1["DEL_TIME"] = rw.Cells[3].Value;
                        pq1["JOB_NO"] = rw.Cells[4].Value;
                        pq1["FACILITY"] = rw.Cells[5].Value;
                        pq1["PARTNUMBER"] = rw.Cells[6].Value;
                        pq1["PULL_QTY"] = rw.Cells[7].Value;
                        pq1["STOCK_QUANTITY"] = totalStock;
                        pq1["END_BALANCE"] = EndBalance;
                        pq1["QUANTITY_DELIVERED"] = Quantity;
                        pq1["GO_NUMBER"] = "";
                        pq1["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                        pq1["GO_LINE_NUMBER"] = "";
                        pq1["CELL_NUM"] = rw.Cells[10].Value;
                        pq1["REMARKS"] = newRemarks;
                        pq1["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                        pq1["LINE"] = rw.Cells[13].Value;
                        pq1["VENDOR_REMARKS"] = rw.Cells[15].Value;
                        pq1["PULLTYPE"] = rw.Cells[23].Value;
                        pq1["PULLNO"] = consoCounter;
                        pq1["QTY_DEL"] = rw.Cells[24].Value;
                        pq1["ORIGINAL_PULL"] = rw.Cells[25].Value;
                        consolidatedTable.Rows.Add(pq1);
                        consolidatedTable.AcceptChanges();
                        if (EndBalance <= 0)
                        {
                            foreach (DataRow backrow2 in z2)
                            {
                                backrow2.Delete();
                            }
                            inventoryTable.AcceptChanges();
                        }
                        else
                        {
                            z2[0][2] = EndBalance; // updates the value of FG in the inventory datatable
                            inventoryTable.AcceptChanges();
                        }
                        continue;
                    }
                    else if (z2.Length > 1)
                    {
                        // //////////////////////////////////// for KANBAN ////////////////////////////////////
                        if (rw.Cells[12].Value.ToString().ToUpper().Contains("KANBAN"))
                        {
                            // checks if there is sufficient PO
                            string poCheck = "NAME LIKE '%" + rw.Cells[6].Value + " %' or NAME = '" + rw.Cells[6].Value + "'";
                            DataRow[] poCheckRow;
                            poCheckRow = poTable.Select(poCheck); // selects rows in PO datatable with the same product number
                            int totSMpo = 0;
                            int totSMfg = 0;
                            for (int a = 0; a < poCheckRow.Length; a++)
                            {
                                totSMpo = totSMpo + Convert.ToInt32(poCheckRow[a]["OPEN_QTY"]); // computes the total PO qty
                            }
                            for (int a = 0; a < z2.Length; a++)
                            {
                                totSMfg = totSMfg + int.Parse(z2[a][2].ToString()); // computes the total FG/stocks
                            }                                // ///////////////////////////////////////////////////////////////////////////////
                            if (totSMfg < int.Parse(rw.Cells[7].Value.ToString()))
                            {
                                EndBalance = totSMfg - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                // creates a new row in the datatable
                                DataRow pq2 = consolidatedTable.NewRow();
                                pq2["TRANS_DATE"] = date_time;
                                pq2["PROD_DATE"] = rw.Cells[0].Value;
                                pq2["PROD_TIME"] = rw.Cells[1].Value;
                                pq2["DEL_DATE"] = rw.Cells[2].Value;
                                pq2["DEL_TIME"] = rw.Cells[3].Value;
                                pq2["JOB_NO"] = rw.Cells[4].Value;
                                pq2["FACILITY"] = rw.Cells[5].Value;
                                pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                pq2["PULL_QTY"] = rw.Cells[7].Value;
                                pq2["STOCK_QUANTITY"] = totSMfg;
                                pq2["END_BALANCE"] = EndBalance;
                                pq2["QUANTITY_DELIVERED"] = 0;
                                pq2["GO_NUMBER"] = "";
                                pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                pq2["GO_LINE_NUMBER"] = "";
                                pq2["CELL_NUM"] = rw.Cells[10].Value;
                                pq2["REMARKS"] = "LACKING FG";
                                pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                pq2["LINE"] = rw.Cells[13].Value;
                                pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                pq2["PULLTYPE"] = rw.Cells[23].Value;
                                pq2["PULLNO"] = consoCounter;
                                pq2["QTY_DEL"] = rw.Cells[24].Value;
                                pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                consolidatedTable.Rows.Add(pq2);
                                consolidatedTable.AcceptChanges();
                                continue; // continue to the next datagridview row
                            }
                            else if (totSMpo < Convert.ToInt32(rw.Cells[7].Value))
                            {
                                EndBalance = totSMfg - Convert.ToInt32(rw.Cells[7].Value); // FG - pull qty
                                // creates a new row in the datatable
                                DataRow pq2 = consolidatedTable.NewRow();
                                pq2["TRANS_DATE"] = date_time;
                                pq2["PROD_DATE"] = rw.Cells[0].Value;
                                pq2["PROD_TIME"] = rw.Cells[1].Value;
                                pq2["DEL_DATE"] = rw.Cells[2].Value;
                                pq2["DEL_TIME"] = rw.Cells[3].Value;
                                pq2["JOB_NO"] = rw.Cells[4].Value;
                                pq2["FACILITY"] = rw.Cells[5].Value;
                                pq2["PARTNUMBER"] = rw.Cells[6].Value;
                                pq2["PULL_QTY"] = rw.Cells[7].Value;
                                pq2["STOCK_QUANTITY"] = totSMfg;
                                pq2["END_BALANCE"] = EndBalance;
                                pq2["QUANTITY_DELIVERED"] = 0;
                                pq2["GO_NUMBER"] = "";
                                pq2["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                                pq2["GO_LINE_NUMBER"] = "";
                                pq2["CELL_NUM"] = rw.Cells[10].Value;
                                pq2["REMARKS"] = "LACKING PO";
                                pq2["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                                pq2["LINE"] = rw.Cells[13].Value;
                                pq2["VENDOR_REMARKS"] = rw.Cells[15].Value;
                                pq2["PULLTYPE"] = rw.Cells[23].Value;
                                pq2["PULLNO"] = consoCounter;
                                pq2["QTY_DEL"] = rw.Cells[24].Value;
                                pq2["ORIGINAL_PULL"] = rw.Cells[25].Value;
                                consolidatedTable.Rows.Add(pq2);
                                consolidatedTable.AcceptChanges();
                                continue; // continue to the next datagridview row
                            }
                        }
                        // ///////////////////////////////////////////////////////////////////////////////////
                        for (int a = 0; a < z2.Length; a++) // loops through the selected inventory datatable rows
                        {
                            // ////////////////////// for pairing total FG with total PO //////////////////
                            if (z2[a][0].ToString().ToUpper().Contains("-SM"))
                            {
                                string expr3 = "PRODUCT_NO = '" + z2[a][0] + "'";
                                DataRow[] z3 = poTable.Select(expr3); // selects rows in PO datatable with the same product number
                                int poTotal = 0;
                                for (int b = 0; b < z3.Length; b++)
                                {
                                    poTotal = poTotal + Convert.ToInt32(z3[b]["OPEN_QTY"]);// computes the total PO qty
                                }
                                if (Convert.ToInt32(z2[a][2]) > poTotal)
                                    z2[a][2] = poTotal; // change the value of FG equal to the total PO
                                else if (Convert.ToInt32(z2[a][2]) < poTotal)
                                {
                                    int poSum = 0;
                                    bool equalFGPO = false;
                                    for (int c = 0; c < z3.Length; c++) // loops through the selected PO datatable rows
                                    {
                                        if (equalFGPO == false)
                                        {
                                            poSum = poSum + Convert.ToInt32(z3[c]["OPEN_QTY"]); // computes the total PO qty
                                            if (poSum > Convert.ToInt32(z2[a][2]))
                                            {
                                                int overPO = poSum - Convert.ToInt32(z2[a][2]); // gets the remaining PO
                                                z3[c]["OPEN_QTY"] = Convert.ToInt32(z3[c]["OPEN_QTY"]) - overPO; // changes the value of current PO to match the total FG
                                                poTable.AcceptChanges();
                                                equalFGPO = true;
                                            }
                                            else if (poSum == Convert.ToInt32(z2[a][2]))
                                            {
                                                equalFGPO = true;
                                            }
                                        }
                                        else
                                        {
                                            string POsmGO = "SO_NO = '" + int.Parse(z3[c]["SO_NO"].ToString()) + "' and SO_QTY ='" + int.Parse(z3[c]["SO_QTY"].ToString()) + "'";
                                            DataRow[] POsmGORow2;
                                            POsmGORow2 = poTable.Select(POsmGO); // selects rows in PO datatable with the same SO number
                                            // deletes the rows in PO datatable
                                            foreach (DataRow POsmGORow2a in POsmGORow2)
                                            {
                                                POsmGORow2a.Delete();
                                            }
                                            poTable.AcceptChanges();
                                        }
                                    }
                                }
                                totalStock = totalStock + Convert.ToInt32(z2[a][2]); // computes the total stock/FG
                                // adds new row to the datatable for FG with "SM"
                                DataRow pq = fgSM.NewRow();
                                pq["PROD_NUM"] = z2[a][0];
                                pq["STOCK"] = totalStock;
                                pq["PULLNO"] = consoCounter;
                                pq["FGCOUNT"] = 1;
                                fgSM.Rows.Add(pq);
                                fgSM.AcceptChanges();
                            }
                            else if (z2[a][0].ToString().ToUpper().Contains("-3"))
                            {
                                string expr3 = "PRODUCT_NO = '" + z2[a][0] + "'";
                                DataRow[] z3 = poTable.Select(expr3); // selects rows in PO datatable with the same product number
                                int poTotal = 0;
                                for (int b = 0; b < z3.Length; b++)
                                {
                                    poTotal = poTotal + Convert.ToInt32(z3[b]["OPEN_QTY"]);// computes the total PO qty
                                }
                                if (Convert.ToInt32(z2[a][2]) > poTotal)
                                {
                                    z2[a][2] = poTotal; // change the value of FG equal to the total PO
                                }
                                else if (Convert.ToInt32(z2[a][2]) < poTotal)
                                {
                                    int poSum = 0;
                                    bool equalFGPO = false;
                                    for (int c = 0; c < z3.Length; c++) // loops through the selected PO datatable rows
                                    {
                                        if (equalFGPO == false)
                                        {
                                            poSum = poSum + int.Parse(z3[c]["OPEN_QTY"].ToString()); // computes the total PO qty
                                            if (poSum > int.Parse(z2[a][2].ToString()))
                                            {
                                                int overPO = poSum - int.Parse(z2[a][2].ToString()); // gets the remaining PO
                                                z3[c]["OPEN_QTY"] = int.Parse(z3[c]["OPEN_QTY"].ToString()) - overPO; // changes the value of current PO to match the total FG
                                                poTable.AcceptChanges();
                                                equalFGPO = true;
                                            }
                                            else if (poSum == int.Parse(z2[a][2].ToString()))
                                            {
                                                equalFGPO = true;
                                            }

                                        }
                                        else
                                        {
                                            string POsmGO = "SO_NO = '" + z3[c]["SO_NO"] + "' and SO_QTY ='" + z3[c]["SO_QTY"] + "'";
                                            DataRow[] POsmGORow2;
                                            POsmGORow2 = poTable.Select(POsmGO); // selects rows in PO datatable with the same SO number
                                            // deletes the rows in PO datatable
                                            foreach (DataRow POsmGORow2a in POsmGORow2)
                                            {
                                                POsmGORow2a.Delete();
                                            }
                                            poTable.AcceptChanges();
                                        }
                                    }
                                }
                                totalStock = totalStock + Convert.ToInt32(z2[a][2]); // computes the total stock/FG
                            }
                            else if (z2[a][0].ToString().ToUpper().Contains("-REWORK"))
                            {
                                string expr3 = "PRODUCT_NO = '" + z2[a][0] + "'";
                                DataRow[] z3 = poTable.Select(expr3); // selects rows in PO datatable with the same product number
                                int poTotal = 0;
                                for (int b = 0; b < z3.Length; b++)
                                {
                                    poTotal = poTotal + Convert.ToInt32(z3[b]["OPEN_QTY"]);// computes the total PO qty
                                }
                                if (Convert.ToInt32(z2[a][2]) > poTotal)
                                {
                                    z2[a][2] = poTotal; // change the value of FG equal to the total PO
                                }
                                else if (Convert.ToInt32(z2[a][2]) < poTotal)
                                {
                                    int poSum = 0;
                                    bool equalFGPO = false;
                                    for (int c = 0; c < z3.Length; c++) // loops through the selected PO datatable rows
                                    {
                                        if (equalFGPO == false)
                                        {
                                            poSum = poSum + Convert.ToInt32(z3[c]["OPEN_QTY"]); // computes the total PO qty
                                            if (poSum > Convert.ToInt32(z2[a][2]))
                                            {
                                                int overPO = poSum - Convert.ToInt32(z2[a][2]); // gets the remaining PO
                                                z3[c]["OPEN_QTY"] = Convert.ToInt32(z3[c]["OPEN_QTY"]) - overPO; // changes the value of current PO to match the total FG
                                                poTable.AcceptChanges();
                                                equalFGPO = true;
                                            }
                                            else if (poSum == Convert.ToInt32(z2[a][2]))
                                            {
                                                equalFGPO = true;
                                            }
                                        }
                                        else
                                        {
                                            string POsmGO = "SO_NO = '" + z3[c]["SO_NO"] + "' and SO_QTY ='" + z3[c]["SO_QTY"] + "'";
                                            DataRow[] POsmGORow2;
                                            POsmGORow2 = poTable.Select(POsmGO); // selects rows in PO datatable with the same SO number
                                            // deletes the rows in PO datatable
                                            foreach (DataRow POsmGORow2a in POsmGORow2)
                                            {
                                                POsmGORow2a.Delete();
                                            }
                                            poTable.AcceptChanges();
                                        }
                                    }
                                }
                                totalStock = totalStock + Convert.ToInt32(z2[a][2]); // computes the total stock/FG
                            }
                            else if (z2[a][0].ToString().ToUpper().Contains("-KST"))
                            {
                                string expr3 = "PRODUCT_NO = '" + z2[a][0] + "'";
                                DataRow[] z3 = poTable.Select(expr3); // selects rows in PO datatable with the same product number
                                int poTotal = 0;
                                for (int b = 0; b < z3.Length; b++)
                                {
                                    poTotal = poTotal + Convert.ToInt32(z3[b]["OPEN_QTY"]);// computes the total PO qty
                                }
                                if (Convert.ToInt32(z2[a][2]) > poTotal)
                                {
                                    z2[a][2] = poTotal; // change the value of FG equal to the total PO
                                }
                                else if (Convert.ToInt32(z2[a][2]) < poTotal)
                                {
                                    int poSum = 0;
                                    bool equalFGPO = false;
                                    for (int c = 0; c < z3.Length; c++) // loops through the selected PO datatable rows
                                    {
                                        if (equalFGPO == false)
                                        {
                                            poSum = poSum + Convert.ToInt32(z3[c]["OPEN_QTY"]); // computes the total PO qty
                                            if (poSum > Convert.ToInt32(z2[a][2]))
                                            {
                                                int overPO = poSum - Convert.ToInt32(z2[a][2]); // gets the remaining PO
                                                z3[c]["OPEN_QTY"] = Convert.ToInt32(z3[c]["OPEN_QTY"]) - overPO; // changes the value of current PO to match the total FG
                                                poTable.AcceptChanges();
                                                equalFGPO = true;
                                            }
                                            else if (poSum == Convert.ToInt32(z2[a][2]))
                                            {
                                                equalFGPO = true;
                                            }
                                        }
                                        else
                                        {
                                            string POsmGO = "SO_NO = '" + z3[c]["SO_NO"] + "' and SO_QTY ='" + z3[c]["SO_QTY"] + "'";
                                            DataRow[] POsmGORow2;
                                            POsmGORow2 = poTable.Select(POsmGO); // selects rows in PO datatable with the same SO number
                                            // deletes the rows in PO datatable
                                            foreach (DataRow POsmGORow2a in POsmGORow2)
                                            {
                                                POsmGORow2a.Delete();
                                            }
                                            poTable.AcceptChanges();
                                        }
                                    }
                                }
                                totalStock = totalStock + Convert.ToInt32(z2[a][2]); // computes the total stock/FG
                            }
                            else if (z2[a][0].ToString().ToUpper().Contains("-MOLEX"))
                            {
                                string expr3 = "PRODUCT_NO = '" + z2[a][0] + "'";
                                DataRow[] z3 = poTable.Select(expr3); // selects rows in PO datatable with the same product number
                                int poTotal = 0;
                                for (int b = 0; b < z3.Length; b++)
                                {
                                    poTotal = poTotal + Convert.ToInt32(z3[b]["OPEN_QTY"]);// computes the total PO qty
                                }
                                if (Convert.ToInt32(z2[a][2]) > poTotal)
                                {
                                    z2[a][2] = poTotal; // change the value of FG equal to the total PO
                                }
                                else if (Convert.ToInt32(z2[a][2]) < poTotal)
                                {
                                    int poSum = 0;
                                    bool equalFGPO = false;
                                    for (int c = 0; c < z3.Length; c++) // loops through the selected PO datatable rows
                                    {
                                        if (equalFGPO == false)
                                        {
                                            poSum = poSum + int.Parse(z3[c]["OPEN_QTY"].ToString()); // computes the total PO qty
                                            if (poSum > int.Parse(z2[a][2].ToString()))
                                            {
                                                int overPO = poSum - int.Parse(z2[a][2].ToString()); // gets the remaining PO
                                                z3[c]["OPEN_QTY"] = int.Parse(z3[c]["OPEN_QTY"].ToString()) - overPO; // changes the value of current PO to match the total FG
                                                poTable.AcceptChanges();
                                                equalFGPO = true;
                                            }
                                            else if (poSum == int.Parse(z2[a][2].ToString()))
                                            {
                                                equalFGPO = true;
                                            }
                                        }
                                        else
                                        {
                                            string POsmGO = "SO_NO = '" + z3[c]["SO_NO"] + "' and SO_QTY ='" + z3[c]["SO_QTY"] + "'";
                                            DataRow[] POsmGORow2;
                                            POsmGORow2 = poTable.Select(POsmGO); // selects rows in PO datatable with the same SO number
                                            // deletes the rows in PO datatable
                                            foreach (DataRow POsmGORow2a in POsmGORow2)
                                            {
                                                POsmGORow2a.Delete();
                                            }
                                            poTable.AcceptChanges();
                                        }
                                    }
                                }
                                totalStock = totalStock + int.Parse(z2[a][2].ToString()); // computes the total stock/FG
                            }
                            else
                            {
                                totalStock = totalStock + int.Parse(z2[a][2].ToString());// computes the total stock/FG
                            }
                        }

                        EndBalance = totalStock - int.Parse(rw.Cells[7].Value.ToString()); // remaining stocks/fg after subtracting the pull qty

                        string newRemarks;

                        if (EndBalance < 0)
                        {
                            Quantity = totalStock; // changes the pull qty equal to FG if the total fg/stocks is less than the pull qty
                        }
                        else
                        {
                            Quantity = int.Parse(rw.Cells[7].Value.ToString());// pull qty is the same since fg/stocks is sufficient
                        }

                        if (totalStock == 0)
                        {
                            newRemarks = "NO FG"; // if the total stock is 0
                        }
                        else
                        {
                            if (rw.Cells[11].Value == DBNull.Value)
                            {
                                newRemarks = "-";
                            }
                            else
                            {
                                newRemarks = rw.Cells[11].Value.ToString();
                            }
                        }


                        // adds new row to the final pull datatable
                        DataRow pq1 = consolidatedTable.NewRow();
                        pq1["TRANS_DATE"] = date_time;
                        pq1["PROD_DATE"] = rw.Cells[0].Value;
                        pq1["PROD_TIME"] = rw.Cells[1].Value;
                        pq1["DEL_DATE"] = rw.Cells[2].Value;
                        pq1["DEL_TIME"] = rw.Cells[3].Value;
                        pq1["JOB_NO"] = rw.Cells[4].Value;
                        pq1["FACILITY"] = rw.Cells[5].Value;
                        pq1["PARTNUMBER"] = rw.Cells[6].Value;
                        pq1["PULL_QTY"] = rw.Cells[7].Value;
                        pq1["STOCK_QUANTITY"] = totalStock;
                        pq1["END_BALANCE"] = EndBalance;
                        pq1["QUANTITY_DELIVERED"] = Quantity;
                        pq1["GO_NUMBER"] = "";
                        pq1["SKU_ASSEMBLY"] = rw.Cells[9].Value;
                        pq1["GO_LINE_NUMBER"] = "";
                        pq1["CELL_NUM"] = rw.Cells[10].Value;
                        pq1["REMARKS"] = newRemarks;
                        pq1["PULL_TICKET_NUMBER"] = rw.Cells[12].Value;
                        pq1["LINE"] = rw.Cells[13].Value;
                        pq1["VENDOR_REMARKS"] = rw.Cells[15].Value;
                        pq1["PULLTYPE"] = rw.Cells[23].Value;
                        pq1["PULLNO"] = consoCounter;
                        pq1["QTY_DEL"] = rw.Cells[24].Value;
                        pq1["ORIGINAL_PULL"] = rw.Cells[25].Value;
                        consolidatedTable.Rows.Add(pq1);
                        consolidatedTable.AcceptChanges();
                        if (EndBalance <= 0)
                        {
                            string backexp = "[4] LIKE '" + rw.Cells[6].Value + " %' or [4] = '" + rw.Cells[6].Value + "'";
                            DataRow[] backrow;
                            backrow = inventoryTable.Select(backexp);
                            foreach (DataRow backrow2 in backrow)
                            {
                                backrow2.Delete();
                            }
                            inventoryTable.AcceptChanges();
                        }
                        else
                        {
                            expr2 = "[4] LIKE '%" + rw.Cells[6].Value + "%' or [4] = '" + rw.Cells[6].Value + "'";
                            z2 = inventoryTable.Select(expr2, "[1] DESC"); // selects rows in inventory datatable base on partnumber
                            int prevPull = 0;
                            int remainingPull = Convert.ToInt32(rw.Cells[7].Value.ToString()); // value of pull qty
                            for (int a = 0; a < z2.Length; a++) // loops through the selected rows
                            {
                                if (remainingPull > 0)
                                {
                                    prevPull = remainingPull; // saves the value of the remaining pull
                                    remainingPull = remainingPull - int.Parse(z2[a][2].ToString()); // subtracts the current FG/stocks
                                    if (remainingPull >= 0)
                                    {
                                        string POsmGO = "[1] = '" + z2[a][0] + "'";
                                        DataRow[] POsmGORow2;
                                        POsmGORow2 = inventoryTable.Select(POsmGO);

                                        foreach (DataRow POsmGORow2a in POsmGORow2)
                                        {
                                            POsmGORow2a.Delete();
                                        }
                                        inventoryTable.AcceptChanges();
                                    }
                                    else
                                    {
                                        z2[a][2] = int.Parse(z2[a][2].ToString()) - remainingPull;// updates the, subtracting the remaining pull to the current FG/stocks
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        private void GenerateStep2Final()
        {
            Form2 f2 = new Form2();
            conso_final_formatTable = new DataTable(); // creates a new instance of the datatable, this will store the final pull ticket data
            date_time = DateTime.Now.ToString("yyyy/MM/dd HH:mm:00");
            // adds columns to the datatable
            DataColumn pullQTcol = conso_final_formatTable.Columns.Add("TRANS_DATE", typeof(string));
            conso_final_formatTable.Columns.Add("PROD_DATE", typeof(string));
            conso_final_formatTable.Columns.Add("PROD_TIME", typeof(string));
            conso_final_formatTable.Columns.Add("DEL_DATE", typeof(string));
            conso_final_formatTable.Columns.Add("DEL_TIME", typeof(string));
            conso_final_formatTable.Columns.Add("JOB_NO", typeof(string));
            conso_final_formatTable.Columns.Add("FACILITY", typeof(string));
            conso_final_formatTable.Columns.Add("PARTNUMBER", typeof(string));
            conso_final_formatTable.Columns.Add("PULL_QTY", typeof(string));
            conso_final_formatTable.Columns.Add("OPEN_QTY", typeof(string));
            conso_final_formatTable.Columns.Add("STOCK_QTY", typeof(string));
            conso_final_formatTable.Columns.Add("QUANTITY_DELIVERED", typeof(string));
            conso_final_formatTable.Columns.Add("END_BALANCE", typeof(string));
            conso_final_formatTable.Columns.Add("GO_NUMBER", typeof(string));
            conso_final_formatTable.Columns.Add("SKU_ASSEMBLY", typeof(string));
            conso_final_formatTable.Columns.Add("GO_LINE_NUMBER", typeof(string));
            conso_final_formatTable.Columns.Add("CELL_NUM", typeof(string));
            conso_final_formatTable.Columns.Add("REMARKS", typeof(string));
            conso_final_formatTable.Columns.Add("PULL_TICKET_NUMBER", typeof(string));
            conso_final_formatTable.Columns.Add("LINE", typeof(string));
            conso_final_formatTable.Columns.Add("VENDOR_REMARKS", typeof(string));
            conso_final_formatTable.Columns.Add("PULLTYPE", typeof(string));
            conso_final_formatTable.Columns.Add("PULLNO", typeof(int));
            conso_final_formatTable.Columns.Add("QTY_DEL", typeof(int));
            conso_final_formatTable.Columns.Add("ORIGINAL_PULL", typeof(int));
            int pb = 0;
            DataRow[] z3;
            for (int w = 0; w < consolidatedTable.Rows.Count; w++) // loops through the datatable created in step1
            {
                f2.label2.Text = "Computing Inventory and PO . . .";
                f2.progressBar1.Maximum = consolidatedTable.Rows.Count;
                pb = pb + 1;
                f2.progressBar1.Value = w;
                f2.Refresh();
                int pull_qty = 0; // pull qty
                int open_qty = 0; // PO qty
                int quantity_del = 0; // quantity to be delivered
                int updated_openqty = 0; // updated PO qty
                int updated_stock = 0; // updated FG/stocks
                int endbalance = 0; // remaining pull qty
                int lessPOend = 0; // remaining PO
                poTable.DefaultView.Sort = "CUST_NO ASC"; // sorts the PO datatable
                poTable = poTable.DefaultView.ToTable();
                poTable.AcceptChanges();
                string expr2 = "NAME LIKE '%" + consolidatedTable.Rows[w]["PARTNUMBER"].ToString() + " %' or NAME = '" + consolidatedTable.Rows[w]["PARTNUMBER"].ToString() + "'";
                z3 = poTable.Select(expr2, "PRODUCT_NO DESC, CUST_NO ASC, CUST_NOLAST ASC");
                // selects rows in PO datatable base on partnumber to check for available PO then sort the results
                string Go_line = default, Go_line_number = default; // GO line and line number
                int end1 = 0;
                int ppp;
                //int stock_quantity = (int)consolidatedTable.Rows[w]["Stock_quantity"];
                int stock_quantity = Convert.ToInt32(consolidatedTable.Rows[w]["stock_quantity"]);// available stocks FG or PO
                //quantity_del = (int)consolidatedTable.Rows[w]["Quantity_delivered"];
                quantity_del = Convert.ToInt32(consolidatedTable.Rows[w]["Quantity_delivered"]);// the quantity to be delivered
                int myOpenQTY; // PO qty
                int newEndbal = 0;
                if (consolidatedTable.Rows[w][16] == DBNull.Value)
                {
                    if (z3.Length == 0)
                    {
                        myOpenQTY = 0; // no PO
                        Go_line = ""; // no line number
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["QUANTITY_DELIVERED"] = "0"; // 0 since there is no available PO
                        pq["END_BALANCE"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = "";
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = "NO PO"; // NO PO remarks if there is no available PO
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                        continue; // continue to the next row
                    }
                }
                else //if the remarks column is not null
                {
                    if (consolidatedTable.Rows[w][16].Equals("NO FG") && z3.Length > 0)
                    {
                        myOpenQTY = int.Parse(z3[0][4].ToString()); // PO qty
                        Go_line = z3[0][2].ToString(); // go line
                        Go_line_number = z3[0][3].ToString(); // go line number
                        endbalance = 0 - Convert.ToInt32(consolidatedTable.Rows[w]["PULL_QTY"]); // 0 because there is no FG
                        // adds new row to the datatable
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["QUANTITY_DELIVERED"] = "0"; // 0 because there is no FG
                        pq["END_BALANCE"] = endbalance;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = Go_line_number;
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = consolidatedTable.Rows[w]["REMARKS"];
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                        continue; // continue to the next row
                    }
                    if (consolidatedTable.Rows[w][16].Equals("NO FG") && z3.Length == 0) //checks if there is NO FG and there is no PO
                    {
                        myOpenQTY = 0; // no PO
                        Go_line = ""; // no GO line
                        endbalance = 0 - Convert.ToInt32(consolidatedTable.Rows[w]["PULL_QTY"]); // 0 because there is no FG
                        // adds new row to the datatable
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["QUANTITY_DELIVERED"] = "0"; // 0 because there is no PO and FG
                        pq["END_BALANCE"] = endbalance;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = "";
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = "NO FG, NO PO"; // NO FG, NO PO if both are not available
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                        continue; // continue to the next row
                    }
                    // for LACKING FG and PO ////////////////////////////
                    if (consolidatedTable.Rows[w][16].Equals("LACKING FG")) //checks if there is not enough FG
                    {
                        if (z3.Length > 0)
                        {
                            myOpenQTY = int.Parse(z3[0][4].ToString()); // PO qty
                            Go_line = z3[0][2].ToString(); // GO line
                            Go_line_number = z3[0][3].ToString(); // GO line number
                        }
                        else
                        {
                            myOpenQTY = 0; // 0 since there is no PO
                            Go_line = ""; // no GO line
                            Go_line_number = ""; // no GO line number
                        }
                        endbalance = Convert.ToInt32(consolidatedTable.Rows[w]["STOCK_QUANTITY"]) - Convert.ToInt32(consolidatedTable.Rows[w]["PULL_QTY"]); // remaining pull qty
                        // adds new row to the datatable
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["QUANTITY_DELIVERED"] = "0";
                        pq["END_BALANCE"] = endbalance;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = "";
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = "LACKING PO";
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                        continue; // continue to the next row65
                    }
                    if (consolidatedTable.Rows[w][16].Equals("LACKING PO"))
                    {
                        int sumPO = 0; // total PO qty
                        if (z3.Length > 0)
                        {
                            for (int sPO = 0; sPO <= z3.Length; sPO++)
                            {
                                sumPO = sumPO + int.Parse(z3[0][4].ToString());// computes the total PO
                            }
                        }
                        myOpenQTY = sumPO; // total PO qty
                        Go_line = ""; // no GO line
                        endbalance = Convert.ToInt32(consolidatedTable.Rows[w]["STOCK_QUANTITY"]) - Convert.ToInt32(consolidatedTable.Rows[w]["PULL_QTY"]); // remaining pull qty
                        // adds new row to the datatable
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["QUANTITY_DELIVERED"] = "0";
                        pq["END_BALANCE"] = endbalance;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = "";
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = "LACKING PO";
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                        continue; // continue to the next row
                    }
                    // //////////////////////////////////////////////////////////////////////////////
                    if (consolidatedTable.Rows[w][16] != ("NO FG") & z3.Length == 0)
                    {
                        myOpenQTY = 0; // 0 because there is no PO
                        Go_line = ""; // no GO line
                        // adds new row to the datatable
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["QUANTITY_DELIVERED"] = "0";
                        pq["END_BALANCE"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = "";
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = "NO PO";
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                        continue; // continue to the next row
                    }
                }
                // ////////////////////////////////////////////////////////////////////
                for (int i = 0; i < z3.Length; i++) // loops through the selected row in PO datatable
                {
                    quantity_del = quantity_del - open_qty;
                    try
                    {
                        open_qty = int.Parse(z3[i][4].ToString());
                    }
                    catch (Exception ex) { }
                    updated_stock = stock_quantity; // updated available stocks
                    myOpenQTY = int.Parse(z3[i][4].ToString()); // current PO qty
                    int v = 0; // used for checking where code in the loop was executed
                    int k = 0; // used for checking where code in the loop was executed
                    if (i == 0 & quantity_del <= open_qty)
                    {
                        Go_line = (string)z3[i][2]; // sets the GO line
                        try
                        {
                            Go_line_number = (string)z3[i][3]; // sets the GO line number
                        }
                        catch (Exception ex)
                        {
                            Go_line_number = "1";
                        }// give a value if there is an error
                        updated_openqty = stock_quantity - open_qty; // updated PO qty
                        updated_stock = updated_openqty; // updated available stocks
                        //endbalance = (int)consolidatedTable.Rows[w]["QUANTITY_DELIVERED"];
                        endbalance = Convert.ToInt32(consolidatedTable.Rows[w]["QUANTITY_DELIVERED"]);// remaining pull qty
                        updated_stock = stock_quantity;
                        int ab; // for storing remaining PO
                        //ab = open_qty - (int)consolidatedTable.Rows[w]["QUANTITY_DELIVERED"];
                        ab = open_qty - Convert.ToInt32(consolidatedTable.Rows[w]["QUANTITY_DELIVERED"]);
                        z3[i][4] = ab; // updates the PO qty
                        // /////////////////////////
                        if (consolidatedTable.Rows[w]["QUANTITY_DELIVERED"].Equals(0))
                        {
                            //open_qty = (int)z3[i][4];
                            open_qty = Convert.ToInt32(z3[i][4]);// sets the available PO
                            updated_stock = Convert.ToInt32(consolidatedTable.Rows[w]["STOCK_QUANTITY"]) - Convert.ToInt32(consolidatedTable.Rows[w]["PULL_QTY"]); // updates the remaining stocks after subtracting the pull qty
                        }
                        else if (Convert.ToInt32(consolidatedTable.Rows[w]["QUANTITY_DELIVERED"]) > 0 && (Convert.ToInt32(consolidatedTable.Rows[w]["END_BALANCE"]) < 0))
                        {
                            updated_stock = Convert.ToInt32(consolidatedTable.Rows[w]["STOCK_QUANTITY"]) - Convert.ToInt32(consolidatedTable.Rows[w]["PULL_QTY"]);// updates the remaining stocks after subtracting the pull qty
                        }
                        else
                        {
                            open_qty = Convert.ToInt32(consolidatedTable.Rows[w]["QUANTITY_DELIVERED"]);// sets the available PO to the value of qty to be delivered
                        }

                        // //////////////////////////////////////////////////
                        if (consolidatedTable.Rows[w]["PULL_QTY"] != DBNull.Value && int.TryParse(consolidatedTable.Rows[w]["PULL_QTY"].ToString(), out int pullQty))
                        {
                            lessPOend = stock_quantity - pullQty;
                        }
                        else
                        {
                            // Handle the case where the value is not a valid integer
                            lessPOend = stock_quantity; // or some other default value
                        }
                        // ///////////////////////////////////////////////////////////////
                        if (ab == 0)
                        {
                            string expr = "SO_NO = '" + Go_line + "' and SO_QTY ='" + z3[i][3] + "'";
                            DataRow[] prt;
                            prt = poTable.Select(expr);
                            foreach (DataRow prt2 in prt)
                            {
                                prt2.Delete();
                            }
                            poTable.AcceptChanges();
                        }
                        v = 1;
                    }
                    else if (i != 0 & quantity_del <= open_qty)
                    {
                        Go_line = z3[i][2].ToString(); // sets the GO line
                        try
                        {
                            Go_line_number = z3[i][3].ToString(); // sets the GO line number
                        }
                        catch (Exception ex)
                        {
                            Go_line_number = "1";
                        }// gives a value if there is there is no value
                        updated_openqty = stock_quantity - open_qty; // computes the available PO qty
                        int ab;
                        ab = open_qty - quantity_del;  // computes the remaining PO qty after subtracting the qty to be delivered
                        updated_stock = updated_openqty; // sets the updated stocks with the value of remaining PO
                        endbalance = updated_stock - open_qty; // computes the remaining pull qty
                        z3[i][4] = ab; // changes the value of the PO qty in the PO datatable
                        if (ab == 0)
                        {
                            string expr = "SO_NO = '" + Go_line + "' and SO_QTY ='" + z3[i][3] + "'";
                            DataRow[] prt;
                            prt = poTable.Select(expr);
                            foreach (DataRow prt2 in prt)
                            {
                                prt2.Delete();
                            }
                            poTable.AcceptChanges();
                        }
                        v = 1;
                        k = 1;
                    }
                    else if (i == 0 & quantity_del > open_qty)
                    {
                        Go_line = z3[i][2].ToString(); // sets the GO line
                        try
                        {
                            Go_line_number = z3[i][3].ToString(); // sets the GO line number
                        }
                        catch (Exception ex)
                        {
                            Go_line_number = "1";
                        }// gives a value if there is no GO line number

                        updated_openqty = stock_quantity - open_qty; // sets the updated PO qty
                        lessPOend = stock_quantity - open_qty; // sets the remaining PO
                        updated_stock = stock_quantity; // sets the remaining stock
                        endbalance = open_qty; // sets the remaining pull qty
                        end1 = updated_stock - open_qty; // sets the remaining stocks available
                        ppp = end1 - open_qty;
                        // selects and deletes the row since it is fully used
                        string expr = "SO_NO = '" + Go_line + "' and SO_QTY ='" + z3[i][3] + "'";
                        DataRow[] prt;
                        prt = poTable.Select(expr);
                        foreach (DataRow prt2 in prt)
                        {
                            prt2.Delete();
                        }
                        poTable.AcceptChanges();
                    }
                    else if (i != 0 & quantity_del > open_qty)
                    {
                        ppp = end1 - open_qty;
                        end1 = ppp;
                        Go_line = z3[i][2].ToString(); // sets the GO line
                        try
                        {
                            Go_line_number = z3[i][3].ToString(); // sets the GO line number
                        }
                        catch (Exception ex)
                        {
                            Go_line_number = "1";
                        }// gives a value if there is no GO line number
                        updated_openqty = stock_quantity - open_qty; // sets the remaining PO qty
                        updated_stock = updated_openqty; // sets the available stocks same as the value of the remaining PO
                        k = 2;
                        int ab;
                        ab = open_qty - Convert.ToInt32(consolidatedTable.Rows[w]["QUANTITY_DELIVERED"]); // sets the remaining PO qty
                        // selects the rows in PO datatable and deletes it since it is fully used
                        string expr = "SO_NO = '" + Go_line + "' and SO_QTY ='" + z3[i][3] + "'";
                        DataRow[] prt;
                        prt = poTable.Select(expr);
                        foreach (DataRow prt2 in prt)
                        {
                            prt2.Delete();
                        }
                        poTable.AcceptChanges();
                    }
                    if (i == 0 && z3.Length >= 1 && k == 0 && consolidatedTable.Rows[w]["QUANTITY_DELIVERED"].Equals(0))
                    {
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["QUANTITY_DELIVERED"] = endbalance;
                        pq["END_BALANCE"] = updated_stock;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = Go_line_number;
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = consolidatedTable.Rows[w]["REMARKS"];
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                    }
                    else if (i == 0 && z3.Length >= 1 && k == 0 && consolidatedTable.Rows[w]["QUANTITY_DELIVERED"].Equals(0) == false && consolidatedTable.Rows[w]["END_BALANCE"].Equals(0) == false && z3.Length == 1)
                    {
                        //newEndbal = stock_quantity - (int)consolidatedTable.Rows[w]["PULL_QTY"];
                        newEndbal = stock_quantity - Convert.ToInt32(consolidatedTable.Rows[w]["PULL_QTY"]);
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["QUANTITY_DELIVERED"] = endbalance;
                        pq["END_BALANCE"] = newEndbal;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = Go_line_number;
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = consolidatedTable.Rows[w]["REMARKS"];
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                    }
                    else if (i == 0 & z3.Length >= 1 & k == 0 & Convert.ToInt32(consolidatedTable.Rows[w]["QUANTITY_DELIVERED"]) > 0 & Convert.ToInt32(consolidatedTable.Rows[w]["END_BALANCE"]) < 0)
                    {
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = consolidatedTable.Rows[w]["STOCK_QUANTITY"];
                        pq["QUANTITY_DELIVERED"] = endbalance;
                        pq["END_BALANCE"] = updated_stock;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = Go_line_number;
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = consolidatedTable.Rows[w]["REMARKS"];
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                    }
                    else if (i == 0 & z3.Length == 1 & k == 0 && myOpenQTY < Convert.ToInt32(consolidatedTable.Rows[w]["PULL_QTY"]))
                    {
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = updated_stock;
                        pq["QUANTITY_DELIVERED"] = endbalance;
                        pq["END_BALANCE"] = updated_stock - open_qty;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = Go_line_number;
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = "PO LACKING";
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                    }
                    else if (i == 0 & z3.Length >= 1 & k == 0)
                    {
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["PROD_DATE"] = consolidatedTable.Rows[w]["PROD_DATE"];
                        pq["PROD_TIME"] = consolidatedTable.Rows[w]["PROD_TIME"];
                        pq["DEL_DATE"] = consolidatedTable.Rows[w]["DEL_DATE"];
                        pq["DEL_TIME"] = consolidatedTable.Rows[w]["DEL_TIME"];
                        pq["JOB_NO"] = consolidatedTable.Rows[w]["JOB_NO"];
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["PULL_QTY"] = consolidatedTable.Rows[w]["PULL_QTY"];
                        pq["OPEN_QTY"] = myOpenQTY;
                        pq["STOCK_QTY"] = updated_stock;
                        pq["QUANTITY_DELIVERED"] = endbalance;
                        pq["END_BALANCE"] = updated_stock - open_qty;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = Go_line_number;
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["REMARKS"] = consolidatedTable.Rows[w]["REMARKS"];
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["VENDOR_REMARKS"] = consolidatedTable.Rows[w]["VENDOR_REMARKS"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["QTY_DEL"] = consolidatedTable.Rows[w]["QTY_DEL"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                    }
                    else if (k == 1)
                    {
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["OPEN_QTY"] = open_qty;
                        pq["STOCK_QTY"] = end1;
                        pq["QUANTITY_DELIVERED"] = quantity_del;
                        pq["END_BALANCE"] = end1 - quantity_del;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = Go_line_number;
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                    }
                    else if (k == 2)
                    {
                        DataRow pq = conso_final_formatTable.NewRow();
                        pq["TRANS_DATE"] = date_time;
                        pq["FACILITY"] = consolidatedTable.Rows[w]["FACILITY"];
                        pq["PARTNUMBER"] = consolidatedTable.Rows[w]["PARTNUMBER"];
                        pq["OPEN_QTY"] = open_qty;
                        pq["STOCK_QTY"] = end1 + open_qty;
                        pq["QUANTITY_DELIVERED"] = open_qty;
                        pq["END_BALANCE"] = end1;
                        pq["GO_NUMBER"] = Go_line;
                        pq["SKU_ASSEMBLY"] = consolidatedTable.Rows[w]["SKU_ASSEMBLY"];
                        pq["GO_LINE_NUMBER"] = Go_line_number;
                        pq["CELL_NUM"] = consolidatedTable.Rows[w]["CELL_NUM"];
                        pq["PULL_TICKET_NUMBER"] = consolidatedTable.Rows[w]["PULL_TICKET_NUMBER"];
                        pq["LINE"] = consolidatedTable.Rows[w]["LINE"];
                        pq["PULLTYPE"] = consolidatedTable.Rows[w]["PULLTYPE"];
                        pq["PULLNO"] = consolidatedTable.Rows[w]["PULLNO"];
                        pq["ORIGINAL_PULL"] = consolidatedTable.Rows[w]["ORIGINAL_PULL"];
                        conso_final_formatTable.Rows.Add(pq);
                        conso_final_formatTable.AcceptChanges();
                    }

                    if (v == 1)
                    {
                        break;
                    }

                }
            }
        }
        //Method for restoring data from dmp
        private void btnrestore_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            string resFile;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = @"\\192.168.50.40\jmctest\";
                openFileDialog.Title = "Select the Backup File";
                // Add more settings or event handlers as needed
                // ...
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    resFile = openFileDialog.FileName;
                    try
                    {
                        // Drops the following tables before importing the data
                        //drop backlog table in the database
                        string backDel = "drop table BACKLOG";
                        OracleCommand backDelCom = new OracleCommand(backDel, con);
                        backDelCom.ExecuteNonQuery();
                    }
                    catch (Exception) { }
                    //Drop manualDr table in the database.
                    try
                    {
                        string backdel1 = "drop table MANUAL_DR";
                        OracleCommand backDelCom1 = new OracleCommand(backdel1, con);
                        backDelCom1.ExecuteNonQuery();
                    }
                    catch (Exception) { }
                    //Drop pullquantityrev table in the database
                    try
                    {
                        string backdel1 = "drop table PULL_QUANTITY_REV";
                        OracleCommand backDelCom1 = new OracleCommand(backdel1, con);
                        backDelCom1.ExecuteNonQuery();
                    }
                    catch (Exception) { }
                    //Drop pullticketrecord table in the database
                    try
                    {
                        string backdel1 = "drop table PULL_TICKET_RECORD";
                        OracleCommand backDelCom1 = new OracleCommand(backdel1, con);
                        backDelCom1.ExecuteNonQuery();
                    }
                    catch (Exception) { }
                    try
                    {
                        // Opens the command prompt and executes the following command to import the file
                        Process myprocess = new Process();
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.FileName = "cmd";
                        startInfo.RedirectStandardInput = true;
                        startInfo.RedirectStandardOutput = true;
                        startInfo.UseShellExecute = false;
                        startInfo.CreateNoWindow = true;
                        myprocess.StartInfo = startInfo;
                        myprocess.Start();
                        StreamReader SR = myprocess.StandardOutput;
                        StreamWriter SW = myprocess.StandardInput;
                        SW.WriteLine(@"cd C:\oraclexe\app\oracle\product\10.2.0\server\BIN");
                        SW.WriteLine($"imp mec/mec2024@192.168.50.40 buffer=4096 grants=Y file={resFile} tables=(BACKLOG, MANUAL_DR, PULL_QUANTITY_REV, PULL_TICKET_RECORD)");
                        SW.WriteLine("exit");
                        //exits command prompt window
                        //Checks if cmd is still running and shows a form with progress bar to notify the user that the process of restoring data is not yet finished
                        int procID = myprocess.Id;
                        f2.progressBar1.Style = ProgressBarStyle.Marquee;
                        f2.Show();
                        while (ProcessExists(procID).Equals(true))
                        {
                            f2.label1.Text = "RESTORING DATA...";
                            f2.Refresh();
                        }
                        f2.Close();
                        SW.Close();
                        SR.Close();
                        // Displays a form telling the user that restoration was successful
                        MessageBox.Show("Restored Succesfully");
                        f2.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    /* Helper method to check if a process is still running
                    bool ProcessExists(int processId)
                    {
                        return Process.GetProcessById(processId) != null;
                    }*/
                }
            }

        }
        private void deliverydatechange()
        {
            if (DelChanged == true)
            {
                pullQuanTempTable = new DataTable(); // creates a new instance of the pullticket datatable to store the new data
                pullQuanTempTable = cloneQuanTempTable.Clone(); // clones the columns of the datatable
                forFilter = new DataTable();
                forFilter = cloneQuanTempTable.Clone();
                datePick1 = deliverydatestart.Value;
                datePick2 = DateTime.Now;
                while (datePick1 <= datePick2)
                {
                    chkALL.Checked = true;
                    DataView negaView = new DataView(cloneQuanTempTable);
                    negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                    forFilter = new DataTable();
                    forFilter = cloneQuanTempTable.Clone();
                    forFilter = negaView.ToTable();
                    foreach (DataRow row in forFilter.Rows)
                    {
                        pullQuanTempTable.ImportRow(row);
                        pullQuanTempTable.AcceptChanges();
                    }
                    datePick1 = datePick1.AddDays(1);

                }
                for (int a = 0; a < selAdvBackTbl.Rows.Count; a++)
                {
                    string zexp = "PULL_TICKET_NUMBER = '" + selAdvBackTbl.Rows[a]["PULL_TICKET_NUMBER"] + "' AND LINE = '" + selAdvBackTbl.Rows[a]["LINE"] + "'";
                    DataRow[] z = backNewDelDate.Select(zexp);
                    foreach (DataRow z2 in z)
                    {
                        z2.Delete();
                        backNewDelDate.AcceptChanges();
                    }
                    z = pullQuanTempTable.Select(zexp);
                    foreach (DataRow z2 in z)
                    {
                        z2.Delete();
                        pullQuanTempTable.AcceptChanges();
                    }
                    int pullQTY = int.Parse(selAdvBackTbl.Rows[a]["ORIGINAL_PULL"].ToString()) - int.Parse(selAdvBackTbl.Rows[a]["QTY_DEL"].ToString());
                    DataRow pq = pullQuanTempTable.NewRow();
                    pq["PROD_DATE"] = selAdvBackTbl.Rows[a]["PROD_DATE"];
                    pq["PROD_TIME"] = selAdvBackTbl.Rows[a]["PROD_TIME"];
                    pq["DEL_DATE"] = selAdvBackTbl.Rows[a]["DEL_DATE"];
                    pq["DEL_TIME"] = selAdvBackTbl.Rows[a]["DEL_TIME"];
                    pq["JOB_NO"] = selAdvBackTbl.Rows[a]["JOB_NO"];
                    pq["FACILITY"] = selAdvBackTbl.Rows[a]["FACILITY"];
                    pq["PARTNUMBER"] = selAdvBackTbl.Rows[a]["PARTNUMBER"];
                    pq["PULL_QTY"] = pullQTY;
                    pq["VENDOR_NAME"] = "MEC ELECTRONICS PHILIPPINES CORP.";
                    pq["SKU_ASSEMBLY"] = selAdvBackTbl.Rows[a]["SKU_ASSEMBLY"];
                    pq["CELLNUMBER"] = selAdvBackTbl.Rows[a]["CELL_NUM"];
                    pq["REMARKS"] = "";
                    pq["PULL_TICKET_NUMBER"] = selAdvBackTbl.Rows[a]["PULL_TICKET_NUMBER"];
                    pq["LINE"] = selAdvBackTbl.Rows[a]["LINE"];
                    pq["FILEUPLOADDATE"] = DateTime.Now.ToString("M/d/yyy");
                    pq["VENDOR_REMARKS"] = selAdvBackTbl.Rows[a]["VENDOR_REMARKS"];
                    pq["ACKNOWLEDGMENT_DATE"] = "";
                    pq["ACKNOWLEDGMENT_REMARKS"] = "";
                    pq["COMMIT_QTY"] = "";
                    pq["COMMIT_DATE"] = "";
                    pq["BUYER_REMARKS_FOR_VENDOR"] = "";
                    pq["QTY_DELIVERED"] = "";
                    pq["DL_VARIENCE"] = "";
                    pq["HITMISS"] = "";
                    pq["STATUS"] = "";
                    pq["PULLTYPE"] = "BACKLOG";
                    pq["QTY_DEL"] = selAdvBackTbl.Rows[a]["QTY_DEL"];
                    pq["ORIGINAL_PULL"] = selAdvBackTbl.Rows[a]["ORIGINAL_PULL"];
                    pullQuanTempTable.Rows.Add(pq);
                    pullQuanTempTable.AcceptChanges();
                }
            }
            // saves the latest data of the pullticket to the datagridview
            pulltktgrid.DataSource = null;
            pullQuanTempTable.DefaultView.Sort = "DEL_DATE ASC, DEL_TIME ASC";
            pullQuanTempTable = pullQuanTempTable.DefaultView.ToTable();
            pullQuanTempTable.AcceptChanges();
            pulltktgrid.DataSource = pullQuanTempTable;
        }
        //import axmr430 button
        private void btnaxmr340_Click(object sender, EventArgs e)
        {
            if (btnaxmr432.BackColor == Color.Navy)
            {
                string strFileName;
                openFD.InitialDirectory = "'C:\'";
                openFD.Filter = "Excel Office | *xxlsx;*xls;";
                openFD.FilterIndex = 2;
                openFD.RestoreDirectory = true;
                if (openFD.ShowDialog().Equals(DialogResult.OK))
                {
                    strFileName = openFD.FileName;
                    if (!string.IsNullOrEmpty(FileSystem.Dir(openFD.FileName)))
                    {
                        if (!openFD.SafeFileName.ToString().ToUpper().Contains("AXMR430"))
                        {
                            Interaction.MsgBox("The file you ae trying to import is named " + openFD.SafeFileName + Microsoft.VisualBasic.Constants.vbCrLf + Microsoft.VisualBasic.Constants.vbCrLf + " Make sure you are importing the correct file!");
                            return;
                        }
                        else
                        {
                            if (strFileName != "")
                            {
                                getAxmrPO(strFileName);
                                btnaxmr340.BackColor = Color.Navy;
                                dataGridView4.Visible = true;
                                pulltktgrid.Visible = false;
                                slider.Visible = false;
                                refesherOrb.Visible = true;
                                btncmpl.Visible = true;
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Import AXMR4320 first!", "Warning!", MessageBoxButtons.OK);
            }
        }
        //method for importing AXMR430
        private void getAxmrPO(string filename)
        {
            DataTable getPOTbl = new DataTable(); //datatable( to hold data temporarily)
            Form2 f2 = new Form2(); //to access objects of form2
            f2.label2.Text = "Reading data from excel. . .";
            f2.Show(); //show form2 as dialog
            oExcel2 = new Excel.Application();
            oExcel2.ScreenUpdating = false;
            oExcel2.EnableEvents = false;
            xlWorkbook = oExcel2.Workbooks.Open(filename, XlFileAccess.xlReadOnly);
            xlWorksheet = xlWorkbook.Worksheets[1];
            int lRow = xlWorksheet.Range["A" + xlWorksheet.Rows.Count.ToString()].End[XlDirection.xlUp].Row;
            Range range = xlWorksheet.Range["A11:Z" + lRow.ToString()];
            object[,] data = (object[,])range.Value;
            for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                f2.progressBar1.Maximum = range.Columns.Count;
                f2.progressBar1.Value = cCnt;
                f2.Show(); //show form2 as dialog
                DataColumn Column = new DataColumn();
                Column.DataType = typeof(string);
                Column.ColumnName = cCnt.ToString();
                getPOTbl.Columns.Add(Column);
                for (int rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {
                    string rCnt1 = Convert.ToString(rCnt);
                    string cCnt1 = Convert.ToString(cCnt);
                    string cellVal = string.Empty;
                    cellVal = Convert.ToString(data[rCnt, cCnt]);
                    DataRow Row;
                    //Adds the row to the DataTable
                    if (cCnt == 1)
                    {
                        Row = getPOTbl.NewRow();
                        Row[Column.ColumnName.ToString()] = cellVal;
                        getPOTbl.Rows.Add(Row);
                    }
                    else
                    {
                        Row = getPOTbl.Rows[rCnt - 2];
                        Row[Column.ColumnName.ToString()] = cellVal;
                    }
                }
            }
            POTbl = getPOTbl.DefaultView.ToTable(false, "1", "5", "7", "10", "14", "16", "18");
            getPOTbl.AcceptChanges();
            MessageBox.Show("Import " + filename + " " +
                "successfully!!", "Success!!!!");
            pulltktgrid.DataSource = POTbl;
            f2.Close();
            xlWorkbook.Close();
            try
            {
                xlApp.Quit();
                oExcel2.Quit();
            }
            catch (Exception ex)
            {
                return;
            }
            ReleaseObject(lRow);
            ReleaseObject(xlWorksheet);
            ReleaseObject(xlWorkbook);
            ReleaseObject(oExcel2);
        }
        //filter button
        private void btnfilter_Click(object sender, EventArgs e)
        {
            if (backlogDel == true)
            {
                MessageBox.Show("Cannot filter items from BACKLOG! Go to records the BACKLOG and select only the ones you want to deliver.");
            }
            if (importCxmr.Equals(true))
            {
                MessageBox.Show("Inventory has already been computed! Please click the new button to start again or continue with the process!");
            }
            else
            {
                pnl_deladvise.Visible = true;
                btncancel1.Visible = true;
                orlabel.Visible = true;
                btncancel2.Visible = true;
                btnasper.Visible = true;
                btndeldate.Visible = true;
            }
        }
        private DateTime GetWeekStartDate(int weekNumber, int year)
        {
            DateTime startDate = new DateTime(year, 1, 1);
            DateTime weekDate = startDate.AddDays(7 * (weekNumber - 1));
            return weekDate.AddDays(-(int)weekDate.DayOfWeek);
        }
        //delivery date button in filter
        private void btndeldate_Click(object sender, EventArgs e)
        {

            DataTable distinctDT = new DataTable(); // creates a new datatable to filter FACILITY column
            distinctDT = cloneQuanTempTable.DefaultView.ToTable(true, "FACILITY"); // returns rows that have distinct values for column FACILITY

            dropDownFacilityFilter.Items.Clear();
            dropDownFacilityFilter.Items.Add("ALL");
            for (int sn = 0; sn < distinctDT.Rows.Count; sn++)
            {
                dropDownFacilityFilter.Items.Add(distinctDT.Rows[sn][0]);
            }
            // display dates from previous workweek and 2 days after current date
            DateTime weekStart = GetWeekStartDate(CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday) - 1, DateTime.Now.Year);
            DateTime dayCtr = weekStart.AddDays(-1);
            DateTime weekEnd = GetWeekStartDate(CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday), DateTime.Now.Year);
            DateTime dayCtr2 = weekEnd.AddDays(6);
            // Sets the minimum and maximum date of the datetimepicker
            if (Convert.ToDateTime(pullQuanTempTable.Rows[pullQuanTempTable.Rows.Count - 1][2].ToString()) < Convert.ToDateTime(DateTime.Now.ToString()))
            {
                deliverydatedocker1.MinDate = Convert.ToDateTime(pullQuanTempTable.Rows[pullQuanTempTable.Rows.Count - 1][2].ToString());
            }
            else
            {
                deliverydatedocker1.MinDate = DateTime.Now;
            }
            try
            {
                deliverydatedocker1.MaxDate = Convert.ToDateTime(cloneQuanTempTable.Rows[cloneQuanTempTable.Rows.Count - 1][2].ToString());
            }
            catch (Exception ex)
            {
                deliverydatedocker1.MaxDate = dayCtr2;
            }
            deliverydatestart.Value = Convert.ToDateTime(pullQuanTempTable.Rows[0][2]);
            deliverydatedocker1.Value = Convert.ToDateTime(pullQuanTempTable.Rows[pullQuanTempTable.Rows.Count - 1][2].ToString());
            date1sel = false;
            date2sel = false;
            facFil = false;
            pullticketfilter = false;
            datePick1 = new DateTime();
            datePick2 = new DateTime();
            forFilter = new DataTable();
            // for advance BACKLOG
            if (deliverydatedocker1.Value > DateTime.Now)
            {
                chkCAV1.Enabled = true;
                chkCAV2.Enabled = true;
                chkCAV3.Enabled = true;
                chkCAV5.Enabled = true;
                chkDANAM.Enabled = true;
                chkMACRO.Enabled = true;
                chkIPAI1.Enabled = true;
                chkIPAI2.Enabled = true;
                chkIPAI3.Enabled = true;
                chkALL.Enabled = true;
                if (selAdvBackTbl.Rows.Count > 0)
                {
                    DataTable samepartTbl = selAdvBackTbl.Clone(); // creates a clone of the datatable
                    samepartTbl = selAdvBackTbl.DefaultView.ToTable(true, "FACILITY");
                    //select the checkboxes that has the same label
                    for (int a = 0; a < samepartTbl.Rows.Count; a++)
                    {
                        if (samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == "CAV1")
                        {
                            chkCAV1.Checked = true;
                            chkCAV1.Enabled = false;
                        }
                        else if (samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == "CAV2")
                        {
                            chkCAV2.Checked = true;
                            chkCAV2.Enabled = false;
                        }
                        else if (samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == ("CAV3") || samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == ("CAV3/JAC"))
                        {
                            chkCAV3.Checked = true;
                            chkCAV3.Enabled = false;
                        }
                        else if (samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == "CAV5")
                        {
                            chkCAV5.Checked = true;
                            chkCAV5.Enabled = false;
                        }
                        else if (samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == "DANAM T" || samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == ("DANAMT"))
                        {
                            chkIPAI1.Checked = true;
                            chkIPAI1.Enabled = false;
                        }
                        else if (samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == "DKP")
                        {
                            chkIPAI2.Checked = true;
                            chkIPAI2.Enabled = false;
                        }
                        else if (samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == "CLP" || samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == ("CLP/CAV5"))
                        {
                            chkIPAI3.Checked = true;
                            chkIPAI3.Enabled = false;
                        }
                        else if (samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == "MACRO")
                        {
                            chkMACRO.Checked = true;
                            chkMACRO.Enabled = false;
                        }
                        else if (samepartTbl.Rows[a]["FACILITY"].ToString().ToUpper() == "DANAM")
                        {
                            chkDANAM.Checked = true;
                            chkDANAM.Enabled = false;
                        }
                    }
                    if (chkCAV1.Checked && chkCAV2.Checked && chkCAV3.Checked && chkCAV5.Checked && chkIPAI1.Checked && chkIPAI2.Checked && chkIPAI3.Checked && chkDANAM.Checked && chkMACRO.Checked)
                    {
                        chkALL.Checked = true;
                        chkALL.Enabled = false;
                    }
                    facFil = true; // facility values were filtered
                    date2sel = true; // datetimepicker "TO" was filtered
                }
                else
                {
                    selAdvBackTbl = new DataTable();
                    chkCAV1.Enabled = true;
                    chkCAV2.Enabled = true;
                    chkCAV3.Enabled = true;
                    chkCAV5.Enabled = true;
                    chkDANAM.Enabled = true;
                    chkMACRO.Enabled = true;
                    chkIPAI1.Enabled = true;
                    chkIPAI2.Enabled = true;
                    chkIPAI3.Enabled = true;
                    chkALL.Enabled = true;
                    //unchecks all the checkboxes
                    chkCAV1.Checked = false;
                    chkCAV2.Checked = false;
                    chkCAV3.Checked = false;
                    chkCAV5.Checked = false;
                    chkDANAM.Checked = false;
                    chkMACRO.Checked = false;
                    chkIPAI1.Checked = false;
                    chkIPAI2.Checked = false;
                    chkIPAI3.Checked = false;
                    chkALL.Checked = false;

                    facFil = false; //facility values weren't filtered
                    date2sel = false; //datetimepicker "TO" wasn't filtered
                }
            }
            else
            {
                dropDownFacilityFilter.Visible = false;
                chkCAV1.Enabled = false;
                chkCAV2.Enabled = false;
                chkCAV3.Enabled = false;
                chkCAV5.Enabled = false;
                chkDANAM.Enabled = false;
                chkMACRO.Enabled = false;
                chkIPAI1.Enabled = false;
                chkIPAI2.Enabled = false;
                chkIPAI3.Enabled = false;
                chkALL.Enabled = false;
                chkCAV1.Checked = false;
                chkCAV2.Checked = false;
                chkCAV3.Checked = false;
                chkCAV5.Checked = false;
                chkDANAM.Checked = false;
                chkMACRO.Checked = false;
                chkIPAI1.Checked = false;
                chkIPAI2.Checked = false;
                chkIPAI3.Checked = false;
                chkALL.Checked = false;
                facFil = false;
                date2sel = false;
                selAdvBackTbl = new DataTable();
                backNewDelDate = advBacklogTbl.Copy();
            }
            if (btndeldate.Visible == true)
            {
                btndeldate.Visible = false;
                btnasper.Visible = false;
                btncancel1.Visible = false;
                orlabel.Visible = false;
                btncancel2.Visible = false;
                pnldeliverydate.Visible = true;
            }
            else btndeldate.Visible = false;


            if (DelChanged == true)
            {
                deliverydatedocker1.Value = DateTime.Now;
            }
        }
        //set arrangement of data in AS PER ADVISE page
        public void asperadvse()
        {
            tempdatatbl.Columns["PROD_DATE"].SetOrdinal(1);
            tempdatatbl.Columns["PROD_TIME"].SetOrdinal(2);
            tempdatatbl.Columns["DEL_DATE"].SetOrdinal(3);
            tempdatatbl.Columns["DEL_TIME"].SetOrdinal(4);
            tempdatatbl.Columns["JOB_NO"].SetOrdinal(5);
            tempdatatbl.Columns["FACILITY"].SetOrdinal(6);
            tempdatatbl.Columns["PARTNUMBER"].SetOrdinal(7);
            tempdatatbl.Columns["PULL_QTY"].SetOrdinal(8);
            tempdatatbl.Columns["VENDOR_NAME"].SetOrdinal(9);
            tempdatatbl.Columns["SKU_ASSEMBLY"].SetOrdinal(10);
            tempdatatbl.Columns["CELLNUMBER"].SetOrdinal(11);
            tempdatatbl.Columns["REMARKS"].SetOrdinal(12);
            tempdatatbl.Columns["PULL_TICKET_NUMBER"].SetOrdinal(13);
            tempdatatbl.Columns["LINE"].SetOrdinal(14);
            tempdatatbl.Columns["FILEUPLOADDATE"].SetOrdinal(15);
            tempdatatbl.Columns["VENDOR_REMARKS"].SetOrdinal(16);
            tempdatatbl.Columns["ACKNOWLEDGMENT_DATE"].SetOrdinal(17);
            tempdatatbl.Columns["ACKNOWLEDGMENT_REMARKS"].SetOrdinal(18);
            tempdatatbl.Columns["BUYER_REMARKS_FOR_VENDOR"].SetOrdinal(19);
            tempdatatbl.Columns["QTY_DELIVERED"].SetOrdinal(20);
            tempdatatbl.Columns["DL_VARIENCE"].SetOrdinal(21);
            tempdatatbl.Columns["HITMISS"].SetOrdinal(22);
            tempdatatbl.Columns["STATUS"].SetOrdinal(23);
            tempdatatbl.Columns["PULLTYPE"].SetOrdinal(24);
            tempdatatbl.Columns["QTY_DEL"].SetOrdinal(25);
            tempdatatbl.Columns["ORIGINAL_PULL"].SetOrdinal(26);
        }
        //asperadvise button in filter
        private void btnasper_Click(object sender, EventArgs e)
        {

            DataGridView3.DataSource = pullQuanTempTable;
            DataGridViewCheckBoxColumn addColumn = new DataGridViewCheckBoxColumn();
            addColumn.HeaderText = "SELECTED";
            addColumn.Name = "cancelled";
            addColumn.Width = 80;
            DataGridView3.Columns.Insert(0, addColumn);
            DataGridView3.AllowUserToAddRows = false;
            //DataGridView3.DataSource = tempdatatbl;
            if (this.WindowState == FormWindowState.Maximized)
            {
                WindowState = FormWindowState.Maximized;
                AsPerPanel.BringToFront();
                AsPerPanel.Visible = true;
            }
            else
            {
                AsPerPanel.BringToFront();
                AsPerPanel.Visible = true;
            }
            DataGridView3.Columns["PROD_DATE"].Visible = false;
            DataGridView3.Columns["PROD_TIME"].Visible = false;
            DataGridView3.Columns["JOB_NO"].Visible = false;
            DataGridView3.Columns["VENDOR_NAME"].Visible = false;
            DataGridView3.Columns["SKU_ASSEMBLY"].Visible = false;
            DataGridView3.Columns["CELLNUMBER"].Visible = false;
            //DataGridView3.ClearSelection();
            foreach (DataGridViewRow row in DataGridView3.Rows)
            {
                row.Cells[0].Value = false;
                isSelectAllButtonAsperAdviceClicked = false;
            }
        }
        //ok button in deliverydate(filter)
        private void btnok_Click(object sender, EventArgs e)
        {
            DelChanged = false;
            /// checks if the starting date is ahead of the end date
            if (deliverydatestart.Value > deliverydatedocker1.Value)
            {
                MessageBox.Show("Invalid date!");
                return;
            }
            // checks if the dates or facility has been changed
            if (!(date1sel || date2sel || facFil))
            {
                return;
            }
            // checks if the user selects any facility
            if (deliverydatedocker1.Value > DateTime.Now)
            {
                if (!chkCAV1.Checked && !chkCAV2.Checked && !chkCAV3.Checked && !chkCAV5.Checked && !chkDANAM.Checked && !chkMACRO.Checked && !chkIPAI1.Checked && !chkIPAI2.Checked && !chkIPAI3.Checked && selAdvBackTbl.Rows.Count == 0)
                {
                    MessageBox.Show("Please select at least 1 facility to advance or select ALL!");
                    return;
                }
            }
            // checks if any facility was selected
            if (chkCAV1.Checked || chkCAV2.Checked || chkCAV3.Checked || chkCAV5.Checked || chkDANAM.Checked || chkMACRO.Checked || chkIPAI1.Checked || chkIPAI2.Checked || chkIPAI3.Checked || selAdvBackTbl.Rows.Count > 0)
            {
                facFil = true;
            }
            else
            {
                facFil = false;
            }
            pullQuanTempTable = new DataTable(); // creates a new instance of the pullticket datatable to store the new data
            pullQuanTempTable = cloneQuanTempTable.Clone(); // clones the columns of the datatable
            forFilter = new DataTable();
            forFilter = cloneQuanTempTable.Clone();
            date1sel = false; // changed to false because the date can't be changed
            if (date1sel && date2sel && facFil)
            {
                datePick1 = deliverydatestart.Value;
                datePick2 = deliverydatedocker1.Value;
                while (datePick1 <= datePick2)
                {
                    if (dropDownFacilityFilter.SelectedValue.Equals("ALL"))
                    {
                        DataView negaView = new DataView(cloneQuanTempTable);
                        negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                        forFilter = new DataTable();
                        forFilter = cloneQuanTempTable.Clone();
                        forFilter = negaView.ToTable();
                        foreach (DataRow row in forFilter.Rows)
                        {
                            pullQuanTempTable.ImportRow(row);
                            pullQuanTempTable.AcceptChanges();
                        }
                        datePick1 = datePick1.AddDays(1);
                    }
                    else
                    {
                        DataView negaView = new DataView(cloneQuanTempTable);
                        negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY = '" + dropDownFacilityFilter.SelectedValue + "'";
                        forFilter = new DataTable();
                        forFilter = cloneQuanTempTable.Clone();
                        forFilter = negaView.ToTable();
                        foreach (DataRow row in forFilter.Rows)
                        {
                            pullQuanTempTable.ImportRow(row);
                            pullQuanTempTable.AcceptChanges();
                        }
                        datePick1 = datePick1.AddDays(1);
                    }
                }
            }
            else if (date1sel == true && date2sel == true)
            {
                datePick1 = deliverydatestart.Value;
                datePick2 = deliverydatedocker1.Value;
                while (datePick1 <= datePick2)
                {
                    DataView negaView = new DataView(cloneQuanTempTable);
                    negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                    forFilter = new DataTable();
                    forFilter = cloneQuanTempTable.Clone();
                    forFilter = negaView.ToTable();
                    foreach (DataRow row in forFilter.Rows)
                    {
                        pullQuanTempTable.ImportRow(row);
                        pullQuanTempTable.AcceptChanges();
                    }
                    datePick1 = datePick1.AddDays(1);
                }
            }
            else if (date1sel == true && facFil == true)
            {
                datePick1 = deliverydatestart.Value;
                datePick2 = deliverydatedocker1.Value;
                while (datePick1 <= datePick2)
                {
                    if (dropDownFacilityFilter.SelectedValue.Equals("ALL"))
                    {
                        DataView negaView = new DataView(cloneQuanTempTable);
                        negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                        forFilter = new DataTable();
                        forFilter = cloneQuanTempTable.Clone();
                        forFilter = negaView.ToTable();
                        foreach (DataRow row in forFilter.Rows)
                        {
                            pullQuanTempTable.ImportRow(row);
                            pullQuanTempTable.AcceptChanges();
                        }
                        datePick1 = datePick1.AddDays(1);
                    }
                    else
                    {
                        DataView negaView = new DataView(cloneQuanTempTable);
                        negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY = '" + dropDownFacilityFilter.SelectedValue + "'";
                        forFilter = new DataTable();
                        forFilter = cloneQuanTempTable.Clone();
                        forFilter = negaView.ToTable();
                        foreach (DataRow row in forFilter.Rows)
                        {
                            pullQuanTempTable.ImportRow(row);
                            pullQuanTempTable.AcceptChanges();
                        }
                        datePick1 = datePick1.AddDays(1);
                    }
                }
            }
            else if (date1sel == true)
            {
                datePick1 = deliverydatestart.Value;
                datePick2 = deliverydatedocker1.Value;

                while (datePick1 <= datePick2)
                {
                    DataView negaView = new DataView(cloneQuanTempTable);
                    negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                    forFilter = new DataTable();
                    forFilter = cloneQuanTempTable.Clone();
                    forFilter = negaView.ToTable();
                    foreach (DataRow row in forFilter.Rows)
                    {
                        pullQuanTempTable.ImportRow(row);
                        pullQuanTempTable.AcceptChanges();
                    }
                    datePick1 = datePick1.AddDays(1);
                }
            }
            else if (facFil == true && date2sel == false && date1sel == false) //if only facilities are filtered and not the date
            {
                datePick1 = deliverydatestart.Value;
                datePick2 = deliverydatedocker1.Value;

                while (datePick1 <= datePick2)
                {
                    if (dropDownFacilityFilter.SelectedValue == "ALL")
                    {
                        DataView negaView = new DataView(cloneQuanTempTable);
                        negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                        forFilter = new DataTable();
                        forFilter = cloneQuanTempTable.Clone();
                        forFilter = negaView.ToTable();

                        foreach (DataRow row in forFilter.Rows)
                        {
                            pullQuanTempTable.ImportRow(row);
                            pullQuanTempTable.AcceptChanges();
                        }
                        datePick1 = datePick1.AddDays(1);
                    }
                    else
                    {
                        DataView negaView = new DataView(cloneQuanTempTable);
                        negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY = '" + dropDownFacilityFilter.SelectedValue + "'";
                        forFilter = new DataTable();
                        forFilter = cloneQuanTempTable.Clone();
                        forFilter = negaView.ToTable();

                        foreach (DataRow row in forFilter.Rows)
                        {
                            pullQuanTempTable.ImportRow(row);
                            pullQuanTempTable.AcceptChanges();
                        }

                        datePick1 = datePick1.AddDays(1);
                    }
                }
            }
            else if (date2sel == true && facFil == true && date1sel == false) //if facilities and end date are filtered
            {
                datePick1 = deliverydatestart.Value;
                datePick2 = deliverydatedocker1.Value;
                while (datePick1 <= datePick2)
                {
                    if (datePick1 > DateTime.Now)
                    {
                        if (chkALL.Checked)
                        {
                            DataView negaView = new DataView(cloneQuanTempTable);
                            negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                            forFilter = new DataTable();
                            forFilter = cloneQuanTempTable.Clone();
                            forFilter = negaView.ToTable();
                            foreach (DataRow row in forFilter.Rows)
                            {
                                pullQuanTempTable.ImportRow(row);
                                pullQuanTempTable.AcceptChanges();
                            }
                        }
                        else if (chkCAV1.Checked && chkCAV2.Checked && chkCAV3.Checked && chkCAV5.Checked && chkIPAI2.Checked && chkIPAI1.Checked && chkIPAI3.Checked && chkMACRO.Checked && chkDANAM.Checked)
                        {
                            DataView negaView = new DataView(cloneQuanTempTable);
                            negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                            forFilter = new DataTable();
                            forFilter = cloneQuanTempTable.Clone();
                            forFilter = negaView.ToTable();
                            foreach (DataRow row in forFilter.Rows)
                            {
                                pullQuanTempTable.ImportRow(row);
                                pullQuanTempTable.AcceptChanges();
                            }
                        }
                        else
                        {
                            if (chkCAV1.Checked)
                            {
                                DataView negaView = new DataView(cloneQuanTempTable);
                                negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY = 'CAV1'";
                                forFilter = new DataTable();
                                forFilter = cloneQuanTempTable.Clone();
                                forFilter = negaView.ToTable();

                                foreach (DataRow row in forFilter.Rows)
                                {
                                    pullQuanTempTable.ImportRow(row);
                                    pullQuanTempTable.AcceptChanges();
                                }
                            }
                            if (chkCAV2.Checked)
                            {
                                DataView negaView = new DataView(cloneQuanTempTable);
                                negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY = 'CAV2'";
                                forFilter = new DataTable();
                                forFilter = cloneQuanTempTable.Clone();
                                forFilter = negaView.ToTable();

                                foreach (DataRow row in forFilter.Rows)
                                {
                                    pullQuanTempTable.ImportRow(row);
                                    pullQuanTempTable.AcceptChanges();
                                }
                            }
                            if (chkCAV3.Checked)
                            {
                                DataView negaView = new DataView(cloneQuanTempTable);
                                negaView.RowFilter = $"DEL_DATE = '{datePick1}' and FACILITY in ('CAV3', 'CAV3/JAC')";
                                forFilter = new DataTable();
                                forFilter = cloneQuanTempTable.Clone();
                                forFilter = negaView.ToTable();
                                foreach (DataRow row in forFilter.Rows)
                                {
                                    pullQuanTempTable.ImportRow(row);
                                    pullQuanTempTable.AcceptChanges();
                                }
                            }
                            if (chkCAV5.Checked)
                            {
                                DataView negaView = new DataView(cloneQuanTempTable);
                                negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY = 'CAV5'";
                                forFilter = new DataTable();
                                forFilter = cloneQuanTempTable.Clone();
                                forFilter = negaView.ToTable();
                                foreach (DataRow row in forFilter.Rows)
                                {
                                    pullQuanTempTable.ImportRow(row);
                                    pullQuanTempTable.AcceptChanges();
                                }
                            }
                            if (chkIPAI1.Checked)
                            {
                                DataView negaView = new DataView(cloneQuanTempTable);
                                negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY in ('DANAM T', 'DANAMT')";
                                forFilter = new DataTable();
                                forFilter = cloneQuanTempTable.Clone();
                                forFilter = negaView.ToTable();
                                foreach (DataRow row in forFilter.Rows)
                                {
                                    pullQuanTempTable.ImportRow(row);
                                    pullQuanTempTable.AcceptChanges();
                                }
                            }
                            if (chkIPAI2.Checked)
                            {
                                DataView negaView = new DataView(cloneQuanTempTable);
                                negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY = 'DKP'";
                                forFilter = new DataTable();
                                forFilter = cloneQuanTempTable.Clone();
                                forFilter = negaView.ToTable();
                                foreach (DataRow row in forFilter.Rows)
                                {
                                    pullQuanTempTable.ImportRow(row);
                                    pullQuanTempTable.AcceptChanges();
                                }
                            }
                            if (chkIPAI3.Checked)
                            {
                                DataView negaView = new DataView(cloneQuanTempTable);
                                negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY in ('CLP', 'CLP/CAV5')";
                                forFilter = new DataTable();
                                forFilter = cloneQuanTempTable.Clone();
                                forFilter = negaView.ToTable();
                                foreach (DataRow row in forFilter.Rows)
                                {
                                    pullQuanTempTable.ImportRow(row);
                                    pullQuanTempTable.AcceptChanges();
                                }
                            }
                            if (chkDANAM.Checked)
                            {
                                DataView negaView = new DataView(cloneQuanTempTable);
                                negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY = 'DANAM'";
                                forFilter = new DataTable();
                                forFilter = cloneQuanTempTable.Clone();
                                forFilter = negaView.ToTable();
                                foreach (DataRow row in forFilter.Rows)
                                {
                                    pullQuanTempTable.ImportRow(row);
                                    pullQuanTempTable.AcceptChanges();
                                }
                            }
                            if (chkMACRO.Checked)
                            {
                                DataView negaView = new DataView(cloneQuanTempTable);
                                negaView.RowFilter = "DEL_DATE = '" + datePick1 + "' and FACILITY = 'MACRO'";
                                forFilter = new DataTable();
                                forFilter = cloneQuanTempTable.Clone();
                                forFilter = negaView.ToTable();
                                foreach (DataRow row in forFilter.Rows)
                                {
                                    pullQuanTempTable.ImportRow(row);
                                    pullQuanTempTable.AcceptChanges();
                                }
                            }
                        }
                    }
                    else
                    {
                        DataView negaView = new DataView(cloneQuanTempTable);
                        negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                        forFilter = new DataTable();
                        forFilter = cloneQuanTempTable.Clone();
                        forFilter = negaView.ToTable();
                        foreach (DataRow row in forFilter.Rows)
                        {
                            pullQuanTempTable.ImportRow(row);
                            pullQuanTempTable.AcceptChanges();
                        }
                    }
                    datePick1 = datePick1.AddDays(1);
                }
                //deleting of data from pullticket datatable and BACKLOG datatable
                for (int a = 0; a < selAdvBackTbl.Rows.Count; a++)
                {
                    string zexp = "PULL_TICKET_NUMBER = '" + selAdvBackTbl.Rows[a]["PULL_TICKET_NUMBER"] + "' AND LINE = '" + selAdvBackTbl.Rows[a]["LINE"] + "'";
                    DataRow[] z = backNewDelDate.Select(zexp);
                    foreach (DataRow z2 in z)
                    {
                        z2.Delete();
                        backNewDelDate.AcceptChanges();
                    }
                    z = pullQuanTempTable.Select(zexp);
                    foreach (DataRow z2 in z)
                    {
                        z2.Delete();
                        pullQuanTempTable.AcceptChanges();
                    }
                    int pullQTY = 0;
                    pullQTY = int.Parse(selAdvBackTbl.Rows[a]["ORIGINAL_PULL"].ToString()) - int.Parse(selAdvBackTbl.Rows[a]["QTY_DEL"].ToString());
                    DataRow pq = pullQuanTempTable.NewRow();
                    pq["PROD_DATE"] = selAdvBackTbl.Rows[a]["PROD_DATE"];
                    pq["PROD_TIME"] = selAdvBackTbl.Rows[a]["PROD_TIME"];
                    pq["DEL_DATE"] = selAdvBackTbl.Rows[a]["DEL_DATE"];
                    pq["DEL_TIME"] = selAdvBackTbl.Rows[a]["DEL_TIME"];
                    pq["JOB_NO"] = selAdvBackTbl.Rows[a]["JOB_NO"];
                    pq["FACILITY"] = selAdvBackTbl.Rows[a]["FACILITY"];
                    pq["PARTNUMBER"] = selAdvBackTbl.Rows[a]["PARTNUMBER"];
                    pq["PULL_QTY"] = pullQTY;
                    pq["VENDOR_NAME"] = "MEC ELECTRONICS PHILIPPINES CORP.";
                    pq["SKU_ASSEMBLY"] = selAdvBackTbl.Rows[a]["SKU_ASSEMBLY"];
                    pq["CELLNUMBER"] = selAdvBackTbl.Rows[a]["CELL_NUM"];
                    pq["REMARKS"] = "";
                    pq["PULL_TICKET_NUMBER"] = selAdvBackTbl.Rows[a]["PULL_TICKET_NUMBER"];
                    pq["LINE"] = selAdvBackTbl.Rows[a]["LINE"];
                    pq["FILEUPLOADDATE"] = DateTime.Now.ToString("M/d/yyyy");
                    pq["VENDOR_REMARKS"] = selAdvBackTbl.Rows[a]["VENDOR_REMARKS"];
                    pq["ACKNOWLEDGMENT_DATE"] = "";
                    pq["ACKNOWLEDGMENT_REMARKS"] = "";
                    pq["COMMIT_QTY"] = "";
                    pq["COMMIT_DATE"] = "";
                    pq["BUYER_REMARKS_FOR_VENDOR"] = "";
                    pq["QTY_DELIVERED"] = "";
                    pq["DL_VARIENCE"] = "";
                    pq["HITMISS"] = "";
                    pq["STATUS"] = "";
                    pq["PULLTYPE"] = "BACKLOG";
                    pq["QTY_DEL"] = selAdvBackTbl.Rows[a]["QTY_DEL"];
                    pq["ORIGINAL_PULL"] = selAdvBackTbl.Rows[a]["ORIGINAL_PULL"];
                    pullQuanTempTable.Rows.Add(pq);
                    pullQuanTempTable.AcceptChanges();
                }
            }
            else if (date2sel == true && facFil == false && date1sel == false)//if only end date or "TO" date was filtered
            {
                datePick1 = deliverydatestart.Value;
                datePick2 = deliverydatedocker1.Value;
                while (datePick1 <= datePick2)
                {
                    // filters the datatable based on the current date value of the loop
                    DataView negaView = new DataView(cloneQuanTempTable);
                    negaView.RowFilter = "DEL_DATE = '" + datePick1 + "'";
                    forFilter = new DataTable();
                    forFilter = cloneQuanTempTable.Clone();
                    forFilter = negaView.ToTable();
                    foreach (DataRow row in forFilter.Rows)
                    {
                        pullQuanTempTable.ImportRow(row);
                        pullQuanTempTable.AcceptChanges();
                    }
                    datePick1 = datePick1.AddDays(1);
                }
                backNewDelDate = advBacklogTbl.Copy();
                selAdvBackTbl = new DataTable();
            }
            // saves the latest data of the pullticket to the datagridview
            pulltktgrid.DataSource = null;
            pullQuanTempTable.DefaultView.Sort = "DEL_DATE ASC, DEL_TIME ASC";
            pullQuanTempTable = pullQuanTempTable.DefaultView.ToTable();
            pullQuanTempTable.AcceptChanges();
            pulltktgrid.DataSource = pullQuanTempTable;
            // DataGridView2.DataSource = pullQuanTempTable;
            // pullticketfilter = true;
            //Panel17.Visible = false;
            //Panel17.Width = 0;
            facFil = false;
            date1sel = false;
            date2sel = false;

            // shows a popup form that tells the user the pullticket has been filtered
            //SalesAdd salesAddnew = new SalesAdd();
            //salesAddnew.Label1.Text = "FILTERED";
            // PanelAnimator2.ShowSync(salesAddnew);
            MessageBox.Show("FILTERED");
            
            pnldeliverydate.Visible = false;
            pnl_deladvise.Visible = false;
        }
        //cancel button in pnl_deladvise(filter)
        private void btncancel1_Click(object sender, EventArgs e)
        {
            pnl_deladvise.Visible = false;
        }
        private void refesherOrb_MouseEnter(object sender, EventArgs e)
        {
            refesherOrb.BackColor = Color.Transparent;
        }
        //clear imported data
        private void refesherOrb_Click(object sender, EventArgs e)
        {
            DialogResult DivineReaper = MessageBox.Show("Are you sure you want to refresh?", "Are you Sure?", MessageBoxButtons.YesNo);
            switch (DivineReaper)
            {
                case DialogResult.Yes:
                    if (importPull.Equals(true) && pullQuanTempTable != null)
                    {
                        if (excelData3 != null && importCxmr.Equals(true))
                        {
                            excelData3.Clear();
                            if (importInv.Equals(true) && invTblCopy != null)
                            {
                                invTblCopy.Clear();
                                if (importPO.Equals(true) && unpostedTbl != null)
                                {
                                    unpostedTbl.Clear();
                                    btncmpl.Visible = false;
                                    if (poTable.Equals(null))
                                    {
                                        poTable.Clear();
                                    }
                                }
                            }
                        }
                        pullQuanTempTable.Clear();
                        backNewDelDate.Clear();
                        pullrecordTable.Clear();
                        pullQuantityTable.Clear();
                        cloneQuanTempTable.Clear();
                        pullDeleteTable.Clear();
                        pullRevTable.Clear();
                        manualDrTable.Clear();
                        advBacklogTbl.Clear();
                        amarokTbl.Clear();
                        kanbanTbl.Clear();
                        historyBack.Clear();
                        backTempotbl.Clear();
                        forFilter.Clear();
                        excelData2.Clear();
                        pullbackTable.Clear();
                        DataGridView3.DataSource = null;
                        btnfilter.Visible = false;
                        btnpull_tkt.Enabled = true;
                        btncxmr620.Enabled = true;
                        btnaimr407.Enabled = true;
                        btnaxmr432.Enabled = true;
                        btnpull_tkt.BackColor = Color.SteelBlue;
                        btncxmr620.BackColor = Color.SteelBlue;
                        btnaimr407.BackColor = Color.SteelBlue;
                        btnaxmr432.BackColor = Color.SteelBlue;
                        btnaxmr340.BackColor = Color.SteelBlue;
                        pnl_import2.Location = new System.Drawing.Point(35, 60);
                        dataGridView4.Visible = false;
                        pulltktgrid.Visible = false;
                        slider.Visible = true;
                        refesherOrb.Visible = false;
                        btndownload.Visible = false;
                        importPull = false;
                        importCxmr = false;
                        importInv = false;
                        importPO = false;
                        calculateBtn = false;
                        if (genClick == true || asperclick == true)
                        {
                            pulltktgrid.Columns.Remove("cancelled");
                        }
                        pulltktgrid.DataSource = null;
                        DataGridView3.DataSource = null;
                        dataGridView4.DataSource = null;
                        genClick = false;
                        asperclick = false;
                        backlogDel = false;
                        asperMarked = false;
                        pullticketfilter = false;
                        asperTable = new DataTable();
                        pullQuanTempTable = new DataTable();
                        pullQuantityTable = new DataTable();
                        inventoryTable = new DataTable();
                        pullbackTable = new DataTable();
                        poTable = new DataTable();
                        consolidatedTable = new DataTable();
                        conso_final_formatTable = new DataTable();
                        backTempotbl = new DataTable();
                        cloneQuanTempTable = new DataTable();
                        manualDrTable = new DataTable();
                        asperQuanTable = new DataTable();
                        backNewDelDate = new DataTable();
                        pullRevTable = new DataTable();
                        pullDeleteTable = new DataTable();
                        fgSM = new DataTable();
                        poTableCopy = new DataTable();
                        totalPullSM = new DataTable();
                        colorPOsm = new DataTable();
                        consoBinding = new BindingSource();
                        historyBack = new DataTable();
                        invTblCopy = new DataTable();
                        invSMorNonTbl = new DataTable();
                        selAdvBackTbl = new DataTable();
                        advBacklogTbl = new DataTable();
                        amarokTbl = new DataTable();
                        kanbanTbl = new DataTable();
                        if (pnldeliverydate.Visible == true || pnl_deladvise.Visible == true)
                        {
                            pnldeliverydate.Visible = false;
                            pnl_deladvise.Visible = false;
                        }
                        //xlWorkbook.Close();
                        xlApp.Quit();
                        //oExcel2.Quit();
                        if (xlWorkbook != null)
                        {
                            ReleaseObject(xlWorkbook);
                            xlWorkbook = null;
                        }
                        ReleaseObject(xlWorksheet);
                        ReleaseObject(oExcel2);
                        break;
                    }
                    else
                    {
                        MessageBox.Show("No Data. . . .", "Error!");
                    }
                    break;
                case DialogResult.No:
                    break;
            }
        }
        //download button
        private void download_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            //Checks if the user has imported all the necessary files

            if (importPull == false || importCxmr == false || importInv == false || importPO == false || calculateBtn == false)
            {
                MessageBox.Show("Please complete all the process before printing!");
                return;
            }
            //Checks if there are data on the datagridview to print
            if (DataGridView3.Rows.Count == 0)
            {
                MessageBox.Show("There is no data to print!");
                return;
            }
            //Displays a prompt that lets the user to select the location where the files will be saved
            date_time_file = DateTime.Now.ToString("yyyy_MM_dd (hhmmtt)");
            SaveFileDialog saveFD = new SaveFileDialog();
            saveFD.Filter = "Excel Workbook|*.xls|Excel Workbook 2011|*.xlsx";
            saveFD.Title = "Save Excel File";
            saveFD.FileName = $"Delivery {date_time_file}.xls";
            saveFD.InitialDirectory = @"C:\";
            string strFileName;
            bool blnFileOpen;
            if (saveFD.ShowDialog() == DialogResult.OK)
            {
                if (saveFD.FileName != "")
                {
                    try
                    {
                        FileStream fs = (FileStream)saveFD.OpenFile();
                        fs.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("File Not Found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                strFileName = saveFD.FileName;
                blnFileOpen = false;
                try
                {
                    FileStream fileTemp = File.OpenWrite(strFileName);
                    fileTemp.Close();
                }
                catch (Exception ex)
                {
                    blnFileOpen = false;
                    return;
                }
                if (File.Exists(strFileName))
                {
                    File.Delete(strFileName);
                }
            }
            else
            {
                return;
            }
            f2.Show(); //shows a form that display the status of the program
                       //Opens the command prompt and executes the commands for backing up the database
            try
            {
                string fulldate = DateTime.Now.ToString("MMM-dd-yyyy-hh-mm-ss");
                fulldate += "-BEFORE-PRINTING";
                Process myprocess = new Process();
                ProcessStartInfo StartInfo = new ProcessStartInfo();
                StartInfo.FileName = "cmd"; // starts cmd window
                StartInfo.RedirectStandardInput = true;
                StartInfo.RedirectStandardOutput = true;
                StartInfo.UseShellExecute = false; // required to redirect
                StartInfo.CreateNoWindow = true;
                myprocess.StartInfo = StartInfo;
                myprocess.Start();
                StreamReader SR = myprocess.StandardOutput;
                StreamWriter SW = myprocess.StandardInput;
                // The command to back up the database
                SW.WriteLine(@"cd C:\oraclee\app\oracle\product\10.2.0\server\BIN");
                SW.WriteLine(@"exp mec/mec2024@192.168.50.40 buffer=4096 grants=Y file=\\192.168.50.40\jmctest\" + fulldate + ".dmp tables=(BACKLOG, MANUAL_DR, PULL_QUANTITY_REV, PULL_TICKET_RECORD)");
                SW.WriteLine("exit"); // exits command prompt window
                                      // txtResults.Text = SR.ReadToEnd; // returns results of the command window
                int procID = myprocess.Id;

                // Checks if the process is still ongoing and displays a text telling the user that the back up process is still running
                while (ProcessExists(procID))
                {
                    f2.label2.Text = "Backing up data . . .";
                    f2.Refresh();
                }
                SW.Close();
                SR.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ". Please check your connection then try again.");
                return;
            }
            ////////////////////////////////////////////////////////////////////////////////
            conso_delTable = new DataTable(); // creates a new instance of the datatable, this table will store the data for printing
            backlogTempoTable = new DataTable(); // creates a new instance of the datatable, this table will store the temporary BACKLOG data
                                                 // adds columns to the datatable
            DataColumn pullQTcol = conso_delTable.Columns.Add("TRANS_DATE", typeof(string));
            conso_delTable.Columns.Add("PROD_DATE", typeof(string));
            conso_delTable.Columns.Add("PROD_TIME", typeof(string));
            conso_delTable.Columns.Add("DEL_DATE", typeof(string));
            conso_delTable.Columns.Add("DEL_TIME", typeof(string));
            conso_delTable.Columns.Add("JOB_NO", typeof(string));
            conso_delTable.Columns.Add("FACILITY", typeof(string));
            conso_delTable.Columns.Add("PARTNUMBER", typeof(string));
            conso_delTable.Columns.Add("PULL_QTY", typeof(string));
            conso_delTable.Columns.Add("OPEN_QTY", typeof(string));
            conso_delTable.Columns.Add("STOCK_QTY", typeof(string));
            conso_delTable.Columns.Add("QUANTITY_FOR_DELIVERY", typeof(string));
            conso_delTable.Columns.Add("END_BALANCE", typeof(string));
            conso_delTable.Columns.Add("GO_NUMBER", typeof(string));
            conso_delTable.Columns.Add("SKU_ASSEMBLY", typeof(string));
            conso_delTable.Columns.Add("GO_LINE_NUMBER", typeof(string));
            conso_delTable.Columns.Add("CELL_NUM", typeof(string));
            conso_delTable.Columns.Add("REMARKS", typeof(string));
            conso_delTable.Columns.Add("PULL_TICKET_NUMBER", typeof(string));
            conso_delTable.Columns.Add("LINE", typeof(string));
            conso_delTable.Columns.Add("VENDOR_REMARKS", typeof(string));
            conso_delTable.Columns.Add("DELNO", typeof(Int32));
            conso_delTable.Columns.Add("QTY_DEL", typeof(int));
            conso_delTable.Columns.Add("ORIGINAL_PULL", typeof(int));
            conso_delTable.Columns.Add("UL_LABEL", typeof(string));
            // adds columns to the datatable
            DataColumn backlogCol = backlogTempoTable.Columns.Add("TRANS_DATE", typeof(string));
            backlogTempoTable.Columns.Add("PROD_DATE", typeof(string));
            backlogTempoTable.Columns.Add("PROD_TIME", typeof(string));
            backlogTempoTable.Columns.Add("DEL_DATE", typeof(string));
            backlogTempoTable.Columns.Add("DEL_TIME", typeof(string));
            backlogTempoTable.Columns.Add("JOB_NO", typeof(string));
            backlogTempoTable.Columns.Add("FACILITY", typeof(string));
            backlogTempoTable.Columns.Add("PARTNUMBER", typeof(string));
            backlogTempoTable.Columns.Add("PULL_QTY", typeof(string));
            backlogTempoTable.Columns.Add("OPEN_QTY", typeof(string));
            backlogTempoTable.Columns.Add("STOCK_QTY", typeof(string));
            backlogTempoTable.Columns.Add("QUANTITY_DELIVERED", typeof(string));
            backlogTempoTable.Columns.Add("END_BALANCE", typeof(string));
            backlogTempoTable.Columns.Add("GO_NUMBER", typeof(string));
            backlogTempoTable.Columns.Add("SKU_ASSEMBLY", typeof(string));
            backlogTempoTable.Columns.Add("GO_LINE_NUMBER", typeof(string));
            backlogTempoTable.Columns.Add("CELL_NUM", typeof(string));
            backlogTempoTable.Columns.Add("REMARKS", typeof(string));
            backlogTempoTable.Columns.Add("PULL_TICKET_NUMBER", typeof(string));
            backlogTempoTable.Columns.Add("LINE", typeof(string));
            backlogTempoTable.Columns.Add("VENDOR_REMARKS", typeof(string));
            backlogTempoTable.Columns.Add("QTY_DEL", typeof(int));
            backlogTempoTable.Columns.Add("ORIGINAL_PULL", typeof(int));
            DataTable consoRecord = new DataTable(); // creates a new datatable
            f2.Refresh();
            int pb = 0;
            // checks if the pull is not from BACKLOG
            if (backlogDel == false)
            {
                // deleting of pullticket data in the database base from the new pull ticket from excel file
                for (int pd = 0; pd < pullDeleteTable.Rows.Count; pd++)
                {
                    // deletes data in the pull ticket table in the database
                    string delPullTick = $"delete from pull_ticket_record where PULL_TICKET_NUMBER ='{pullDeleteTable.Rows[pd]["PULL_TICKET_NUMBER"]}' AND LINE = '{pullDeleteTable.Rows[pd]["LINE"]}'";
                    OracleCommand delPullTickAd = new OracleCommand(delPullTick, con);
                    delPullTickAd.ExecuteNonQuery();
                    // deletes data in the pull ticket revision table in the database
                    string delPullTick2 = $"delete from PULL_QUANTITY_REV where PULL_TICKET_NUMBER ='{pullDeleteTable.Rows[pd]["PULL_TICKET_NUMBER"]}' AND LINE = '{pullDeleteTable.Rows[pd]["LINE"]}'";
                    OracleCommand delPullTickAd2 = new OracleCommand(delPullTick2, con);
                    delPullTickAd2.ExecuteNonQuery();
                    // deletes data in the backlog table in the database
                    string delPullTick3 = $"delete from backlog where PULL_TICKET_NUMBER ='{pullDeleteTable.Rows[pd]["PULL_TICKET_NUMBER"]}' AND LINE = '{pullDeleteTable.Rows[pd]["LINE"]}'";
                    OracleCommand delPullTickAd3 = new OracleCommand(delPullTick3, con);
                    delPullTickAd3.ExecuteNonQuery();
                }
                //loops through the datatable that contains data of pull ticket with revisions
                for (int pullrev = 0; pullrev < pullRevTable.Rows.Count; pullrev++)
                {
                    // updates the record of pull ticket in the database
                    string upPullTick = $"UPDATE pull_ticket_record SET DEL_DATE = '{pullRevTable.Rows[pullrev]["DEL_DATE"]}', PULL_QTY ='{pullRevTable.Rows[pullrev]["NEW_PULL"]}', REMARKS ='{pullRevTable.Rows[pullrev]["REMARKS"]}' WHERE PULL_TICKET_NUMBER ='{pullRevTable.Rows[pullrev]["PULL_TICKET_NUMBER"]}' AND LINE = '{pullRevTable.Rows[pullrev]["LINE"]}'";
                    OracleCommand upPullTickAd = new OracleCommand(upPullTick, con);
                    upPullTickAd.ExecuteNonQuery();
                    // inserts in the pull ticket revision table in the database
                    string insPullTick = $"insert into PULL_QUANTITY_REV values('{ pullRevTable.Rows[pullrev]["DEL_DATE"] }','{pullRevTable.Rows[pullrev]["FACILITY"]}','{pullRevTable.Rows[pullrev]["PARTNUMBER"]}',{pullRevTable.Rows[pullrev]["PREVIOUS_PULL"]},{pullRevTable.Rows[pullrev]["NEW_PULL"]},'{pullRevTable.Rows[pullrev]["REMARKS"]}','{pullRevTable.Rows[pullrev]["PULL_TICKET_NUMBER"]}','{pullRevTable.Rows[pullrev]["LINE"]}','{pullRevTable.Rows[pullrev]["DATE_REVISED"]}')";
                    OracleCommand insPullTickAd = new OracleCommand(insPullTick, con);
                    insPullTickAd.ExecuteNonQuery();
                }
            }
            ///////////////////////////////////// recording of pull ticket data ///////////////////////////////////////
            //selecting new pull ticket in the pull ticket datatable
            DataView newView = new DataView(pullQuanTempTable);
            newView.RowFilter = "PULLTYPE = 'NEW'";
            pullQuanTempTable = newView.ToTable();
            pullQuanTempTable.AcceptChanges();
            for (int rw = 0; rw < pullQuanTempTable.Rows.Count; rw++) // loops through the pull ticket datatable
            {
                f2.label2.Text = "Recording of Pull Ticket . . .";
                f2.progressBar1.Maximum = pullQuanTempTable.Rows.Count;
                pb++;
                f2.progressBar1.Value = pb;
                f2.Refresh();
                //insert data in the pull ticket table in database
                string insRec = $"insert into pull_ticket_record (Prod_date, Prod_time, Del_date, Del_time, Job_no,Facility, Partnumber,Pull_qty, Vendor_Name, SKU_Assembly, CellNumber, Remarks, Pull_ticket_number, Line,FileUploadDate,Vendor_remarks,Acknowledgment_Remarks,Buyer_Remarks_for_Vendor,QTY_DELIVERED,DL_Varience,HitMiss,Status) values ('{pullQuanTempTable.Rows[rw][0] + "','" + pullQuanTempTable.Rows[rw][1]}','{pullQuanTempTable.Rows[rw][2]}','{pullQuanTempTable.Rows[rw][3]}','{pullQuanTempTable.Rows[rw][4]}','{pullQuanTempTable.Rows[rw][5]}','{pullQuanTempTable.Rows[rw][6]}',{pullQuanTempTable.Rows[rw][7]},'{pullQuanTempTable.Rows[rw][8]}','{pullQuanTempTable.Rows[rw][9]}','{pullQuanTempTable.Rows[rw][10]}','{pullQuanTempTable.Rows[rw][11]}','{pullQuanTempTable.Rows[rw][12]}','{pullQuanTempTable.Rows[rw][13]}','{pullQuanTempTable.Rows[rw][14]}','{pullQuanTempTable.Rows[rw][15]}','{pullQuanTempTable.Rows[rw][17]}','{pullQuanTempTable.Rows[rw][18]}','{pullQuanTempTable.Rows[rw][19]}','{pullQuanTempTable.Rows[rw][20]}','{pullQuanTempTable.Rows[rw][21]}','{pullQuanTempTable.Rows[rw][22]}')";
                OracleCommand insRecAd = new OracleCommand(insRec, con);
                insRecAd.ExecuteNonQuery();
            }
            if (backlogDel == false) // checks if the pull is not from BACKLOG
            {
                // recording of new pull ticket that is not going to be delivered because of "AS PER" status
                for (int rw = 0; rw < asperQuanTable.Rows.Count; rw++) // loops through the datatable containing "AS PER" pull
                {
                    // checks if there are already records saved in pull ticket table in database
                    string backlogsel = $"SELECT * from PULL_TICKET_RECORD where PULL_TICKET_NUMBER = '{asperQuanTable.Rows[rw]["PULL_TICKET_NUMBER"]}' and LINE = '{asperQuanTable.Rows[rw]["LINE"]}'";
                    OracleDataAdapter backlogselAd = new OracleDataAdapter(backlogsel, con);
                    DataTable backChange = new DataTable();
                    backlogselAd.Fill(backChange);
                    if (backChange.Rows.Count > 0)
                    {
                        continue; // continue to the next row if there is already a data saved
                    }
                    f2.label2.Text = "Recording of Pull Ticket . . .";
                    f2.progressBar1.Maximum = asperQuanTable.Rows.Count;
                    f2.progressBar1.Value = rw + 1;
                    f2.Refresh();
                    // Insert the new data in pull ticket table in database
                    string insRec = $"insert into pull_ticket_record (Prod_date, Prod_time, Del_date, Del_time, Job_no,Facility, Partnumber,Pull_qty, Vendor_Name, SKU_Assembly, CellNumber, Remarks, Pull_ticket_number, Line,FileUploadDate,Vendor_remarks,Acknowledgment_Remarks, Buyer_Remarks_for_Vendor,QTY_DELIVERED,DL_Varience,HitMiss,Status) values ('{asperQuanTable.Rows[rw][0]}', '{asperQuanTable.Rows[rw][1]}', '{asperQuanTable.Rows[rw][2]}', '{asperQuanTable.Rows[rw][3]}', '{asperQuanTable.Rows[rw][4]}', '{asperQuanTable.Rows[rw][5]}', '{asperQuanTable.Rows[rw][6]}', {asperQuanTable.Rows[rw][7]}, '{asperQuanTable.Rows[rw][8]}', '{asperQuanTable.Rows[rw][9]}', '{asperQuanTable.Rows[rw][10]}', '{asperQuanTable.Rows[rw][11]}', '{asperQuanTable.Rows[rw][12]}', '{asperQuanTable.Rows[rw][13]}', '{asperQuanTable.Rows[rw][14]}', '{asperQuanTable.Rows[rw][15]}', '{asperQuanTable.Rows[rw][17]}', '{asperQuanTable.Rows[rw][18]}', '{asperQuanTable.Rows[rw][19]}', '{asperQuanTable.Rows[rw][20]}', '{asperQuanTable.Rows[rw][21]}', '{asperQuanTable.Rows[rw][22]}')";
                    OracleCommand insRecAd = new OracleCommand(insRec, con);
                    insRecAd.ExecuteNonQuery();
                }
            }
            if (backlogDel == false)
            {
                for (int rw = 0; rw < manualDrTable.Rows.Count; rw++)
                {
                    // deletes the data in the database
                    string insRec = $"delete from MANUAL_DR where PULL_TICKET_NUMBER = '{manualDrTable.Rows[rw]["PULL_TICKET_NUMBER"]}' and LINE = '{manualDrTable.Rows[rw]["LINE"]}'";
                    OracleCommand insRecAd = new OracleCommand(insRec, con);
                    insRecAd.ExecuteNonQuery();
                    //inserts data in the pull ticket table in database
                    string insRec2 = $"insert into pull_ticket_record (Prod_date, Prod_time, Del_date, Del_time, Job_no,Facility, Partnumber,Pull_qty, Vendor_Name, SKU_Assembly, CellNumber, Remarks, Pull_ticket_number, Line,FileUploadDate,Vendor_remarks,Acknowledgment_Remarks, Buyer_Remarks_for_Vendor,QTY_DELIVERED,DL_Varience,HitMiss,Status) values ('{manualDrTable.Rows[rw][0]}','{manualDrTable.Rows[rw][1]}','{manualDrTable.Rows[rw][2]}','{manualDrTable.Rows[rw][3]}','{manualDrTable.Rows[rw][4]}','{manualDrTable.Rows[rw][5]}','{manualDrTable.Rows[rw][6]}',{manualDrTable.Rows[rw][7]},'{manualDrTable.Rows[rw][8]}','{manualDrTable.Rows[rw][9]}','{manualDrTable.Rows[rw][10]}','{manualDrTable.Rows[rw][11]}','{manualDrTable.Rows[rw][12]}','{manualDrTable.Rows[rw][13]}','{manualDrTable.Rows[rw][14]}','{manualDrTable.Rows[rw][15]}','{manualDrTable.Rows[rw][16]}','{manualDrTable.Rows[rw][17]}','{manualDrTable.Rows[rw][18]}','{manualDrTable.Rows[rw][19]}','{manualDrTable.Rows[rw][20]}','{manualDrTable.Rows[rw][21]}')";
                    OracleCommand insRecAd2 = new OracleCommand(insRec2, con);
                    insRecAd2.ExecuteNonQuery();
                }
            }
            if (backlogDel == false)
            {
                // deletes the saved data in BACKLOG in the database
                string backDel = $"delete from BACKLOG";
                OracleCommand backDelCom = new OracleCommand(backDel, con);
                backDelCom.ExecuteNonQuery();
            }
            else
            {
                // deletes the saved data in BACKLOG in the database
                for (int ab = 0; ab < pulltktgrid.Rows.Count; ab++)
                {
                    string backDel = $"delete from BACKLOG where PULL_TICKET_NUMBER = '{pulltktgrid.Rows[ab].Cells["PULL_TICKET_NUMBER"].Value}' and LINE = '{pulltktgrid.Rows[ab].Cells["LINE"].Value}'";
                    OracleCommand backDelCom = new OracleCommand(backDel, con);
                    backDelCom.ExecuteNonQuery();
                }
            }
            //////////////////////////////// recording of BACKLOG////////////////////////////////
            if (backlogDel == false)
            {
                string hisBack; //history of backlogs
                int balanceBack; //remaining balance
                for (int cr = 0; cr < asperTable.Rows.Count; cr++)
                {
                    //for BACKLOG history //////////////////////
                    string ckPull = "PULL_TICKET_NUMBER = '" + asperTable.Rows[cr]["PULL_TICKET_NUMBER"] + "' AND LINE ='" + asperTable.Rows[cr]["LINE"] + "'";
                    DataRow[] p04 = historyBack.Select(ckPull);
                    if (p04.Length == 1)
                    {
                        if (p04[0]["HISTORY"] is DBNull)
                        {
                            hisBack = "";
                        }
                        else
                        {
                            hisBack = p04[0]["HISTORY"].ToString();
                        }
                    }
                    else
                    {
                        hisBack = "";
                    }
                    balanceBack = int.Parse(asperTable.Rows[cr]["ORIGINAL_PULL"].ToString()) - int.Parse(asperTable.Rows[cr]["QTY_DEL"].ToString()); //gets the remaining balance
                    string dropqry4 = $"insert into BACKLOG (Trans_date,Prod_date,Prod_time,Del_date,Del_time,Job_no,Facility,Partnumber,Pull_qty,Open_qty,Stock_qty,Quantity_delivered,End_Balance,GO_Number,SKU_ASSEMBLY,GO_Line_Number,CELL_NUM,REMARKS,PULL_TICKET_NUMBER,LINE,VENDOR_REMARKS, QTY_DEL, ORIGINAL_PULL, BACKLOGTYPE, HISTORY, BALANCE) values ('{asperTable.Rows[cr]["TRANS_DATE"]} ','{asperTable.Rows[cr]["PROD_DATE"]}','{asperTable.Rows[cr]["PROD_TIME"]}','{asperTable.Rows[cr]["DEL_DATE"].ToString().Replace(" 12:00:00 AM", "")}','{asperTable.Rows[cr]["DEL_TIME"]}','{asperTable.Rows[cr]["JOB_NO"]}','{asperTable.Rows[cr]["FACILITY"]}','{asperTable.Rows[cr]["PARTNUMBER"]}','{asperTable.Rows[cr]["PULL_QTY"]}','{asperTable.Rows[cr]["OPEN_QTY"]}','{asperTable.Rows[cr]["STOCK_QTY"]}','{asperTable.Rows[cr]["QUANTITY_DELIVERED"]}','{asperTable.Rows[cr]["END_BALANCE"]}','{asperTable.Rows[cr]["GO_Number"]}','{asperTable.Rows[cr]["SKU_ASSEMBLY"]}','{asperTable.Rows[cr]["GO_LINE_NUMBER"]}','{asperTable.Rows[cr]["CELL_NUM"]}','{asperTable.Rows[cr]["REMARKS"]}','{asperTable.Rows[cr]["PULL_TICKET_NUMBER"]}','{asperTable.Rows[cr]["LINE"]}','{asperTable.Rows[cr]["VENDOR_REMARKS"]}',{int.Parse(asperTable.Rows[cr]["QTY_DEL"].ToString())},{asperTable.Rows[cr]["ORIGINAL_PULL"]},'{asperTable.Rows[cr]["BACKLOGTYPE"]}','{hisBack}','{balanceBack}')";
                    OracleCommand oracommand4 = new OracleCommand(dropqry4, con);
                    oracommand4.ExecuteNonQuery();
                }
                ///////////////for updating adjusted BACKLOG //////////////////////////////////
                for (int cr = 0; cr < backNewDelDate.Rows.Count; cr++)
                {
                    string ckPull = $"PULL_TICKET_NUMBER = '{backNewDelDate.Rows[cr]["PULL_TICKET_NUMBER"]}' AND LINE ='{backNewDelDate.Rows[cr]["LINE"]}'";
                    DataRow[] p04 = historyBack.Select(ckPull);
                    if (p04.Length == 1)
                    {
                        if (p04[0]["HISTORY"] is DBNull)
                        {
                            hisBack = "";
                        }
                        else
                        {
                            hisBack = p04[0]["HISTORY"].ToString();
                        }
                    }
                    else
                    {
                        hisBack = "";
                    }
                    balanceBack = Convert.ToInt32(backNewDelDate.Rows[cr]["ORIGINAL_PULL"]) - Convert.ToInt32(backNewDelDate.Rows[cr]["QTY_DEL"]); //gets the remaining balance
                                                                                                                                                   //inserts data in BACKLOG table in the database
                    string dropqry4 = $"insert into BACKLOG (Trans_date,Prod_date,Prod_time,Del_date,Del_time,Job_no,Facility,Partnumber,Pull_qty,Open_qty,Stock_qty,Quantity_delivered,End_Balance,GO_Number,SKU_ASSEMBLY,GO_Line_Number,CELL_NUM,REMARKS,PULL_TICKET_NUMBER,LINE,VENDOR_REMARKS, QTY_DEL, ORIGINAL_PULL, BACKLOGTYPE, HISTORY, BALANCE) values ('{backNewDelDate.Rows[cr]["TRANS_DATE"]}','{backNewDelDate.Rows[cr]["PROD_DATE"]}','{backNewDelDate.Rows[cr]["PROD_TIME"]}','{backNewDelDate.Rows[cr]["DEL_DATE"].ToString().Replace(" 12:00:00 AM", "")}','{backNewDelDate.Rows[cr]["DEL_TIME"]}','{backNewDelDate.Rows[cr]["JOB_NO"]}','{backNewDelDate.Rows[cr]["FACILITY"]}','{backNewDelDate.Rows[cr]["PARTNUMBER"]}','{backNewDelDate.Rows[cr]["PULL_QTY"]}','{backNewDelDate.Rows[cr]["OPEN_QTY"]}','{backNewDelDate.Rows[cr]["STOCK_QTY"]}','{backNewDelDate.Rows[cr]["QUANTITY_DELIVERED"]}','{backNewDelDate.Rows[cr]["END_BALANCE"]}','{backNewDelDate.Rows[cr]["GO_Number"]}','{backNewDelDate.Rows[cr]["SKU_ASSEMBLY"]}','{backNewDelDate.Rows[cr]["GO_LINE_NUMBER"]}','{backNewDelDate.Rows[cr]["CELL_NUM"]}','{backNewDelDate.Rows[cr]["REMARKS"]}','{backNewDelDate.Rows[cr]["PULL_TICKET_NUMBER"]}','{backNewDelDate.Rows[cr]["LINE"]}','{backNewDelDate.Rows[cr]["VENDOR_REMARKS"]}',{int.Parse(backNewDelDate.Rows[cr]["QTY_DEL"].ToString())},{backNewDelDate.Rows[cr]["ORIGINAL_PULL"]},'{backNewDelDate.Rows[cr]["BACKLOGTYPE"]}','{hisBack}','{balanceBack}')";
                    OracleCommand oracommand4 = new OracleCommand(dropqry4, con);
                    oracommand4.ExecuteNonQuery();
                }
            }
            pb = 0;
            string PROD_DATE, PROD_TIME, DEL_DATE, DEL_TIME, JOB_NO, REMARKS, VENDOR_REMARKS;
            int QTY_DEL, ORIGINAL_PULL, totalOpen;
            DataTable samePull = new DataTable(); // creates a new datatable
            samePull.Columns.Add("PULLNO"); // adds new columns
            samePull.Columns.Add("COUNT"); // adds new columns
                                           // selects unique pullno from the pull ticket data
            var query = conso_final_formatTable.AsEnumerable().Select(d => new { PULLNO = d["PULLNO"] }).Select(dr => dr.PULLNO).Distinct();
            foreach (int colName in query)
            {
                int cName = int.Parse(colName.ToString());
                int cCount = conso_final_formatTable.Rows.Cast<DataRow>().Count(row => row["PULLNO"].ToString() == cName.ToString());
                samePull.Rows.Add(colName, cCount);
            }
            DataRow[] uniquePullno;
            string uniqueExpr = "COUNT < 2";
            uniquePullno = samePull.Select(uniqueExpr); // selects rows with unique pullno
            for (int a = 0; a < uniquePullno.Length; a++) // loops through the selected rows
            {
                f2.label2.Text = "Saving BACKLOG . . .";
                f2.progressBar1.Maximum = uniquePullno.Length;
                f2.progressBar1.Value = a + 1;
                f2.Refresh();
                int cr = consoBinding.Find("PULLNO", uniquePullno[a][0]); // find the index of the row
                string hisBack; // for BACKLOG history
                int balanceBack; // for remaining balance
                if (Convert.ToInt32(DataGridView3.Rows[cr].Cells["Quantity_delivered"].Value) == 0 || Convert.ToInt32(DataGridView3.Rows[cr].Cells["End_Balance"].Value) < 0 || Convert.ToBoolean(DataGridView3.Rows[cr].Cells["cancelled"].Value) == true || DataGridView3.Rows[cr].Cells["Remarks"].Value.ToString().Contains("PO LACKING"))
                {
                    int endBal = Convert.ToInt32(DataGridView3.Rows[cr].Cells["End_Balance"].Value);
                    int quanDel = Convert.ToInt32(DataGridView3.Rows[cr].Cells["Quantity_delivered"].Value);
                    if (Convert.ToBoolean(DataGridView3.Rows[cr].Cells["cancelled"].Value) == true)// checks if the datagridcheckboxcolumn has been checked/selected
                    {
                        quanDel = 0; // changed to 0 because it is cancelled
                        endBal = 0; // changed to 0 because it is cancelled
                        if (DataGridView3.Rows[cr].Cells["Pull_qty"].Value == DBNull.Value) { }
                        else
                        {
                            string ckPull = $"PULL_TICKET_NUMBER = '{DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value}' AND LINE ='{DataGridView3.Rows[cr].Cells["LINE"].Value}'";
                            DataRow[] p04 = historyBack.Select(ckPull);
                            if (p04.Length == 1)
                            {
                                if (p04[0]["HISTORY"] == DBNull.Value)
                                {
                                    hisBack = "";
                                }
                                else
                                {
                                    hisBack = p04[0]["HISTORY"].ToString();
                                }
                            }
                            else
                            {
                                hisBack = "";
                            }
                            balanceBack = Convert.ToInt32(DataGridView3.Rows[cr].Cells["ORIGINAL_PULL"].Value) - Convert.ToInt32(DataGridView3.Rows[cr].Cells["QTY_DEL"].Value);
                            //insert the data into BACKLOG datatable in database
                            string dropqry4 = $"insert into BACKLOG (Trans_date,Prod_date,Prod_time,Del_date,Del_time,Job_no,Facility,Partnumber,Pull_qty,Open_qty,Stock_qty,Quantity_delivered,End_Balance,GO_Number,SKU_ASSEMBLY,GO_Line_Number,CELL_NUM,REMARKS,PULL_TICKET_NUMBER,LINE,VENDOR_REMARKS, QTY_DEL, ORIGINAL_PULL,HISTORY,BALANCE) values ('{DataGridView3.Rows[cr].Cells["Trans_date"].Value}', '{DataGridView3.Rows[cr].Cells["Prod_date"].Value}', '{DataGridView3.Rows[cr].Cells["Prod_time"].Value}', '{DataGridView3.Rows[cr].Cells["Del_date"].Value.ToString().Replace(" 12:00:00 AM", "")}', '{DataGridView3.Rows[cr].Cells["Del_time"].Value}', '{DataGridView3.Rows[cr].Cells["Job_no"].Value}', '{DataGridView3.Rows[cr].Cells["Facility"].Value}', '{DataGridView3.Rows[cr].Cells["Partnumber"].Value}', '{DataGridView3.Rows[cr].Cells["Pull_qty"].Value}', '{DataGridView3.Rows[cr].Cells["Open_qty"].Value}', '{DataGridView3.Rows[cr].Cells["Stock_qty"].Value}', '{quanDel}', '{endBal}', '{DataGridView3.Rows[cr].Cells["GO_Number"].Value}', '{DataGridView3.Rows[cr].Cells["SKU_ASSEMBLY"].Value}', '{DataGridView3.Rows[cr].Cells["GO_Line_Number"].Value}', '{DataGridView3.Rows[cr].Cells["CELL_NUM"].Value}', '{DataGridView3.Rows[cr].Cells["REMARKS"].Value}', '{DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value}', '{DataGridView3.Rows[cr].Cells["LINE"].Value}', '{DataGridView3.Rows[cr].Cells["VENDOR_REMARKS"].Value}', '{int.Parse(DataGridView3.Rows[cr].Cells["QTY_DEL"].Value.ToString()) + quanDel}', '{DataGridView3.Rows[cr].Cells["ORIGINAL_PULL"].Value}', '{hisBack}', '{balanceBack}')";
                            OracleCommand oracommand4 = new OracleCommand(dropqry4, con);
                            oracommand4.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        string ckPull = $"PULL_TICKET_NUMBER = '{DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value}' AND LINE ='{DataGridView3.Rows[cr].Cells["LINE"].Value}'";
                        DataRow[] p04 = historyBack.Select(ckPull);
                        if (p04.Length == 1)
                        {
                            if (Convert.ToInt32(DataGridView3.Rows[cr].Cells["Quantity_delivered"].Value) == 0)
                            {
                                if (p04[0]["HISTORY"] == DBNull.Value)
                                {
                                    hisBack = "";
                                }
                                else
                                {
                                    hisBack = p04[0]["HISTORY"].ToString();
                                }
                            }
                            else
                            {
                                if (p04[0]["HISTORY"] == DBNull.Value)
                                {
                                    hisBack = quanDel + " PCS - " + DateTime.Now.ToString("MMM-dd-yy").ToUpper();
                                }
                                else
                                {
                                    hisBack = p04[0]["HISTORY"] + ", " + quanDel + " PCS - " + DateTime.Now.ToString("MMM-dd-yy").ToUpper();
                                }
                            }
                        }
                        else
                        {
                            if (int.Parse(DataGridView3.Rows[cr].Cells["Quantity_delivered"].Value.ToString()) == 0)
                            {
                                hisBack = "";
                            }
                            else
                            {
                                hisBack = quanDel + " PCS - " + DateTime.Now.ToString("MMM-dd-yy").ToUpper();
                            }
                        }
                        balanceBack = Convert.ToInt32(DataGridView3.Rows[cr].Cells["ORIGINAL_PULL"].Value) - (Convert.ToInt32(DataGridView3.Rows[cr].Cells["QTY_DEL"].Value) + quanDel);//computes the remaining balance
                                                                                                                                                                                    //insert the data into BACKLOG datatable in database
                        string dropqry4 = $"insert into BACKLOG (Trans_date,Prod_date,Prod_time,Del_date,Del_time,Job_no,Facility,Partnumber,Pull_qty,Open_qty,Stock_qty,Quantity_delivered,End_Balance,GO_Number,SKU_ASSEMBLY,GO_Line_Number,CELL_NUM,REMARKS,PULL_TICKET_NUMBER,LINE,VENDOR_REMARKS, QTY_DEL, ORIGINAL_PULL, HISTORY,BALANCE) values ('" + DataGridView3.Rows[cr].Cells["Trans_date"].Value + "','" + DataGridView3.Rows[cr].Cells["Prod_date"].Value + "','" + DataGridView3.Rows[cr].Cells["Prod_time"].Value + "','" + DataGridView3.Rows[cr].Cells["Del_date"].Value.ToString().Replace("  12:00:00 AM", "") + "','" + DataGridView3.Rows[cr].Cells["Del_time"].Value + "','" + DataGridView3.Rows[cr].Cells["Job_no"].Value + "','" + DataGridView3.Rows[cr].Cells["Facility"].Value + "','" + DataGridView3.Rows[cr].Cells["Partnumber"].Value + "','" + DataGridView3.Rows[cr].Cells["Pull_qty"].Value + "','" + DataGridView3.Rows[cr].Cells["Open_qty"].Value + "','" + DataGridView3.Rows[cr].Cells["Stock_qty"].Value + "','" + quanDel + "','" + endBal + "','" + DataGridView3.Rows[cr].Cells["GO_Number"].Value + "','" + DataGridView3.Rows[cr].Cells["SKU_ASSEMBLY"].Value + "','" + DataGridView3.Rows[cr].Cells["GO_Line_Number"].Value + "','" + DataGridView3.Rows[cr].Cells["CELL_NUM"].Value + "','" + DataGridView3.Rows[cr].Cells["REMARKS"].Value + "','" + DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value + "','" + DataGridView3.Rows[cr].Cells["LINE"].Value + "','" + DataGridView3.Rows[cr].Cells["VENDOR_REMARKS"].Value + "'," + (Convert.ToInt32(DataGridView3.Rows[cr].Cells["QTY_DEL"].Value) + Convert.ToInt32(quanDel)) + "," + DataGridView3.Rows[cr].Cells["ORIGINAL_PULL"].Value + ",'" + hisBack + "','" + balanceBack + "')";
                        OracleCommand oracommand4 = new OracleCommand(dropqry4, con);
                        oracommand4.ExecuteNonQuery();
                    }
                }
            }
            DataRow[] z4;
            string expr2 = "COUNT > 1";
            z4 = samePull.Select(expr2);
            if (z4.Length > 0)
            {
                int[] rowDupNum = new int[z4.Length];
                int pullQuan;
                int rowDunIndex = 0;
                int newLength = 0;
                DataRow[] backz;
                int pullStock;
                string pullGO, pullGOline;
                string hisBack;
                int balanceBack;
                /////////////////////////////////////////////////////////
                pb = 0; // for progressbar value
                for (int rw = 0; rw < z4.Length; rw++)
                {
                    int totalDel = 0;
                    totalOpen = 0;
                    f2.label2.Text = "Saving BACKLOG . . .";
                    f2.progressBar1.Maximum = z4.Length;
                    pb = pb + 1;
                    f2.progressBar1.Value = pb;
                    f2.Refresh();
                    int dg = consoBinding.Find("PULLNO", z4[rw][0]);
                    if (Convert.ToBoolean(DataGridView3.Rows[dg].Cells["cancelled"].Value) == true) //checks if the datagridviewcheckboxcolumn is checked/ selected
                    {
                        string ckPull = "PULLNO = " + z4[rw][0] + " AND PULL_QTY IS NOT NULL";
                        DataRow[] p04 = conso_final_formatTable.Select(ckPull);
                        string backHispull = "PULL_TICKET_NUMBER = '" + p04[0]["PULL_TICKET_NUMBER"] + "' AND LINE ='" + p04[0]["LINE"] + "'";
                        DataRow[] p05 = historyBack.Select(backHispull);

                        if (p05.Length == 1)
                        {
                            if (p04[0]["HISTORY"] is DBNull)
                            {
                                hisBack = "";
                            }
                            else
                            {
                                hisBack = p04[0]["HISTORY"].ToString();
                            }
                        }
                        else
                        {
                            hisBack = "";
                        }
                        balanceBack = Convert.ToInt32(p04[0]["ORIGINAL_PULL"]) - Convert.ToInt32(p04[0]["QTY_DEL"]);//computes the ramaining balance
                                                                                                                    //insert the data into BACKLOG datatable in database
                        string puTek = $"insert into BACKLOG (Trans_date,Prod_date,Prod_time,Del_date,Del_time,Job_no,Facility,Partnumber,Pull_qty,Open_qty,Stock_qty,Quantity_delivered,End_Balance,GO_Number,SKU_ASSEMBLY,GO_Line_Number,CELL_NUM,REMARKS,PULL_TICKET_NUMBER,LINE,VENDOR_REMARKS, QTY_DEL, ORIGINAL_PULL, HISTORY,BALANCE) values ('{p04[0]["Trans_date"]}','{p04[0]["Prod_date"]}','{p04[0]["Prod_time"]}','{p04[0]["Del_date"]}','{p04[0]["Del_time"]}','{p04[0]["Job_no"]}','{p04[0]["Facility"]}','{p04[0]["Partnumber"]}','{p04[0]["Pull_qty"]}','{p04[0]["Open_qty"]}','{p04[0]["Stock_qty"]}','0','{p04[0]["Stock_qty"]}','{p04[0]["GO_Number"]}','{p04[0]["SKU_ASSEMBLY"]}','{p04[0]["GO_Line_Number"]}','{p04[0]["CELL_NUM"]}','{p04[0]["REMARKS"]}','{p04[0]["PULL_TICKET_NUMBER"]}','{p04[0]["LINE"]}','{p04[0]["VENDOR_REMARKS"]}',{int.Parse(p04[0]["QTY_DEL"].ToString())},{p04[0]["ORIGINAL_PULL"]},'{hisBack}','{balanceBack}')";
                        OracleCommand checkPullcom = new OracleCommand(puTek, con);
                        checkPullcom.ExecuteNonQuery();
                        continue;
                    }
                    DataRow[] z5;
                    string expr3 = "PULLNO = " + z4[rw][0];
                    z5 = conso_final_formatTable.Select(expr3);
                    for (int a = 0; a < z5.Length; a++)
                    {
                        totalDel = totalDel + int.Parse(z5[a]["QUANTITY_DELIVERED"].ToString());
                        totalOpen = totalOpen + int.Parse(z5[a]["OPEN_QTY"].ToString());
                    }
                    string backexpr = "PULLNO = " + z4[rw][0] + " AND PULL_QTY IS NOT NULL";
                    backz = conso_final_formatTable.Select(backexpr);
                    if (backz[0]["STOCK_QTY"] == DBNull.Value)
                    {
                        pullStock = 0;
                    }
                    else
                    {
                        pullStock = Convert.ToInt32(backz[0]["STOCK_QTY"]);
                    }
                    if (backz[0]["GO_NUMBER"] == DBNull.Value)
                    {
                        pullGO = "";
                    }
                    else
                    {
                        pullGO = backz[0]["GO_NUMBER"].ToString();
                    }
                    if (backz[0]["GO_LINE_NUMBER"] == DBNull.Value)
                    {
                        pullGOline = "";
                    }
                    else
                    {
                        pullGOline = backz[0]["GO_LINE_NUMBER"].ToString();
                    }
                    pullQuan = Convert.ToInt32(backz[0]["Pull_qty"]);
                    if (backz[0]["PROD_DATE"] == DBNull.Value)
                    {
                        PROD_DATE = "";
                    }
                    else
                    {
                        PROD_DATE = backz[0]["PROD_DATE"].ToString();
                    }
                    if (backz[0]["PROD_TIME"] == DBNull.Value)
                    {
                        PROD_TIME = "";
                    }
                    else
                    {
                        PROD_TIME = backz[0]["PROD_TIME"].ToString();
                    }
                    if (backz[0]["DEL_DATE"] == DBNull.Value)
                    {
                        DEL_DATE = DateTime.Now.ToString("M/d/yyyy");
                    }
                    else
                    {
                        DEL_DATE = backz[0]["DEL_DATE"].ToString().Replace("12:00:00 AM", "");
                    }
                    if (backz[0]["DEL_TIME"] == DBNull.Value)
                    {
                        DEL_TIME = "10:00:00 AM";
                    }
                    else
                    {
                        DEL_TIME = backz[0]["DEL_TIME"].ToString();
                    }
                    if (backz[0]["JOB_NO"] == DBNull.Value)
                    {
                        JOB_NO = "";
                    }
                    else
                    {
                        JOB_NO = backz[0]["JOB_NO"].ToString();
                    }
                    if (backz[0]["REMARKS"] == DBNull.Value)
                    {
                        REMARKS = "";
                    }
                    else
                    {
                        REMARKS = backz[0]["REMARKS"].ToString();
                    }
                    if (backz[0]["VENDOR_REMARKS"] == DBNull.Value)
                    {
                        VENDOR_REMARKS = "";
                    }
                    else
                    {
                        VENDOR_REMARKS = backz[0]["VENDOR_REMARKS"].ToString();
                    }
                    QTY_DEL = int.Parse(backz[0]["QTY_DEL"].ToString());
                    ORIGINAL_PULL = int.Parse(backz[0]["ORIGINAL_PULL"].ToString());
                    if (totalDel < pullQuan)
                    {
                        string backHispull = $"PULL_TICKET_NUMBER = '{backz[0]["PULL_TICKET_NUMBER"]}' AND LINE ='{backz[0]["LINE"]}'";
                        DataRow[] p05;
                        p05 = historyBack.Select(backHispull);
                        if (p05.Length == 1)
                        {
                            if (p05[0]["HISTORY"] == DBNull.Value)
                            {
                                hisBack = totalDel + " PCS - " + DateTime.Now.ToString("MMM-dd-yy").ToUpper();
                            }
                            else
                            {
                                hisBack = p05[0]["HISTORY"] + ", " + totalDel + " PCS - " + DateTime.Now.ToString("MMM-dd-yy").ToUpper();
                            }
                        }
                        else
                        {
                            hisBack = totalDel + " PCS - " + DateTime.Now.ToString("MMM-dd-yy").ToUpper();
                        }
                        balanceBack = ORIGINAL_PULL - (QTY_DEL + totalDel); //computes the remaining balance
                                                                          //insert the data into BACKLOG datatable in database
                        string dropqry4 = $"insert into BACKLOG (Trans_date,Prod_date,Prod_time,Del_date,Del_time,Job_no,Facility,Partnumber,Pull_qty,Open_qty,Stock_qty,Quantity_delivered,End_Balance,GO_Number,SKU_ASSEMBLY,GO_Line_Number,CELL_NUM,REMARKS,PULL_TICKET_NUMBER,LINE,VENDOR_REMARKS,QTY_DEL,ORIGINAL_PULL,HISTORY,BALANCE) values ('{backz[0]["Trans_date"]}','{PROD_DATE}','{PROD_TIME}','{DEL_DATE}','{DEL_TIME}','{JOB_NO}','{backz[0]["Facility"]}','{backz[0]["Partnumber"]}','{pullQuan}','{totalOpen}','{pullStock}','{totalDel}','{(totalDel - pullQuan)}','{pullGO}','{backz[0]["SKU_ASSEMBLY"]}','{pullGOline}','{backz[0]["CELL_NUM"]}','{REMARKS}','{backz[0]["PULL_TICKET_NUMBER"]}','{backz[0]["LINE"]}','{VENDOR_REMARKS}',{QTY_DEL + totalDel},{ORIGINAL_PULL},'{hisBack}','{balanceBack}')";
                        OracleCommand oracommand4 = new OracleCommand(dropqry4, con);
                        oracommand4.ExecuteNonQuery();
                    }
                }
            }
            // for UL labels
            DataTable ulDataTbl = new DataTable();
            string ulData = $"SELECT * FROM UL_LABEL";
            OracleDataAdapter ulDataAd = new OracleDataAdapter(ulData, con);
            ulDataAd.Fill(ulDataTbl);
            int delCounter = 0;
            pb = 0; // for progressbar value
            for (int cr = 0; cr < DataGridView3.Rows.Count; cr++)
            {
                f2.label2.Text = "Processing items for delivery . . .";
                f2.progressBar1.Maximum = DataGridView3.Rows.Count;
                pb = pb + 1;
                f2.progressBar1.Value = pb;
                f2.Refresh();
                if (int.Parse(DataGridView3.Rows[cr].Cells["Quantity_delivered"].Value.ToString()) > 0 &&
    Convert.ToBoolean(DataGridView3.Rows[cr].Cells["cancelled"].Value) == false)
                {
                    delCounter = delCounter + 1;
                    if (DataGridView3.Rows[cr].Cells["QTY_DEL"].Value == DBNull.Value &&
    DataGridView3.Rows[cr].Cells["ORIGINAL_PULL"].Value == DBNull.Value)
                    {
                        DataRow pq = conso_delTable.NewRow();
                        pq["TRANS_DATE"] = DataGridView3.Rows[cr].Cells["TRANS_DATE"].Value;
                        pq["PROD_DATE"] = DataGridView3.Rows[cr].Cells["PROD_DATE"].Value;
                        pq["PROD_TIME"] = DataGridView3.Rows[cr].Cells["PROD_TIME"].Value;
                        pq["DEL_DATE"] = DataGridView3.Rows[cr].Cells["DEL_DATE"].Value;
                        pq["DEL_TIME"] = DataGridView3.Rows[cr].Cells["DEL_TIME"].Value;
                        pq["JOB_NO"] = DataGridView3.Rows[cr].Cells["JOB_NO"].Value;
                        pq["FACILITY"] = DataGridView3.Rows[cr].Cells["FACILITY"].Value;
                        pq["PARTNUMBER"] = DataGridView3.Rows[cr].Cells["PARTNUMBER"].Value;
                        pq["PULL_QTY"] = DataGridView3.Rows[cr].Cells["PULL_QTY"].Value;
                        pq["OPEN_QTY"] = DataGridView3.Rows[cr].Cells["OPEN_QTY"].Value;
                        pq["STOCK_QTY"] = DataGridView3.Rows[cr].Cells["STOCK_QTY"].Value;
                        pq["QUANTITY_FOR_DELIVERY"] = DataGridView3.Rows[cr].Cells["QUANTITY_DELIVERED"].Value;
                        pq["END_BALANCE"] = DataGridView3.Rows[cr].Cells["END_BALANCE"].Value;
                        pq["GO_NUMBER"] = DataGridView3.Rows[cr].Cells["GO_NUMBER"].Value;
                        pq["SKU_ASSEMBLY"] = DataGridView3.Rows[cr].Cells["SKU_ASSEMBLY"].Value;
                        pq["GO_LINE_NUMBER"] = DataGridView3.Rows[cr].Cells["GO_LINE_NUMBER"].Value;
                        pq["CELL_NUM"] = DataGridView3.Rows[cr].Cells["CELL_NUM"].Value;
                        pq["REMARKS"] = DataGridView3.Rows[cr].Cells["REMARKS"].Value;
                        pq["PULL_TICKET_NUMBER"] = DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value;
                        pq["LINE"] = DataGridView3.Rows[cr].Cells["LINE"].Value;
                        pq["VENDOR_REMARKS"] = DataGridView3.Rows[cr].Cells["VENDOR_REMARKS"].Value;
                        pq["DELNO"] = delCounter;
                        conso_delTable.Rows.Add(pq);
                        conso_delTable.AcceptChanges();
                    }
                    else
                    {
                        DataRow pq = conso_delTable.NewRow();
                        pq["TRANS_DATE"] = DataGridView3.Rows[cr].Cells["TRANS_DATE"].Value;
                        pq["PROD_DATE"] = DataGridView3.Rows[cr].Cells["PROD_DATE"].Value;
                        pq["PROD_TIME"] = DataGridView3.Rows[cr].Cells["PROD_TIME"].Value;
                        pq["DEL_DATE"] = DataGridView3.Rows[cr].Cells["DEL_DATE"].Value;
                        pq["DEL_TIME"] = DataGridView3.Rows[cr].Cells["DEL_TIME"].Value;
                        pq["JOB_NO"] = DataGridView3.Rows[cr].Cells["JOB_NO"].Value;
                        //for amarok and kanban
                        if (backlogDel == false)
                        {
                            amarokTbl.CaseSensitive = false;
                            string amaSel = $"Pull_Ticket_No = '{DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value}' AND Line ='{DataGridView3.Rows[cr].Cells["LINE"].Value}'";
                            DataRow[] amaSelRow = amarokTbl.Select(amaSel);

                            if (amaSelRow.Length == 0)
                            {
                                pq["FACILITY"] = DataGridView3.Rows[cr].Cells["FACILITY"].Value;
                            }
                            else
                            {
                                if (amaSelRow[0][11].ToString().ToUpper().Contains("AMAROK"))
                                {
                                    if (amaSelRow[0][12].ToString().ToUpper().Contains("KANBAN"))
                                    {
                                        pq["FACILITY"] = "IPAI4-KANBAN";
                                    }
                                    else
                                    {
                                        pq["FACILITY"] = "IPAI4";
                                    }
                                }
                                else
                                {
                                    if (amaSelRow[0][12].ToString().ToUpper().Contains("KANBAN"))
                                    {
                                        pq["FACILITY"] = DataGridView3.Rows[cr].Cells["FACILITY"].Value + "-KANBAN";
                                    }
                                    else
                                    {
                                        pq["FACILITY"] = DataGridView3.Rows[cr].Cells["FACILITY"].Value;
                                    }
                                }
                            }
                        }
                        else
                        {
                            pq["FACILITY"] = DataGridView3.Rows[cr].Cells["FACILITY"].Value;
                        }
                        pq["PARTNUMBER"] = DataGridView3.Rows[cr].Cells["PARTNUMBER"].Value;
                        pq["PULL_QTY"] = DataGridView3.Rows[cr].Cells["PULL_QTY"].Value;
                        pq["OPEN_QTY"] = DataGridView3.Rows[cr].Cells["OPEN_QTY"].Value;
                        pq["STOCK_QTY"] = DataGridView3.Rows[cr].Cells["STOCK_QTY"].Value;
                        pq["QUANTITY_FOR_DELIVERY"] = DataGridView3.Rows[cr].Cells["QUANTITY_DELIVERED"].Value;
                        pq["END_BALANCE"] = DataGridView3.Rows[cr].Cells["END_BALANCE"].Value;
                        pq["GO_NUMBER"] = DataGridView3.Rows[cr].Cells["GO_NUMBER"].Value;
                        pq["SKU_ASSEMBLY"] = DataGridView3.Rows[cr].Cells["SKU_ASSEMBLY"].Value;
                        pq["GO_LINE_NUMBER"] = DataGridView3.Rows[cr].Cells["GO_LINE_NUMBER"].Value;
                        pq["CELL_NUM"] = DataGridView3.Rows[cr].Cells["CELL_NUM"].Value;
                        pq["REMARKS"] = DataGridView3.Rows[cr].Cells["REMARKS"].Value;
                        pq["PULL_TICKET_NUMBER"] = DataGridView3.Rows[cr].Cells["PULL_TICKET_NUMBER"].Value;
                        pq["LINE"] = DataGridView3.Rows[cr].Cells["LINE"].Value;
                        pq["VENDOR_REMARKS"] = DataGridView3.Rows[cr].Cells["VENDOR_REMARKS"].Value;
                        pq["DELNO"] = delCounter;
                        pq["QTY_DEL"] = DataGridView3.Rows[cr].Cells["QTY_DEL"].Value;
                        pq["ORIGINAL_PULL"] = DataGridView3.Rows[cr].Cells["ORIGINAL_PULL"].Value;
                        DataRow[] ulRow;
                        string ulExp = "PARTNUMBER LIKE '%" + DataGridView3.Rows[cr].Cells["PARTNUMBER"].Value.ToString().ToUpper() + "%' or PARTNUMBER = '" + DataGridView3.Rows[cr].Cells["PARTNUMBER"].Value.ToString().ToUpper() + "'";
                        ulRow = ulDataTbl.Select(ulExp); // selects row in the UL datatable
                        if (ulRow.Length > 0)
                        {
                            pq["UL_LABEL"] = ulRow[0][1];
                        }
                        conso_delTable.Rows.Add(pq);
                        conso_delTable.AcceptChanges();
                    }
                }
            }
            /////////////////////////creating excel files////////////////////////
            date_time_file = DateTime.Now.ToString("yyyy_MM_dd (hhmmtt)");
            //Creates excel file for export
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wBook = excel.Workbooks.Add();
            Worksheet wSheet = (Worksheet)wBook.Sheets[1];
            //Style style;
            int r = 0; // for progress bar
            /////////////////// for clearing excel sheets //////////////////
            for (int st = excel.Application.Worksheets.Count; st >= 2; st--)
            {
                ((Worksheet)wBook.Sheets[st]).Delete();
            }
            //////////////////////////////////////////////////////////
            DataSet dset = new DataSet();
            // add table to dataset
            dset.Tables.Add();
            // added codes for facility sheets in excel
            DataTable sameFacilityTbl = new DataTable();
            sameFacilityTbl = conso_delTable.DefaultView.ToTable(true, "FACILITY");// selects distinct facility name
            string[] sheetname = new string[sameFacilityTbl.Rows.Count];
            for (int sn = 0; sn < sameFacilityTbl.Rows.Count; sn++)
            {
                sheetname[sn] = sameFacilityTbl.Rows[sn][0].ToString();
            }
            ///////////////////////////////////////////////////////////////////////////
            ////for transaction date //////////////////////////////////
            string transdatePrint = DateTime.Now.ToString("MMM d, yyyy hh:mm tt");
            ///////////////////////////////////////////////////////////////
            for (int x = 0; x < sheetname.Length; x++) // loops through the facility name to insert the data per sheet
            {
                f2.label2.Text = "Creating excel file . . .";
                f2.progressBar1.Maximum = sheetname.Length;
                r = r + 1;
                f2.progressBar1.Value = r;
                f2.Refresh();
                int count = -1;
                DataTable selectTable = new DataTable(); // creates a new datatable for filtered rows
                DataView negaView = new DataView(conso_delTable); // creates a dataview for the final pull ticket datatable
                negaView.RowFilter = "FACILITY='" + sheetname[x] + "'"; // filters the data base on facility name
                selectTable = negaView.ToTable(); // saved the filtered rows to a datatable
                selectTable = selectTable.DefaultView.ToTable();
                // Iterate through the rows in the DataTable
                for (int i = 1; i < selectTable.Rows.Count; i++)
                {
                    // Check if the current row's delivery date is null or empty
                    if (selectTable.Rows[i]["DEL_DATE"] == null || string.IsNullOrEmpty(selectTable.Rows[i]["DEL_DATE"].ToString()))
                    {
                        // Get the previous row's delivery date
                        if (i > 0 && selectTable.Rows[i - 1]["DEL_DATE"] != null)
                        {
                            string prevDeliveryDate = selectTable.Rows[i - 1]["DEL_DATE"].ToString();
                            // Update the current row's delivery date with the previous row's delivery date
                            selectTable.Rows[i]["DEL_DATE"] = prevDeliveryDate;
                        }
                    }
                }
                selectTable.DefaultView.Sort = "DELNO ASC, PARTNUMBER ASC"; // sorts the data
                selectTable = selectTable.DefaultView.ToTable();

                selectTable = selectTable.DefaultView.ToTable(false, "DEL_DATE", "DEL_TIME", "PARTNUMBER", "PULL_QTY", "QUANTITY_FOR_DELIVERY", "PULL_TICKET_NUMBER", "LINE", "GO_NUMBER", "GO_LINE_NUMBER", "TRANS_DATE", "REMARKS", "ORIGINAL_PULL", "QTY_DEL", "CELL_NUM", "UL_LABEL");
                selectTable.Columns.Add("BINCARD", typeof(string)); // adds column in the datatable for bincard
                selectTable = selectTable.DefaultView.ToTable(false, "DEL_DATE", "DEL_TIME", "PARTNUMBER", "PULL_QTY", "BINCARD", "QUANTITY_FOR_DELIVERY", "PULL_TICKET_NUMBER", "LINE", "GO_NUMBER", "GO_LINE_NUMBER", "CELL_NUM", "TRANS_DATE", "REMARKS", "ORIGINAL_PULL", "QTY_DEL", "UL_LABEL"); //////////////// OCT 4 ////////////
                DataRow[] z2 = null;
                DataRow[] z3 = null;
                DataRow[] poBin;
                DataTable samepartTbl = selectTable.DefaultView.ToTable(true, "GO_NUMBER", "PULL_TICKET_NUMBER", "LINE");
                //for bincard
                for (int a = 0; a < samepartTbl.Rows.Count; a++)
                {
                    string poProd = "SO_NO = '" + samepartTbl.Rows[a]["GO_NUMBER"] + "'";
                    poBin = colorPOsm.Select(poProd);
                    if (poBin.Length >= 1)
                    {
                        string bincard = "[1] = '" + poBin[0]["PRODUCT_NO"] + "'";
                        z2 = invTblCopy.Select(bincard);
                        string bincard2 = "GO_NUMBER = '" + samepartTbl.Rows[a]["GO_NUMBER"] + "' and PULL_TICKET_NUMBER = '" + samepartTbl.Rows[a]["PULL_TICKET_NUMBER"] + "' and LINE = '" + samepartTbl.Rows[a]["LINE"] + "'";
                        z3 = selectTable.Select(bincard2);
                    }
                    if (z2.Length > 0)
                    {
                        z3[0]["BINCARD"] = z2[0]["16"];
                    }
                }
                /////////////////////////////
                //checks for excel sheets and assigns a name to it
                if (excel.Application.Sheets.Count < x + 1)
                {
                    wSheet = (Worksheet)wBook.Worksheets.Add();
                }
                else
                {
                    wSheet = (Worksheet)excel.Worksheets[x + 1];
                }
                /////////////////////////////////////// OCT. 10, 2024
                if (sheetname[x].Equals("CLP/CAV5"))
                {
                    wSheet.Name = "CLP-CAV5";
                }
                else if (sheetname[x].Equals("CAV3/JAC"))
                {
                    wSheet.Name = "CAV3-JAC";
                }
                else
                {
                    wSheet.Name = sheetname[x];
                }
                /////////////////////////////////////// OCT. 10, 2024



                //wSheet.Name = sheetname[x];
                // for page layout orientation //////////////////////
                try
                {
                    wSheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                }
                catch (Exception ex) { }
                //////////////////////////////////////////////
                // adds values to the excel cells
                Range labelCell1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 1]];
                labelCell1.Value = "PULL TICKET SCHEDULE";
                labelCell1.Font.Size = 20;
                labelCell1.Font.Bold = true;
                Microsoft.Office.Interop.Excel.Range labelCell2 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[2, 1]];
                labelCell2.Value = "FACILITY: " + sheetname[x];
                labelCell2.Font.Size = 20;
                labelCell2.Font.Bold = true;
                //////////////////////////////////////////////
                DataTable dt = dset.Tables[0];
                DataColumn dc;
                int colIndex = 0;
                int rowIndex = 2;
                int nextRowIndex = 3;
                object[,] arr; // creates an array object to hold the data
                arr = new object[selectTable.Rows.Count, selectTable.Columns.Count];
                // adds values to cells
                colIndex++;
                excel.Cells[3, colIndex] = "DELIVERY DATE";
                excel.Cells[3, colIndex + 1] = "DELIVERY TIME";
                excel.Cells[3, colIndex + 2] = "PART NUMBER";
                excel.Cells[3, colIndex + 3] = "BALANCE / PULL";
                excel.Cells[3, colIndex + 3].ColumnWidth = 10;
                excel.Cells[3, colIndex + 4] = "BINCARD";
                excel.Cells[3, colIndex + 5] = "QTY FOR DELIVERY";
                excel.Cells[3, colIndex + 5].ColumnWidth = 10;
                excel.Cells[3, colIndex + 6] = "PULL TICKET NUMBER";
                excel.Cells[3, colIndex + 7] = "LINE";
                excel.Cells[3, colIndex + 8] = "GO NUMBER";
                excel.Cells[3, colIndex + 9] = "GO LINE NUMBER";
                excel.Cells[3, colIndex + 10] = "CELL NUMBER";
                excel.Cells[3, colIndex + 11] = "TRANSACTION DATE";
                excel.Cells[3, colIndex + 12] = "REMARKS";
                excel.Cells[3, colIndex + 13] = "ORIGINAL PULL";
                excel.Cells[3, colIndex + 13].ColumnWidth = 10;
                excel.Cells[3, colIndex + 14] = "QTY DELIVERED";
                excel.Cells[3, colIndex + 15] = "UL LABEL";
                excel.Cells[3, colIndex + 16] = "QTY BREAKDOWN";
                excel.Cells[3, colIndex + 16].ColumnWidth = 25;
                excel.Cells[3, colIndex + 17] = "UL SERIES NO.";
                excel.Cells[3, colIndex + 18] = "UL DENOMINATION";
                excel.Cells[3, colIndex + 19] = "DR NO.";
                // formats the cell
                Range formatRange2 = wSheet.UsedRange;
                Range cell = formatRange2.Range[wSheet.Cells[3, colIndex], wSheet.Cells[3, colIndex + 19]];
                Borders border = cell.Borders;
                border.Weight = 2.0;
                cell.Font.Size = 10;
                cell.Font.Bold = true;
                cell.WrapText = true;
                cell.RowHeight = 30;
                cell.EntireRow.Font.Bold = true;
                cell.Interior.ColorIndex = 20;
                /////////////////////////////////////////////////////////////////
                Range smcell; // creates a new Excel range
                for (int r2 = 0; r2 < selectTable.Rows.Count; r2++)
                {
                    DataRow dr = selectTable.Rows[r2]; // selects a row in the datatable
                    for (int c = 0; c < selectTable.Columns.Count; c++)
                    {
                        arr[r2, c] = dr[c]; // inserts data to the array
                    }
                    // selects rows that contain "SM" to change the color of cells in the excel file
                    string outputSM = "SO_NO = '" + selectTable.Rows[r2]["GO_NUMBER"] + "'";
                    DataRow[] outputSMrow = colorPOsm.Select(outputSM);
                    if (outputSMrow.Length > 0)
                    {
                        if (outputSMrow[0][0].ToString().Contains("SM"))
                        {
                            smcell = wSheet.Range[wSheet.Cells[4 + r2, 1], wSheet.Cells[4 + r2, 20]];
                            smcell.Interior.ColorIndex = 40;
                        }
                    }
                    //////for A, B, C partnumber highlighting ////////////////
                    if (selectTable.Rows[r2].Field<string>("PARTNUMBER").ToUpper().Contains("A") ||
    selectTable.Rows[r2].Field<string>("PARTNUMBER").ToUpper().Contains("B") ||
    selectTable.Rows[r2].Field<string>("PARTNUMBER").ToUpper().Contains("C"))
                    {
                        smcell = wSheet.Range[wSheet.Cells[4 + r2, 1], wSheet.Cells[4 + r2, 20]];
                        smcell.Interior.ColorIndex = 43;
                    }
                }
                Range c1 = wSheet.Cells[4, 1];
                Range c2 = wSheet.Cells[4 + selectTable.Rows.Count - 1, selectTable.Columns.Count];
                Range range = wSheet.Range[c1, c2];
                range.Value = arr;
                //formats the cell
                Range formatRange3 = wSheet.UsedRange;
                Range cell2 = formatRange3.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + selectTable.Rows.Count, 20]];
                Borders border2 = cell2.Borders;
                border2.LineStyle = XlLineStyle.xlContinuous;
                border2.Weight = 2.0;
                cell2.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                cell2.VerticalAlignment = XlVAlign.xlVAlignCenter;
                Range ulCell = wSheet.Range[wSheet.Cells[4, 16], wSheet.Cells[4 + selectTable.Rows.Count, 16]];
                ulCell.Font.Size = 7.5;
                ulCell.WrapText = true;
                ulCell.ColumnWidth = 18;
                /////////////////////////////////
                wSheet.Columns.AutoFit();
                wSheet.Move(wBook.Worksheets[wBook.Worksheets.Count]); //arranges the sheets in order
                //deletes the data with the same facility name in the loop
                string expr = "FACILITY='" + sheetname[x] + "'";
                DataRow[] prt;
                prt = conso_delTable.Select(expr);
                foreach (DataRow prt2 in prt)
                {
                    prt2.Delete();
                }
                conso_delTable.AcceptChanges();
                if (x == sheetname.Length - 1)
                {
                    f2.label2.Text = "DONE!";
                    f2.Refresh();
                }
            }
            //creates a backup of the database after printing
            try
            {
                string fulldate = DateTime.Now.ToString("MMM-dd-yyyy-hh-mm-ss") + "-AFTER-PRINTING";
                Process myprocess = new Process();
                System.Diagnostics.ProcessStartInfo StartInfo = new System.Diagnostics.ProcessStartInfo();
                StartInfo.FileName = "cmd"; //starts cmd window
                StartInfo.RedirectStandardInput = true;
                StartInfo.RedirectStandardOutput = true;
                StartInfo.UseShellExecute = false; //required to redirect
                StartInfo.CreateNoWindow = true;
                myprocess.StartInfo = StartInfo;
                myprocess.Start();
                System.IO.StreamReader SR = myprocess.StandardOutput;
                System.IO.StreamWriter SW = myprocess.StandardInput;
                SW.WriteLine(@"cd C:\oraclexe\app\oracle\product\10.2.0\server\BIN"); //the command you wish to run.....
                SW.WriteLine(@"exp mec/mec2024@192.168.50.40 buffer=4096 grants=Y file=\\192.168.50.40\jmctest\" + fulldate + ".dmp tables=(BACKLOG, MANUAL_DR, PULL_QUANTITY_REV, PULL_TICKET_RECORD)");
                SW.WriteLine("exit"); //exits command prompt window
                int procID = myprocess.Id;
                while (ProcessExists(procID))
                {
                    f2.label2.Text = "Backing up data . . .";
                    f2.Refresh();
                }
                SW.Close();
                SR.Close();
            }
            catch (Exception ex) { }
            //for inserting and hiding of cell columns for exports
            int scount1 = wBook.Sheets.Count;
            for (int a = 1; a <= scount1; a++)
            {
                f2.label2.Text = "Creating excel file . . .";
                f2.progressBar1.Maximum = scount1;
                f2.progressBar1.Value = a;
                f2.Refresh();
                wSheet = (Worksheet)wBook.Sheets[a];
                wSheet.ResetAllPageBreaks();
                wSheet.VPageBreaks.Add(wSheet.Range["U1"]);
                wSheet.Columns["T:T"].Cut();
                wSheet.Columns["M:M"].Insert();
                Range hideRow = wSheet.Range[wSheet.Cells[3, 7], wSheet.Cells[3, 8]];
                hideRow.EntireColumn.Hidden = true;
                hideRow = wSheet.Range[wSheet.Cells[3, 12], wSheet.Cells[3, 12]];
                hideRow.EntireColumn.Hidden = true;
                hideRow = wSheet.Range[wSheet.Cells[3, 15], wSheet.Cells[3, 16]];
                hideRow.EntireColumn.Hidden = true;
                hideRow = wSheet.Range[wSheet.Cells[3, 18], wSheet.Cells[3, 20]];
                hideRow.EntireColumn.Hidden = true;
                hideRow = wSheet.Range["M3"];
                hideRow.ColumnWidth = 16;
                ///////////////////////////// for trans date ///////////////////////////////
                hideRow = wSheet.Range[wSheet.Cells[1, 13], wSheet.Cells[1, 14]];
                hideRow.Merge();
                hideRow.NumberFormat = "@";
                hideRow.Value = transdatePrint;
                hideRow = wSheet.Range["A1"];
                hideRow.ColumnWidth = 15;
                wSheet.ResetAllPageBreaks();
                wSheet.VPageBreaks.Add(wSheet.Range["U1"]);
            }
            //////////////////////////////////////////////////
            wBook.SaveAs(strFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlXMLSpreadsheet); //saves the workbook


            //for inserting and hiding of cell columns for warehouse copy
            int scount = wBook.Sheets.Count;
            for (int a = 1; a <= scount; a++)
            {
                f2.label2.Text = "Creating excel file . . .";
                f2.progressBar1.Maximum = scount;
                f2.progressBar1.Value = a;
                f2.Refresh();
                wSheet = (Worksheet)wBook.Sheets[a];
                wSheet.ResetAllPageBreaks();
                wSheet.VPageBreaks.Add(wSheet.Range["V1"]);
                Range hideRow = wSheet.Range[wSheet.Cells[3, 7], wSheet.Cells[3, 8]];
                hideRow.EntireColumn.Hidden = false;
                hideRow = wSheet.Range[wSheet.Cells[3, 12], wSheet.Cells[3, 12]];
                hideRow.EntireColumn.Hidden = true;
                hideRow = wSheet.Range[wSheet.Cells[3, 15], wSheet.Cells[3, 16]];
                hideRow.EntireColumn.Hidden = false;
                hideRow = wSheet.Range[wSheet.Cells[3, 18], wSheet.Cells[3, 20]];
                hideRow.EntireColumn.Hidden = false;
                hideRow = wSheet.Range[wSheet.Cells[3, 7], wSheet.Cells[3, 13]];
                hideRow.EntireColumn.Hidden = true;
                hideRow = wSheet.Range[wSheet.Cells[3, 15], wSheet.Cells[3, 16]];
                hideRow.EntireColumn.Hidden = true;
                wSheet.Columns["Q:Q"].Cut();
                wSheet.Columns["U:U"].Insert();
                hideRow = wSheet.Range["N3"];
                hideRow.ColumnWidth = 0;
                hideRow = wSheet.Range["Q3"];
                hideRow.ColumnWidth = 23;
                hideRow = wSheet.Range["R3"];
                hideRow.ColumnWidth = 10;
                hideRow = wSheet.Range["S3"];
                hideRow.ColumnWidth = 14;
                hideRow = wSheet.Range["T3"];
                hideRow.ColumnWidth = 18;


                int lastRow = wSheet.UsedRange.Rows.Count;
                ////////// Add "Prepared by" and "Checked by" labels /////////////////////

                hideRow = wSheet.Cells[lastRow, 5];
                hideRow.Value = "PREPARED BY:";
                hideRow.Font.Size = 10;
                hideRow.Font.Bold = true;

                hideRow = wSheet.Cells[lastRow, 18];
                hideRow.Value = "CHECKED BY:";
                hideRow.Font.Size = 10;
                hideRow.Font.Bold = true;

                wSheet.Columns["G:G"].Insert();
                wSheet.Cells[3, 7].Value = "TOTAL";




                ////////////////////// MERGING TOTAL VALUES //////////////////////////////////////////////

                Range criteriaRange = wSheet.Range[wSheet.Cells[4, 3], wSheet.Cells[lastRow, 3]];
                Range sumRange = wSheet.Range[wSheet.Cells[4, 6], wSheet.Cells[lastRow, 6]];

                string prevValue = "";
                int startRow = 4;
                int endRow = 4;

                double sum;
                Range mergeRange;
               //DateTime prevDate;
                /*if (!DateTime.TryParse(wSheet.Cells[startRow, 4].Value.ToString(), out prevDate))
                {
                    prevDate = DateTime.MinValue;
                }*/
                int prevStartRow = 4;


                for (int i = startRow; i < lastRow; i++)
                {
                    string currentValue = wSheet.Cells[i, 3].Value.ToString();
                   // DateTime currentDate;
                    object cellValue = wSheet.Cells[i, 4].Value;

                    /*if (cellValue != null)
                    {
                        bool isValidDate = DateTime.TryParse(cellValue.ToString(), out currentDate);
                    }
                    else
                    {
                        // Handle the case where the cell value is null
                        currentDate = DateTime.MinValue;
                    }*/

                    if (currentValue != prevValue /*|| currentDate != prevDate*/)
                    {
                        if (i > startRow)
                        {
                            // Calculate the sum for the previous group
                            Range groupSumRange = wSheet.Range[wSheet.Cells[prevStartRow, 6], wSheet.Cells[endRow, 6]];
                            sum = (double)wSheet.Application.WorksheetFunction.Sum(groupSumRange);

                            // Write the sum in the first cell of the group
                            wSheet.Cells[startRow, 7].Value = sum;

                            // Merge cells in Column S
                            mergeRange = wSheet.Range[wSheet.Cells[startRow, 7], wSheet.Cells[endRow, 7]];
                            mergeRange.Merge();
                        }
                        startRow = i;
                        prevStartRow = i;
                        prevValue = currentValue;
                        //prevDate = currentDate;
                    }
                    endRow = i;
                }

                // Calculate the sum for the last group
                Range lastGroupSumRange = wSheet.Range[wSheet.Cells[prevStartRow, 6], wSheet.Cells[endRow, 6]];
                sum = (double)wSheet.Application.WorksheetFunction.Sum(lastGroupSumRange);

                // Write the sum in the first cell of the group
                wSheet.Cells[startRow, 7].Value = sum;

                // Merge cells in Column S for the last group
                mergeRange = wSheet.Range[wSheet.Cells[startRow, 7], wSheet.Cells[endRow, 7]];
                mergeRange.Merge();



                ///////////////////////////// for trans date ///////////////////////////////
                hideRow = wSheet.Range["R1"];
                hideRow.Value = "";
                hideRow = wSheet.Range[wSheet.Cells[1, 19], wSheet.Cells[1, 20]];
                hideRow.Merge();
                hideRow.NumberFormat = "@";
                hideRow.Value = transdatePrint;
                wSheet.ResetAllPageBreaks();
                wSheet.VPageBreaks.Add(wSheet.Range["V1"]);
            }
            string warehouse = strFileName.Replace(".xlsx", "").Replace(".xls", "") + "-WAREHOUSE.xls";
            wBook.SaveAs(warehouse, Microsoft.Office.Interop.Excel.XlFileFormat.xlXMLSpreadsheet);


            //for inserting and hiding of cell columns for uploading
            int scount2 = wBook.Sheets.Count;

            for (int a = 1; a <= scount2; a++)
            {
                f2.label2.Text = "Creating excel file . . .";
                f2.progressBar1.Maximum = scount2;
                f2.progressBar1.Value = a;
                f2.Refresh();
                wSheet = (Worksheet)wBook.Sheets[a];
                Range hideRow = wSheet.Range[wSheet.Cells[1, 13], wSheet.Cells[1, 14]];
                hideRow.Value = "";
                hideRow.UnMerge();
                hideRow = wSheet.Range[wSheet.Cells[1, 18], wSheet.Cells[1, 19]];
                hideRow.Value = "";
                hideRow.UnMerge();
                hideRow = wSheet.Range[wSheet.Cells[3, 7], wSheet.Cells[3, 13]];
                hideRow.EntireColumn.Hidden = false;
                hideRow = wSheet.Range[wSheet.Cells[3, 15], wSheet.Cells[3, 16]];
                hideRow.EntireColumn.Hidden = false;
                wSheet = (Worksheet)wBook.Sheets[a];
                wSheet.ResetAllPageBreaks();
                wSheet.VPageBreaks.Add(wSheet.Range["U1"]);
                wSheet.Columns["G:G"].Delete();
                wSheet.Columns["O:O"].Cut();
                wSheet.Columns["D:D"].Insert();
                wSheet.Columns["H:H"].Cut();
                wSheet.Columns["E:E"].Insert();
                wSheet.Columns["I:I"].Cut();
                wSheet.Columns["F:F"].Insert();
                wSheet.Columns["I:I"].Cut();
                wSheet.Columns["G:G"].Insert();
                wSheet.Columns["M:M"].Cut();
                wSheet.Columns["H:H"].Insert();
                wSheet.Columns["N:N"].Cut();

                wSheet.Columns["I:I"].Insert();
                wSheet.Columns["L:L"].Cut();
                wSheet.Columns["J:J"].Insert();
                wSheet.Columns["M:M"].Cut();
                wSheet.Columns["K:K"].Insert();

                hideRow = wSheet.Range[wSheet.Cells[3, 9], wSheet.Cells[3, 9]];
                hideRow.EntireColumn.Hidden = false;
                hideRow = wSheet.Range[wSheet.Cells[3, 12], wSheet.Cells[3, 13]];
                hideRow.EntireColumn.Hidden = true;
                hideRow = wSheet.Range[wSheet.Cells[3, 14], wSheet.Cells[3, 19]];
                hideRow.EntireColumn.Hidden = true;
                ///////////////////////////// feb 21 for trans date ///////////////////////////////
                hideRow = wSheet.Range["R1"];
                hideRow.Value = "";
                hideRow = wSheet.Range[wSheet.Cells[1, 8], wSheet.Cells[1, 9]];
                hideRow.Merge();
                hideRow.NumberFormat = "@";
                hideRow.Value = transdatePrint;
                wSheet.ResetAllPageBreaks();
                wSheet.VPageBreaks.Add(wSheet.Range["U1"]);
            }
            string upload = strFileName.Replace(".xlsx", "").Replace(".xls", "").Replace("WAREHOUSE", "") + "-UPLOAD.xls";
            wBook.SaveAs(upload, Microsoft.Office.Interop.Excel.XlFileFormat.xlXMLSpreadsheet);
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            f2.Close(); //closes the form that shows the progress bar and status text
            //////////////////////////////////////////////////////////////////////////////////////////////////
            MessageBox.Show("Excel file created!"); //notify the user that the excel files were created
            //opens the created excel files
            excel.Workbooks.Open(strFileName);
            excel.Workbooks.Open(warehouse);
            excel.Workbooks.Open(upload);
            excel.Visible = true;
            ///////////////////////resetting of controls////////////////////////////
            if (genClick == true || asperclick == true)
            {
                DataGridView3.Columns.Remove("cancelled");
            }
            if (zoomClick == true)
            {
                dataGridView4.Columns.Remove("cancelled");
            }
            genClick = false;
            zoomClick = false;
            pullrecordTable = new DataTable();
            backlogDel = false;
            forFilter = new DataTable();
            pullticketfilter = false;
            dataGridView4.DataSource = null;
            DataGridView3.DataSource = null;
            pulltktgrid.DataSource = null;
            pullQuanTempTable = new DataTable();
            pullQuantityTable = new DataTable();
            inventoryTable = new DataTable();
            pullbackTable = new DataTable();
            poTable = new DataTable();
            consolidatedTable = new DataTable();
            conso_final_formatTable = new DataTable();
            conso_delTable = new DataTable();
            backlogTempoTable = new DataTable();
            cloneQuanTempTable = new DataTable();
            manualDrTable = new DataTable();
            asperTable = new DataTable();
            asperQuanTable = new DataTable();
            backNewDelDate = new DataTable();
            pullRevTable = new DataTable();
            pullDeleteTable = new DataTable();
            manualDrTable = new DataTable();
            fgSM = new DataTable();
            poTableCopy = new DataTable();
            totalPullSM = new DataTable();
            colorPOsm = new DataTable();
            historyBack = new DataTable();
            invTblCopy = new DataTable();
            invSMorNonTbl = new DataTable();
            selAdvBackTbl = new DataTable();
            advBacklogTbl = new DataTable();
            amarokTbl = new DataTable();
            kanbanTbl = new DataTable();
            if (importPull.Equals(true) && pullQuanTempTable != null)
            {
                if (excelData3 != null && importCxmr.Equals(true))
                {
                    excelData3.Clear();
                    if (importInv.Equals(true) && invTblCopy != null)
                    {
                        invTblCopy.Clear();
                        if (importPO.Equals(true) && unpostedTbl != null)
                        {
                            unpostedTbl.Clear();
                            btncmpl.Visible = false;
                            if (poTable.Equals(null))
                            {
                                poTable.Clear();
                            }
                        }
                    }
                }
                /*pullQuanTempTable.Clear();
                backNewDelDate.Clear();
                pullrecordTable.Clear();
                pullQuantityTable.Clear();
                cloneQuanTempTable.Clear();
                pullDeleteTable.Clear();
                pullRevTable.Clear();
                manualDrTable.Clear();
                advBacklogTbl.Clear();
                amarokTbl.Clear();
                kanbanTbl.Clear();
                historyBack.Clear();
                backTempotbl.Clear();
                forFilter.Clear();
                excelData2.Clear();
                pullbackTable.Clear();*/
                btnfilter.Visible = false;
                btnpull_tkt.Enabled = true;
                btncxmr620.Enabled = true;
                btnaimr407.Enabled = true;
                btnaxmr432.Enabled = true;
                btnpull_tkt.BackColor = Color.SteelBlue;
                btncxmr620.BackColor = Color.SteelBlue;
                btnaimr407.BackColor = Color.SteelBlue;
                btnaxmr432.BackColor = Color.SteelBlue;
                btnaxmr340.BackColor = Color.SteelBlue;
                pnl_import2.Location = new System.Drawing.Point(35, 60);
                dataGridView4.Visible = false;
                pulltktgrid.Visible = false;
                slider.Visible = true;
                refesherOrb.Visible = false;
                btndownload.Visible = false;
                importPull = false;
                importCxmr = false;
                importInv = false;
                importPO = false;
                calculateBtn = false;

                genClick = false;
                asperclick = false;
                backlogDel = false;
                asperMarked = false;
                pullticketfilter = false;
                /*asperTable = new DataTable();
                pullQuanTempTable = new DataTable();
                pullQuantityTable = new DataTable();
                inventoryTable = new DataTable();
                pullbackTable = new DataTable();
                poTable = new DataTable();
                consolidatedTable = new DataTable();
                conso_final_formatTable = new DataTable();
                backTempotbl = new DataTable();
                cloneQuanTempTable = new DataTable();
                manualDrTable = new DataTable();
                asperQuanTable = new DataTable();
                backNewDelDate = new DataTable();
                pullRevTable = new DataTable();
                pullDeleteTable = new DataTable();
                fgSM = new DataTable();
                poTableCopy = new DataTable();
                totalPullSM = new DataTable();
                colorPOsm = new DataTable();*/
                consoBinding = new BindingSource();
                /*historyBack = new DataTable();
                invTblCopy = new DataTable();
                invSMorNonTbl = new DataTable();
                selAdvBackTbl = new DataTable();
                advBacklogTbl = new DataTable();
                amarokTbl = new DataTable();
                kanbanTbl = new DataTable();*/
                if (pnldeliverydate.Visible == true || pnl_deladvise.Visible == true)
                {
                    pnldeliverydate.Visible = false;
                    pnl_deladvise.Visible = false;
                }
            }
            else
            {
                MessageBox.Show("No Data. . . .", "Error!");
            }
        }
    }
}
