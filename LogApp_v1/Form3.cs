using LogApp_v1.Models;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;

namespace LogApp_v1
{

    public partial class Form3 : Form
    {
        //oracle temporary holders//
        private string connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.50.40)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User ID=mec;Password=mec2024"; //oracle host source
        private OracleConnection con = new OracleConnection(); //temporary holder for oracle connection
        private OracleCommand cmd = new OracleCommand(); //Temporary holder for oracle command
        private OracleDataAdapter adpt, adpt1, adpt2, adpt3; //Temporary holder for oracle data
        private DataTable dt, dtbacklog, dt2, dt3; //Temporary holder for datatable

        //This is the list of all the backlog data that have checked - cj
        private List<BacklogDataModel> checkedBacklogData = new List<BacklogDataModel>();
        private List<ManualDRDataModel> checkedmanualDrData = new List<ManualDRDataModel>();

        public const int WM_NCLBUTTONDOWN = 0XA1;
        public const int HTCAPTION = 0x2;
        [DllImport("User32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("User32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);



        public Form3()
        {
            InitializeComponent();
            
        }
        public void Form3_Load(object sender, EventArgs e)
        {
            connectdata(); //connect sql method
            showpullticketdata(); //display pull ticket record to pullticketdatagrid
            showbacklogdata(); // display backlog record to backlogdatgrid
            showrevision(); //display revision record to revisiondatagrid
            manualDRrecord(); //display manual dr record to revisiondatagrid

            recorddatagrid.RowHeadersVisible = false;//hide datagridviewrowheaders
            backlogdatagrid.RowHeadersVisible = false;//hide datagridviewrowheaders

            recorddatagridCustomStyle(); //custom style method for pullticketrecord datagrid
           backlogdatagridCustomStyle(); //custom style method for backlogrecord datagrid
           revisiondatagridCustomStyle();//custom style method for revisionrecord datagrid
            manualdrdatagridCustomStyle();//custom style method for manualDrrecord datagrid


            facilbox.SelectedIndex = 0; //facility combobox in pullticketrecord default index to 0
            backlogBox.SelectedIndex = 0; //backlog searchby combobox default index to 0
            backlogTxt.Enabled = false; //backlog searchtextbox disabled by default

            //Checked AM by default
            amCheckbox.Checked = true;

            totalData.Text = "TOTAL DATA: " + recorddatagrid.Rows.Count;
            totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;
            TotalRev.Text = "TOTAL DATA: " + revisiondatagrid.Rows.Count;
            datastatus();
            recorddatagrid.ClearSelection();
            pnlDeldate.Value = DateTime.Now;

            DoubleBuffered = true;
            EnableDoubleBuffering();

            ExtendedMethods.DoubleBuffered(backlogdatagrid, true);
            ExtendedMethods.DoubleBuffered(manualdrdatagrid, true);
            ExtendedMethods.DoubleBuffered(recorddatagrid, true);
            ExtendedMethods.DoubleBuffered(revisiondatagrid, true);

            mdrStatus.BorderStyle = BorderStyle.Fixed3D;
            facilbox.SelectedIndex = 0;
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

        private void datastatus()
        {
            if(revisiondatagrid.Rows.Count > 0)
            {
                TotalRev.Visible = true;
                dataStatus.Visible = false;
                
            }
            else
            {
                dataStatus.Visible = true;
                TotalRev.Visible = false;
                dataStatus.Text = "NO DATA ENTRY";

            }


            if(manualdrdatagrid.Rows.Count > 0)
            {
               totalMDR.Text = "TOTAL DATA: " + manualdrdatagrid.Rows.Count.ToString();
                mdrStatus.Visible = false;
            }
            else
            {
                totalMDR.Visible = false;
                mdrStatus.Text = "NO DATA ENTRY";
            }
        }
        //manualDrrecord Custom style method
        private void manualdrdatagridCustomStyle()
        {
            manualdrdatagrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None; //disable autosizerows
            manualdrdatagrid.AllowUserToResizeRows = false; //disable user resize rows on datagrid 
            manualdrdatagrid.EnableHeadersVisualStyles = false; //headervisualstyle disabling
            manualdrdatagrid.ColumnHeadersDefaultCellStyle.BackColor = Color.Black; // changing header bacground color. 
            manualdrdatagrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; // changing header fore color.
            manualdrdatagrid.Columns[0].Width = 90; // width size on 1st column
        }
        //custom style method for backlogdatagrid
        private void backlogdatagridCustomStyle()
        {
            backlogdatagrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            backlogdatagrid.AllowUserToResizeRows = false;
            backlogdatagrid.EnableHeadersVisualStyles = false;
            backlogdatagrid.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            backlogdatagrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            backlogdatagrid.Columns[0].Width = 90;
            backlogdatagrid.Columns[1].Width = 120;
            backlogdatagrid.Columns[2].Width = 100;
            recorddatagrid.Columns[3].Width = 120;
            backlogdatagrid.Columns[4].Width = 130;
            backlogdatagrid.Columns[5].Width =90;
            backlogdatagrid.Columns[6].Width = 120;
            backlogdatagrid.Columns[7].Width = 290;
            backlogdatagrid.Columns[8].Width = 80;
            backlogdatagrid.Columns[9].Width = 120;
            backlogdatagrid.Columns[10].Width = 120;
            backlogdatagrid.Columns[11].Width = 120;
            backlogdatagrid.Columns[12].Width = 120;
        }
        //custom style method for pullticketdatagrid
        private void recorddatagridCustomStyle()
        {
            recorddatagrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            recorddatagrid.AllowUserToResizeRows = false;
            recorddatagrid.EnableHeadersVisualStyles = false;
            recorddatagrid.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            recorddatagrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            //recorddatagrid.AlternatingRowsDefaultCellStyle.ForeColor = Color.White;
            recorddatagrid.Columns[0].Width = 120;
            recorddatagrid.Columns[1].Width = 120;
            recorddatagrid.Columns[2].Width = 120;
            recorddatagrid.Columns[3].Width = 120;
            recorddatagrid.Columns[4].Width = 120;
            recorddatagrid.Columns[5].Width = 120;
            recorddatagrid.Columns[6].Width = 120;
            recorddatagrid.Columns[7].Width = 90;
            recorddatagrid.Columns[8].Width = 120;
            recorddatagrid.Columns[9].Width = 120;
            recorddatagrid.Columns[10].Width = 120;
            recorddatagrid.Columns[11].Width = 120;
            recorddatagrid.Columns[12].Width = 120;
            recorddatagrid.Columns[13].Width = 90;
            recorddatagrid.Columns[14].Width = 120;
            recorddatagrid.Columns[15].Width = 120;
            recorddatagrid.Columns[16].Width = 120;
            recorddatagrid.Columns[18].Width = 120;
            recorddatagrid.Columns[20].Width = 120;
            recorddatagrid.Columns[21].Width = 120;
        }
        private void revisiondatagridCustomStyle()
        {
            revisiondatagrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            revisiondatagrid.AllowUserToResizeRows = false;
            revisiondatagrid.EnableHeadersVisualStyles = false;
            revisiondatagrid.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            revisiondatagrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            //revisiondatagrid.AlternatingRowsDefaultCellStyle.ForeColor = Color.White;
        }
        //connect to sql //
        private void connectdata()
        {
            con.ConnectionString = connectionString; //declaring connection from host
            con.Open(); //to open database connection
        }
        // method to show data from database to datatable//
        private void showpullticketdata()
        {
            // pull ticket record data //
            adpt = new OracleDataAdapter("SELECT * FROM PULL_TICKET_RECORD", con); //query to select specified data from database (Pull ticket record)
            dt = new DataTable();
            adpt.Fill(dt); //store data to datatable(temporary holder) from datase
            recorddatagrid.DataSource = dt;  //Throwing the data to datagridview

        }
        // backlog data table //
        private void showbacklogdata()
        {
            
            adpt1 = new OracleDataAdapter("SELECT DEL_DATE,DEL_TIME,FACILITY,PARTNUMBER,BALANCE,REMARKS,PULL_TICKET_NUMBER,LINE,QTY_DEL,ORIGINAL_PULL,BACKLOGTYPE,HISTORY FROM BACKLOG", con);//selecting specific data to sql
            dtbacklog = new DataTable(); //backlog datatable that holds the sql data
            adpt1.Fill(dtbacklog);        //throw the data to the datatable
            dtbacklog.Columns.Add(new DataColumn("SELECTED", typeof(bool))); // checkbox 
            
            //backlog datatable columns arrangement //
            dtbacklog.Columns["SELECTED"].SetOrdinal(0);
            dtbacklog.Columns["DEL_DATE"].SetOrdinal(1);
            dtbacklog.Columns["DEL_TIME"].SetOrdinal(2);
            dtbacklog.Columns["FACILITY"].SetOrdinal(3);
            dtbacklog.Columns["PARTNUMBER"].SetOrdinal(4);
            dtbacklog.Columns["BALANCE"].SetOrdinal(5);
            dtbacklog.Columns["REMARKS"].SetOrdinal(6);
            dtbacklog.Columns["PULL_TICKET_NUMBER"].SetOrdinal(7);
            dtbacklog.Columns["LINE"].SetOrdinal(8);
            dtbacklog.Columns["QTY_DEL"].SetOrdinal(9);
            dtbacklog.Columns["ORIGINAL_PULL"].SetOrdinal(10);
            dtbacklog.Columns["BACKLOGTYPE"].SetOrdinal(11);
            dtbacklog.Columns["HISTORY"].SetOrdinal(12);
            backlogdatagrid.DataSource = dtbacklog;
            //dtbacklog.Columns[1].DataType = Type.GetType("System.DateTime");

            
            
        }

        //Method to show data from database to datatable
        private void showrevision()
        {
            //revision
            adpt2 = new OracleDataAdapter("SELECT DEL_DATE, FACILITY, PARTNUMBER, PREVIOUS_PULL, NEW_PULL, REMARKS, PULL_TICKET_NUMBER, LINE FROM PULL_QUANTITY_REV", con); // query to select specified data from database (PULL_QUANTITY_REV).
            dt2 = new DataTable();
            adpt2.Fill(dt2);

            //dt2.Columns.Add(new DataColumn("SELECTED", typeof(bool)));
            dt2.Columns["DEL_DATE"].SetOrdinal(0);
            dt2.Columns["FACILITY"].SetOrdinal(1);
            dt2.Columns["PARTNUMBER"].SetOrdinal(2);
            dt2.Columns["PREVIOUS_PULL"].SetOrdinal(3);
            dt2.Columns["NEW_PULL"].SetOrdinal(4);
            dt2.Columns["REMARKS"].SetOrdinal(5);
            dt2.Columns["PULL_TICKET_NUMBER"].SetOrdinal(6);
            dt2.Columns["LINE"].SetOrdinal(7);

            revisiondatagrid.DataSource = dt2;
        }

        //manualDR datatable
        private void manualDRrecord()
        {
            adpt3 = new OracleDataAdapter("SELECT * FROM MANUAL_DR", con);
            dt3 = new DataTable();
            adpt3.Fill(dt3);

            dt3.Columns.Add(new DataColumn("SELECTED", typeof(bool)));
            dt3.Columns["SELECTED"].SetOrdinal(0);
            dt3.Columns["FACILITY"].SetOrdinal(1);
            dt3.Columns["PARTNUMBER"].SetOrdinal(2);
            dt3.Columns["PULL_QTY"].SetOrdinal(3);
            dt3.Columns["PULL_TICKET_NUMBER"].SetOrdinal(4);
            dt3.Columns["LINE"].SetOrdinal(5);
            dt3.Columns["REMARKS"].SetOrdinal(6);
            dt3.Columns["DATE_ADDED"].SetOrdinal(7);


            manualdrdatagrid.DataSource = dt3;
        }

        //button method to show pull ticket record datagrid
        private void btnrecords_Click(object sender, EventArgs e)
        {
            pnl_manualdr.Visible = false;//hide the manualDR datagrid
            pnl_rec.Visible = true;// show pull ticket record datagrid
            pnl_bcklog.Visible = false;//hide backlog datagrid
            pnl_revision.Visible = false;//hide revision datagrid
            showpullticketdata();
            partnumbox.Clear();
            facilbox.SelectedIndex = 0;
            totalData.Text = "TOTAL DATA: " + recorddatagrid.Rows.Count;
            recorddatagrid.ClearSelection();
        }

        //back button
        private void btnbk_Click(object sender, EventArgs e)
        {   
           Timer2.Start();//timer for transition

        }

        //Upload button
        private void btnupload_Click(object sender, EventArgs e)
        {
            Uploading upd = new Uploading();
            upd.ShowDialog();
        }

        //UL label button[updating ul label list]
        private void btnul_Click(object sender, EventArgs e)
        {

            //Current size of the forms 

            int formHeight = this.Height;
            int formWidth = this.Width;

            //This will open the ul_label form
            UL_label_form ul_form = new UL_label_form( );

            ul_form.Width = formWidth;
            ul_form.Height = formHeight;
        
            ul_form.StartPosition = FormStartPosition.Manual;
            ul_form.Load += delegate (object s2, EventArgs e2)
            {
                ul_form.Location = 
                new Point(this.Bounds.Location.X + this.Bounds.Width / 2 - ul_form.Width / 2,
                this.Bounds.Location.Y + this.Bounds.Height / 2 - ul_form.Height / 2);
            };
            ul_form.Show();

        }

        //backlog button to show backlog datagrid
        private void btnbcklog_Click(object sender, EventArgs e)
        {
            pnl_manualdr.Visible = false;//hide manual dr datagrid
            pnl_rec.Visible = false;//hide pullticket record datagrid
            pnl_bcklog.Visible = true; //show backlog datagrid
            pnl_revision.Visible = false;// hide revision datagrid
            backlogBox.SelectedIndex = 0;
            backlogTxt.Clear();
            showbacklogdata();
            clearCheckedBacklogData();
            totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;
        }

        //revision button to show revision datagrid
        private void btnrevision_Click(object sender, EventArgs e)
        {
            pnl_revision.Visible = true;// show revision datagrid
            pnl_bcklog.Visible = false;//hide backlog datagrid
            pnl_rec.Visible = false;//hide pullticket datagrid
            pnl_manualdr.Visible = false;//hide manualDR datagrid
            showrevision();
            TotalRev.Text = "TOTAL DATA: " + revisiondatagrid.Rows.Count;
            revisiondatagrid.ClearSelection();
        }

        //button to show manual dr datagrid
        private void btnmanualdr_Click(object sender, EventArgs e)
        {
            pnl_manualdr.Visible = true; // show manual dr datagrid
            pnl_rec.Visible = false; // hide pullticketrecord datagrid
            pnl_bcklog.Visible = false;// hide backlog datagrid
            pnl_revision.Visible = false;// hide revision datagrid
            manualDRrecord();
            manualdrdatagrid.ClearSelection();
        }

        //event for changing button on hover
        private void btnbk_MouseEnter(object sender, EventArgs e)
        {
            btnbk.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btnbk_MouseLeave(object sender, EventArgs e)
        {
            btnbk.ForeColor= Color.White;
        }

        //event for changing button on hover
        private void btnupload_MouseEnter(object sender, EventArgs e)
        {
            btnupload.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btnupload_MouseLeave(object sender, EventArgs e)
        {
            btnupload.ForeColor= Color.White;
        }

        //event for changing button on hover
        private void btnul_MouseEnter(object sender, EventArgs e)
        {
            btnul.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btnul_MouseLeave(object sender, EventArgs e)
        {
            btnul.ForeColor= Color.White;
        }

        //event for changing button on hover
        private void btnadd_MouseEnter(object sender, EventArgs e)
        {
            btnadd.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btnadd_MouseLeave(object sender, EventArgs e)
        {
            btnadd.ForeColor= Color.White;
        }

        //event for changing button on hover
        private void btnupdate_MouseEnter(object sender, EventArgs e)
        {
            btnupdate.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btnupdate_MouseLeave(object sender, EventArgs e)
        {
            btnupdate.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btndlt_MouseEnter(object sender, EventArgs e)
        {
            btndlt.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btndlt_MouseLeave(object sender, EventArgs e)
        {
            btndlt.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btnadd_dr_MouseEnter(object sender, EventArgs e)
        {
            btnadd_dr.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btnadd_dr_MouseLeave(object sender, EventArgs e)
        {
            btnadd_dr.ForeColor = Color.White;
        }

        //event for changing button on hover
        private void btnupdate_dr_MouseEnter(object sender, EventArgs e)
        {
            btnupdate_dr.ForeColor= Color.White;
        }

        //event for changing button on hover
        private void btnupdate_dr_MouseLeave(object sender, EventArgs e)
        {
            btnupdate_dr.ForeColor= Color.White;
        }

        //event for changing button on hover
        private void btndel_dr_MouseEnter(object sender, EventArgs e)
        {
            btndel_dr.ForeColor= Color.White;
        }

        //event for changing button on hover
        private void btndel_dr_MouseLeave(object sender, EventArgs e)
        {
            btndel_dr.ForeColor = Color.White;
        }

        //timer for carousel
        private void Timer1_Tick(object sender, EventArgs e)
        {
            if (Opacity == 1)
            {
                Timer1.Stop();
            }
            Opacity += .2;
        }

        //method to trigger the timer
        public void TriggerTimerTick()
        {
            Timer1_Tick(this, EventArgs.Empty);
        }

        //timer for transition
        private void Timer2_Tick(object sender, EventArgs e)
        {
            if (Opacity <= 0)
            {
                this.Close();
            }
            Opacity -= .2;
        }

        //facility combobox in pullticket record [conditions and loops when combobox value changed]
        private void facilbox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DataTable dttemp = new DataTable();
            DataView dv = new DataView(dt);
            //filtering data with Zero PULL_QTY//
            if (facilbox.Text == "ALL")
            {
                recorddatagrid.DataSource = dt;
                //check if partnumber in textbox is null or empty
                if (partnumbox.Text == "")
                {
                    if (recordfilterchk.Checked == true)
                    {
                        dt.DefaultView.RowFilter = "[PULL_QTY] > 0";
                    }
                    else
                    {
                        dt.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }
                }
                else
                {
                    dv.RowFilter = "PARTNUMBER LIKE '%" + partnumbox.Text.ToUpper() +"%'";
                    dttemp = dv.ToTable();
                    recorddatagrid.DataSource = dttemp;
                    //filtering data with Zero PULL_QTY//
                    if
                            (recordfilterchk.Checked == true)
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0";
                    }
                    else
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }
                }
            }
            else
            {
                switch (facilbox.Text)
                {
                    case "CAV3":
                        dv.RowFilter = "FACILITY = 'CAV3'";
                        break;
                    case "DKP":
                        dv.RowFilter = "FACILITY = 'DKP'";
                        break;
                    case "DANAM T":
                        dv.RowFilter = "FACILITY = 'DANAM T'";
                        break;
                    case "DANAM":
                        dv.RowFilter = "FACILITY = 'DANAM'";
                        break;
                    case "MACRO":
                        dv.RowFilter = "FACILITY = 'MACRO'";
                        break;
                    case "CAV2":
                        dv.RowFilter = "FACILITY = 'CAV2'";
                        break;
                    case "CAV5-CPPK":
                        dv.RowFilter = "FACILITY = 'CAV5-CPPK'";
                        break;
                    case "CLP/CAV5":
                        dv.RowFilter = "FACILITY = 'CLP/CAV5'";
                        break;
                    case "CLP":
                        dv.RowFilter = "FACILITY = 'CLP'";
                        break;

                }

                dttemp = dv.ToTable();
                if (recordfilterchk.Checked)
                {
                    dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0";
                }
                else
                {
                    dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                }

                if (partnumbox.Text == "")
                {
                dv.RowFilter = "FACILITY = '" +facilbox.Text+"'";
                dttemp = dv.ToTable();
                recorddatagrid.DataSource = dttemp;
                    //filtering data with Zero PULL_QTY//
                    if (recordfilterchk.Checked == true)
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0";
                    }
                    else
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }
                }
                else
                {
                    dv.RowFilter = "FACILITY = '"+facilbox.Text+"' and PARTNUMBER LIKE '%" +partnumbox.Text.ToUpper()+ "%'";
                    dttemp = dv.ToTable();
                    recorddatagrid.DataSource = dttemp;
                    //filtering data with Zero PULL_QTY//
                    if (recordfilterchk.Checked == true)
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0";
                    }
                    else
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }

                }
            }
        }
        //method for partnumber search textbox [filtering data by partnumber]
        private void partnumbox_TextChanged(object sender, EventArgs e)
        {
            DataTable dttemp = new DataTable();
            DataView dv = new DataView(dt);

            if (facilbox.Text == "ALL")
            {
                recorddatagrid.DataSource = dt;
                if (partnumbox.Text == "")
                {
                    //filtering data with Zero PULL_QTY//
                    if (recordfilterchk.Checked == true)
                    {
                        dt.DefaultView.RowFilter = "[PULL_QTY] > 0";
                    }
                    else
                    {
                        dt.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }
                }
                else
                {
                    dv.RowFilter = "PARTNUMBER LIKE '%" + partnumbox.Text.ToUpper() + "%'";
                    dttemp = dv.ToTable();
                    recorddatagrid.DataSource = dttemp;
                    //filtering data with Zero PULL_QTY//
                    if (recordfilterchk.Checked == true)
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0";
                    }
                    else
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }
                }
            }
            else
            {
                switch (facilbox.Text)
                {
                    case "CAV3":
                        dv.RowFilter = "FACILITY = 'CAV3'";
                        break;
                    case "DKP":
                        dv.RowFilter = "FACILITY = 'DKP'";
                        break;
                    case "DANAM T":
                        dv.RowFilter = "FACILITY = 'DANAM T'";
                        break;
                    case "DANAM":
                        dv.RowFilter = "FACILITY = 'DANAM'";
                        break;
                    case "MACRO":
                        dv.RowFilter = "FACILITY = 'MACRO'";
                        break;
                    case "CAV2":
                        dv.RowFilter = "FACILITY = 'CAV2'";
                        break;
                    case "CAV5-CPPK":
                        dv.RowFilter = "FACILITY = 'CAV5-CPPK'";
                        break;
                    case "CLP/CAV5":
                        dv.RowFilter = "FACILITY = 'CLP/CAV5'";
                        break;
                    case "CLP":
                        dv.RowFilter = "FACILITY = 'CLP'";
                        break;
                }

                dttemp = dv.ToTable();
                if (recordfilterchk.Checked)
                {
                    dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0";
                }
                else
                {
                    dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                }

                if (partnumbox.Text == "")
                {
                    dv.RowFilter = "FACILITY = '" + facilbox.Text + "'";
                    dttemp = dv.ToTable();
                    recorddatagrid.DataSource = dttemp;
                    //filtering data with Zero PULL_QTY//
                    if (recordfilterchk.Checked == true)
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0";
                    }
                    else
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }
                }
                else
                {
                    dv.RowFilter = "FACILITY = '" + facilbox.Text + "' AND PARTNUMBER LIKE '%" + partnumbox.Text.ToUpper() + "%'";
                    dttemp = dv.ToTable();
                    recorddatagrid.DataSource = dttemp;
                    //filtering data with Zero PULL_QTY//
                    if (recordfilterchk.Checked == true)
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0";
                    }
                    else
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }
                }
            }
        }
        //searchby combobox in backlog
        private void backlogBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DataTable dttemp = new DataTable();
            DataView dv = new DataView(dtbacklog);
            if(backlogBox.Text == "ALL")
            {
                backlogTxt.Enabled = false;
                backlogdatagrid.DataSource = dtbacklog;           
            }
            else if (backlogBox.Text == "AS PER ADVISE")
            {
                backlogTxt.Enabled = false;
                dv.RowFilter = "BACKLOGTYPE = '" + backlogBox.Text + "'";
                dttemp = dv.ToTable();
                backlogdatagrid.DataSource = dttemp;
            }else if(backlogBox.Text == "FACILITY" || backlogBox.Text == "PULLTICKET NO." || backlogBox.Text == "PARTNUMBER")
            {
                backlogTxt.Enabled = true;
                dv.RowFilter = "FACILITY LIKE '%" +backlogTxt.Text.ToUpper() + "%'";
                dttemp = dv.ToTable();
                backlogdatagrid.DataSource = dttemp;
            }
            
        }
        //searchby textbox in backlog
        private void backlogTxt_TextChanged(object sender, EventArgs e)
        {
            DataTable dttemp = new DataTable();
            DataView dv = new DataView(dtbacklog);
            if (backlogBox.Text == "FACILITY")
            {
                dv.RowFilter = "FACILITY LIKE '%" + backlogTxt.Text.ToUpper() + "%'";
                dttemp = dv.ToTable();
                backlogdatagrid.DataSource = dttemp;
                //Select/deselect data//
                if (selectallchk.Checked)
                {
                    selectallchk.Text = "Deselect All";
                }
                else
                {
                    selectallchk.Text = "Select All";
                }

                foreach (DataGridViewRow row in backlogdatagrid.Rows)
                {
                    DataGridViewCheckBoxCell chk = row.Cells[0] as DataGridViewCheckBoxCell;
                    if (chk != null)
                    {
                        chk.Value = selectallchk.Checked;
                    }
                    else
                    {
                        chk.Value = !selectallchk.Checked;
                    }
                }
            } else if (backlogBox.Text == "PULLTICKET NO.")
            {
                dv.RowFilter = "PULL_TICKET_NUMBER LIKE '%" + backlogTxt.Text + "%'";
                dttemp = dv.ToTable();
                backlogdatagrid.DataSource = dttemp;
                //Select/deselect data//
                if (selectallchk.Checked)
                {
                    selectallchk.Text = "Deselect All";
                }
                else
                {
                    selectallchk.Text = "Select All";
                }

                foreach (DataGridViewRow row in backlogdatagrid.Rows)
                {
                    DataGridViewCheckBoxCell chk = row.Cells[0] as DataGridViewCheckBoxCell;
                    if (chk != null)
                    {
                        chk.Value = selectallchk.Checked;
                    }
                    else
                    {
                        chk.Value = !selectallchk.Checked;
                    }
                }
            }
            else if (backlogBox.Text == "ALL")
            {
                backlogdatagrid.DataSource = dtbacklog;
                //Select/deselect data//
                if (selectallchk.Checked)
                {
                    selectallchk.Text = "Deselect All";
                }
                else
                {
                    selectallchk.Text = "Select All";
                }

                foreach (DataGridViewRow row in backlogdatagrid.Rows)
                {
                    DataGridViewCheckBoxCell chk = row.Cells[0] as DataGridViewCheckBoxCell;
                    if (chk != null)
                    {
                        chk.Value = selectallchk.Checked;
                    }
                    else
                    {
                        chk.Value = !selectallchk.Checked;
                    }
                }
            } else if (backlogBox.Text == "AS PER ADVISE")
            {
                dv.RowFilter = "BACKLOGTYPE = '" + backlogBox.Text + "'";
                dttemp = dv.ToTable();
                backlogdatagrid.DataSource = dttemp;
                //Select/deselect data//
                if (selectallchk.Checked)
                {
                    selectallchk.Text = "Deselect All";
                }
                else
                {
                    selectallchk.Text = "Select All";
                }

                foreach (DataGridViewRow row in backlogdatagrid.Rows)
                {
                    DataGridViewCheckBoxCell chk = row.Cells[0] as DataGridViewCheckBoxCell;
                    if (chk != null)
                    {
                        chk.Value = selectallchk.Checked;
                    }
                    else
                    {
                        chk.Value = !selectallchk.Checked;
                    }
                }
            }
            else if(backlogBox.Text == "PARTNUMBER")
            {
                dv.RowFilter = "PARTNUMBER LIKE '%" + backlogTxt.Text + "%'";
                dttemp = dv.ToTable();
                backlogdatagrid.DataSource= dttemp;
                //Select/deselect data//
                if (selectallchk.Checked)
                {
                    selectallchk.Text = "Deselect All";
                }
                else
                {
                    selectallchk.Text = "Select All";
                }

                foreach (DataGridViewRow row in backlogdatagrid.Rows)
                {
                    DataGridViewCheckBoxCell chk = row.Cells[0] as DataGridViewCheckBoxCell;
                    if (chk != null)
                    {
                        chk.Value = selectallchk.Checked;
                    }
                    else
                    {
                        chk.Value = !selectallchk.Checked;
                    }
                }


            }
            else if(backlogTxt.Text == "")
            {
                backlogdatagrid.DataSource = dtbacklog;
                //Select/deselect data//
                if (selectallchk.Checked)
                {
                    selectallchk.Text = "Deselect All";
                }
                else
                {
                    selectallchk.Text = "Select All";
                }

                foreach (DataGridViewRow row in backlogdatagrid.Rows)
                {
                    DataGridViewCheckBoxCell chk = row.Cells[0] as DataGridViewCheckBoxCell;
                    if (chk != null)
                    {
                        chk.Value = selectallchk.Checked;
                    }
                    else
                    {
                        chk.Value = !selectallchk.Checked;
                    }
                }
            }
        }
        //add button in backlog
        private void btnadd_Click(object sender, EventArgs e)
        {
            amCheckbox.Checked = true;
            pnlDeltime.SelectedIndex = 0;

            backlogpanel_btnadd.Text = "ADD";
            // This will check if the panel as per advise width is greater than 304 which is means the pnl is open
            // if the panel as per advise is open when the btn add is click the panel as per advise width will turn into 0 or close the as per advise panel
            if (pnl_asperadvice.Width < 276)
            {
                pnl_asperadvice.Width = 276;
                backlogdatagrid.Location = new Point(280, 86);
                searchby.Visible = false;
                backlogBox.Visible = false;
                backlogTxt.Visible = false;
                selectallchk.Visible = false;
                btnadd.Enabled = false;
                btnupdate.Enabled = false;
                btndlt.Enabled = false;
                pnlDeltime.Enabled = true;
                amCheckbox.Enabled = true;
                pmCheckbox.Enabled = true;
                pnlfacility.Enabled = true;
                asppartnumber.Enabled = true;
                pnlpullticketnumber.Enabled = true;
                pnlline.Enabled = true;
                pnldelivered.Enabled = true;
                pnlDeldate.Enabled = true;
                pnlpullqty.Enabled = true;
            }
            else
            {
                //pnl_asperadvice.Width = 0;
                
                backlogdatagrid.Location = new Point(0, 86);
                searchby.Visible = true;
                backlogBox.Visible = true;
                backlogTxt.Visible = true;
                selectallchk.Visible = true;

            }
            pnl_delivery_time_date.Visible = true;


            //These will move the as per advice sub panel 2 content (  )
            asperadvice_sub_panel2.Location = new Point(13, 114);
       

        }

        private void btndlt_Click(object sender, EventArgs e)
        {
            if(checkedBacklogData.Count <= 0)
            {
                MessageBox.Show("No Data selected!", "Warning!");
            }
            else if(checkedBacklogData.Count > 1)
            {
                DialogResult dr0 = MessageBox.Show("You are deleting multiple items?", "Do you want to proceed?", MessageBoxButtons.YesNo);
                switch (dr0)
                {
                    case DialogResult.Yes:

                        foreach (var item in checkedBacklogData)
                        {
                            string delall = "delete from BACKLOG where PULL_TICKET_NUMBER = '" + item.PullTicketNumber + "' and LINE = '" + item.Line + "'";
                            var pulldelcmd = new OracleCommand(delall, con);
                            pulldelcmd.ExecuteNonQuery();
                        }
                        MessageBox.Show("Files Deleted!");

                        showbacklogdata();
                        clearCheckedBacklogData();
                        totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;

                        break;
                        
                    case DialogResult.No:
                        showbacklogdata();
                        clearCheckedBacklogData();
                        totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;
                        break;
                }
                
            }else if(checkedBacklogData.Count == 1)
            {
                pnlpullticketnumber.Text = checkedBacklogData[0].PullTicketNumber;
                pnlline.Text = checkedBacklogData[0].Line;

                DialogResult dr = MessageBox.Show("Are you sure you want to delete this data?", "Are you sure?", MessageBoxButtons.YesNo);
                switch(dr)
                {
                    case DialogResult.Yes:
                       string puldel = "delete from BACKLOG where PULL_TICKET_NUMBER = '" + checkedBacklogData[0].PullTicketNumber + "' and LINE = '" + checkedBacklogData[0].Line + "'";
                        var pulldelcmd = new OracleCommand(puldel, con);
                        pulldelcmd.ExecuteNonQuery();
                        MessageBox.Show("Delete successsuly");

                        clearCheckedBacklogData();
                        showbacklogdata();
                        totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;
                        break;
                    case DialogResult.No:
                        showbacklogdata();
                        clearCheckedBacklogData();
                        totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;
                        break;
                }
            }
        }
        //update button in backlog
        private void btnupdate_Click(object sender, EventArgs e)
        {
            amCheckbox.Checked = true;
            backlogpanel_btnadd.Text = "UPDATE";
            int rowindex = backlogdatagrid.CurrentCell.RowIndex;
            int columnindex = backlogdatagrid.CurrentCell.ColumnIndex;

            //These will fill the data from the list of checked data in the backlog
            if (checkedBacklogData.Count == 1)
            {
                // This will check if the panel as per advise width is greater than 304 which is means the pnl is open
                // if the panel as per advise is open when the btn add is click the panel as per advise width will turn into 0 or close the as per advise panel
                if (pnl_asperadvice.Width < 276)
                {
                    pnl_asperadvice.Width = 276;
                    backlogdatagrid.Location = new Point(280, 86);
                    searchby.Visible = false;
                    backlogBox.Visible = false;
                    backlogTxt.Visible = false;
                    selectallchk.Visible = false;
                    btnupdate.Enabled = false;
                    btnadd.Enabled = false;
                    btndlt.Enabled = false;
                    pnlfacility.Enabled = true;
                    asppartnumber.Enabled = false;
                    pnlpullticketnumber.Enabled = false;
                    pnlline.Enabled = false;
                }
                else
                {
                    //pnl_asperadvice.Width = 0;
                    
                    backlogdatagrid.Location = new Point(0, 86);
                    searchby.Visible = true;
                    backlogBox.Visible = true;
                    backlogTxt.Visible = true;
                    selectallchk.Visible = true;

                }
                pnl_delivery_time_date.Visible = false;
                asperadvice_sub_panel2.Location = new Point(13, 52);
                asppartnumber.Text = checkedBacklogData[0].Partnumber;
                pnlfacility.Text = checkedBacklogData[0].Facility;
                pnlpullqty.Text = checkedBacklogData[0].Original_Pull;
                pnldelivered.Text = checkedBacklogData[0].Qty_Del;
                pnlpullticketnumber.Text = checkedBacklogData[0].PullTicketNumber;
                pnlline.Text = checkedBacklogData[0].Line;
                pnlpullqty.Enabled = false;
                if (checkedBacklogData[0].BackLogType == "")
                {
                    pnlBacklogtype.Checked = false;
                }
                else
                {
                    pnlBacklogtype.Checked = true;
                }
            } else if (checkedBacklogData.Count > 1) {
                //MessageBox.Show("Too many selected data.", "Warning");

                DialogResult dr0 = MessageBox.Show("Updating multiple items", "Do you want to proceed?", MessageBoxButtons.YesNo);
                switch (dr0)
                {
                    case DialogResult.Yes:
                        pnl_asperadvice.Width = 276;
                        backlogdatagrid.Location = new Point(280, 86);
                        searchby.Visible = false;
                        backlogBox.Visible = false;
                        backlogTxt.Visible = false;
                        selectallchk.Visible = false;
                        btnupdate.Enabled = false;
                        btnadd.Enabled = false;
                        btndlt.Enabled = false;
                        pnlfacility.Enabled = false;
                        asppartnumber.Enabled = false;
                        pnlpullticketnumber.Enabled = false;
                        pnlline.Enabled = false;    
                        pnldelivered.Enabled = false;
                        pnlDeldate.Enabled = false;
                        pnlpullqty.Enabled = false;
                        break;
                    case DialogResult.No:
                        showbacklogdata();
                        clearCheckedBacklogData();
                        totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;
                        break;

                }
            }
            else if(checkedBacklogData.Count < 1) {
                MessageBox.Show("Please select one backlogdata.", "Warning");

            }else if(selectallchk.Checked == true)
            {
                DialogResult dr0 = MessageBox.Show("Updating multiple items", "Do you want to proceed?", MessageBoxButtons.YesNo);
                switch (dr0)
                {
                    case DialogResult.Yes:
                        pnl_asperadvice.Width = 276;
                        backlogdatagrid.Location = new Point(280, 86);
                        searchby.Visible = false;
                        backlogBox.Visible = false;
                        backlogTxt.Visible = false;
                        selectallchk.Visible = false;
                        btnupdate.Enabled = false;
                        btnadd.Enabled = false;
                        btndlt.Enabled = false;
                        pnlfacility.Enabled = false;
                        asppartnumber.Enabled = false;
                        pnlpullticketnumber.Enabled = false;
                        pnlline.Enabled = false;
                        pnldelivered.Enabled = false;
                        pnlDeldate.Enabled = false;
                        pnlpullqty.Enabled = false;
                        break;
                    case DialogResult.No:
                        showbacklogdata();
                        clearCheckedBacklogData();
                        totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;
                        break;

                }
            }


        }
        //Back button of asperadvice panel
        private void asperadvice_back_btn_Click(object sender, EventArgs e)
        {
            //These will check if the screen is fullscreen
            //If the screen is not full screen the size of the backlogdatagrid will change into 940 width and 396 height
            pnl_asperadvice.Width = 0;
            backlogdatagrid.Location = new Point(0, 86);
            searchby.Visible = true;
            backlogBox.Visible = true;
            backlogTxt.Visible = true;
            selectallchk.Visible = true;

            pnlBacklogtype.Checked = false;
            pnlDeltime.Text = "";
            pnlfacility.Text = "";
            asppartnumber.Text = "";
            pnlpullqty.Text = "";
            pnldelivered.Text = "";
            pnlpullticketnumber.Text = "";
            pnlline.Text = "";
            amCheckbox.Checked = false;
            pmCheckbox.Checked = false;

            btnadd.Enabled = true;
            btnupdate.Enabled = true;
            btndlt.Enabled = true;
            showbacklogdata();
            clearCheckedBacklogData();
            totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;

            foreach (DataGridViewRow row in backlogdatagrid.Rows)
            {
                DataGridViewCheckBoxCell chk = row.Cells[0] as DataGridViewCheckBoxCell;
                //chk.Value = DBNull.Value;
                checkedBacklogData.Count.Equals(0);
            }
            



        }
        //am Check box if the am is check the pm box will unchecked
        private void amCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            pmCheckbox.Checked = false;
        }
        //pm Check box if the am is check the am box will unchecked
        private void pmCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            amCheckbox.Checked = false;
        }


        //Back button of manual dr pane panel
        private void btnadd_dr_Click(object sender, EventArgs e)
        {
            mdrPullTktTxt.Text = string.Empty;
            mdrFacilTxt.Text = string.Empty;
            mdrQuantTxt.Text = string.Empty;
            mdrPartnumTxt.Text = string.Empty;
            mdrLineTxt.Text = string.Empty;
            mdrRemTxt.Text = string.Empty;
            // This will check if the panel of pnl_manualdr_add width is greater than 276 which is means the pnl is open
            // if the pnl_manualdr_add width is open when the btn add is click the pnl_manualdr_add width width will turn into 0 or close the as pnl_manualdr_add width
            pnl_manualdr_add.Width = 276;
            manualdrdatagrid.Location = new Point(280, 66);
            btnadd_dr.Enabled = false;
            btnupdate_dr.Enabled = false;
            btndel_dr.Enabled = false;

            btnAddMdr.Text = "ADD";

        }
        //pnl_manualdr_add back 
        private void manualdr_add_edit_back_btn_Click(object sender, EventArgs e)
        {
            pnl_manualdr_add.Width = 0;
            manualdrdatagrid.Location = new Point(2, 66);
            btnadd_dr.Enabled = true;
            btnupdate_dr.Enabled = true;
            btndel_dr.Enabled = true;
            mdrFacilTxt.Text = "";
            mdrPartnumTxt.Text = "";
            mdrQuantTxt.Text = "";
            mdrPullTktTxt.Text = "";
            mdrLineTxt.Text = "";
            mdrRemTxt.Text = "";
        }
        //checking backlogdata
        private void clearCheckedBacklogData() { 
            backlogdatagrid.ClearSelection();
            checkedBacklogData.Clear();
            
        }
        private void clearCheckedManualDrData()
        {
            manualdrdatagrid.ClearSelection();
            checkedmanualDrData.Clear();
        }
        //This will tell if the program will run for the first time
        private bool FirstRun = true;
        private bool FirstRundr = true;
        //This will hold the value of the row that current selected
        private bool isRowChecked = false;

        private bool isRowCheckedDr = false;
        //backlogdatagrid rowstate method
        private void backlogdatagrid_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            //these will disregard other event if it was not selected event
            if (e.StateChanged != DataGridViewElementStates.Selected) return;

        }
        //This is when the cell of the backlog datagridview was clicked
        private void backlogdatagrid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            
            if (FirstRun)
            {
                clearCheckedBacklogData();
                FirstRun = false;
            }

            if (backlogdatagrid.SelectedCells.Count > 0)
            {

                int selectedrowindex = backlogdatagrid.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = backlogdatagrid.Rows[selectedrowindex];
                string cellValueSelected = Convert.ToString(selectedRow.Cells["Selected"].Value);


                if (cellValueSelected.Equals("True"))
                {
                    isRowChecked = true;

                }
                else
                {
                    isRowChecked = false;
                }

                String delDate = selectedRow.Cells["DEL_DATE"].Value.ToString();
                String delTime = selectedRow.Cells["DEL_TIME"].Value.ToString();
                String facility = selectedRow.Cells["FACILITY"].Value.ToString();
                String partNumber = selectedRow.Cells["PARTNUMBER"].Value.ToString();
                String balance = selectedRow.Cells["BALANCE"].Value.ToString();
                String remarks = selectedRow.Cells["REMARKS"].Value.ToString();
                String pullTicketNumber = selectedRow.Cells["PULL_TICKET_NUMBER"].Value.ToString();
                String line = selectedRow.Cells["LINE"].Value.ToString();
                String qty_del = selectedRow.Cells["QTY_DEL"].Value.ToString();
                String original_pull = selectedRow.Cells["ORIGINAL_PULL"].Value.ToString();
                String backlogtype = selectedRow.Cells["BACKLOGTYPE"].Value.ToString(); 
                String history = selectedRow.Cells["HISTORY"].Value.ToString();



                if (isRowChecked == false) //if rowCheck is false
                {

                    //checkedBacklogData.RemoveAll(r => r.isChecked == false);
                    //This will check if the backload datamodel if it was unchecked

                    //This will check if the current selected row have
                    //the same del_date, partnumber, balance and line if 
                    //they are have the same value it will not be added.
                    
                    int checkedCountValue = 0;
                    foreach (BacklogDataModel data in checkedBacklogData)
                    {
                      var matches = checkedBacklogData.Where(p => p.Del_Date == delDate &&
                        p.Partnumber == partNumber && p.Balance == balance && p.Line == line);

                        checkedCountValue = matches.Count();
                    }

                    if (checkedCountValue == 0) {

                        checkedBacklogData.Add(new BacklogDataModel(!isRowChecked, delDate, delTime, facility, partNumber, balance, remarks,
                            pullTicketNumber, line, qty_del, original_pull, backlogtype, history));
                        selectedRow.Cells["SELECTED"].Value = true;
                    }

                  

                
                }
                else {

                    int checkedCountValue = 0;
                    foreach (BacklogDataModel data in checkedBacklogData)
                    {
                        var matches = checkedBacklogData.Where(
                              p => p.Del_Date == delDate &&
                        p.Partnumber == partNumber && p.Balance == balance && p.Line == line);
                        
                        checkedCountValue = matches.Count();
                    }

                    if (checkedCountValue > 0)
                    {

                        checkedBacklogData.Remove(checkedBacklogData.Find(p => p.Del_Date == delDate &&
                        p.Partnumber == partNumber && p.Balance == balance && p.Line == line));
                        selectedRow.Cells["SELECTED"].Value = false;
                    }

                }
            }
        }
        private void btnupdate_dr_Click(object sender, EventArgs e)
        {
            
            if(manualdrdatagrid.Rows.Count > 0)
            {
                int rowindexDr = manualdrdatagrid.CurrentCell.RowIndex;
                int collumnindexDr = manualdrdatagrid.CurrentCell.ColumnIndex;


                if (checkedmanualDrData.Count == 1)
                {
                    pnl_manualdr_add.Width = 276;
                    manualdrdatagrid.Location = new Point(280, 66);
                    btnadd_dr.Enabled = false;
                    btnupdate_dr.Enabled = false;
                    btndel_dr.Enabled = false;

                    btnAddMdr.Text = "UPDATE";


                    mdrFacilTxt.Text = checkedmanualDrData[0].Facility;
                    mdrPartnumTxt.Text = checkedmanualDrData[0].Partnumber;
                    mdrQuantTxt.Text = checkedmanualDrData[0].Pull_Qty;
                    mdrPullTktTxt.Text = checkedmanualDrData[0].Pull_Ticker_Number;
                    mdrLineTxt.Text = checkedmanualDrData[0].Line;
                    mdrRemTxt.Text = checkedmanualDrData[0].Remarks;

                    mdrPullTktTxt.Enabled = false;
                    mdrLineTxt.Enabled = false;
                }
                else if (checkedmanualDrData.Count > 1)
                {
                    MessageBox.Show("Too many selected data.", "Warning");
                }
                else
                {
                    MessageBox.Show("Please select only one backlogdata.", "Warning");

                }
            }
            else
            {
                MessageBox.Show("No Data Entry", "Error!");
            }
        }

        private void manualdrdatagrid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (FirstRundr)
            {
                clearCheckedManualDrData();
                FirstRundr = false;
            }

            if (manualdrdatagrid.SelectedCells.Count > 0)
            {

                int selectedrowindexDr = manualdrdatagrid.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRowDr = manualdrdatagrid.Rows[selectedrowindexDr];
                string cellValueSelectedDr = Convert.ToString(selectedRowDr.Cells["Selected"].Value);

                if (cellValueSelectedDr.Equals("True"))
                {
                    isRowCheckedDr = true;

                }
                else
                {
                    isRowCheckedDr = false;
                }

                String facility = selectedRowDr.Cells["FACILITY"].Value.ToString();
                String partNumber = selectedRowDr.Cells["PARTNUMBER"].Value.ToString();
                String pullqty = selectedRowDr.Cells["PULL_QTY"].Value.ToString();
                String pullTicketNumber = selectedRowDr.Cells["PULL_TICKET_NUMBER"].Value.ToString();
                String line = selectedRowDr.Cells["LINE"].Value.ToString();
                String remarks = selectedRowDr.Cells["REMARKS"].Value.ToString();
                String dateadded = selectedRowDr.Cells["DATE_ADDED"].Value.ToString();



                if (isRowCheckedDr == false) //if rowCheck is false
                {

                    //checkedBacklogData.RemoveAll(r => r.isChecked == false);
                    //This will check if the bac	kload datamodel if it was unchecked

                    //This will check if the current selected row have
                    //the same del_date, partnumber, balance and line if 
                    //they are have the same value it will not be added.

                    int checkedCountValueDr = 0;
                    foreach (ManualDRDataModel data in checkedmanualDrData)
                    {
                        var matches = checkedmanualDrData.Where(p => p.Facility == facility &&
                          p.Partnumber == partNumber && p.Pull_Qty == pullqty && p.Pull_Ticker_Number == pullTicketNumber && p.Line == line && p.Remarks == remarks && p.Date_Added == dateadded);

                        checkedCountValueDr = matches.Count();
                    }

                    if (checkedCountValueDr == 0)
                    {

                        checkedmanualDrData.Add(new ManualDRDataModel(!isRowChecked, facility, partNumber, pullqty, pullTicketNumber, line, remarks, dateadded));
                        selectedRowDr.Cells["SELECTED"].Value = true;
                    }




                }
                else
                {

                    int checkedCountValueDr = 0;
                    foreach (ManualDRDataModel data in checkedmanualDrData)
                    {
                        var matches = checkedmanualDrData.Where(
                              p => p.Facility == facility &&
                        p.Partnumber == partNumber && p.Pull_Qty == pullqty && p.Pull_Ticker_Number == pullTicketNumber && p.Line == line && p.Remarks == remarks && p.Date_Added == dateadded);

                        checkedCountValueDr = matches.Count();
                    }

                    if (checkedCountValueDr > 0)
                    {

                        checkedmanualDrData.Remove(checkedmanualDrData.Find(p => p.Facility == facility &&
                        p.Partnumber == partNumber && p.Pull_Qty == pullqty && p.Pull_Ticker_Number == pullTicketNumber && p.Line == line && p.Remarks == remarks && p.Date_Added == dateadded));
                        selectedRowDr.Cells["SELECTED"].Value = false;
                    }

                }
            }

        }

        private void manualdrdatagrid_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            if (e.StateChanged != DataGridViewElementStates.Selected) return;
        }

        private void btnAddMdr_Click_1(object sender, EventArgs e)
        {
            if(btnAddMdr.Text == "ADD")
            {
                var dupcheck = new OracleDataAdapter("Select * from MANUAL_DR where PULL_TICKET_NUMBER = '" +mdrPullTktTxt.Text+ "' or LINE = '" +mdrLineTxt.Text+ "'", con);
                DataTable dubpchecktbl = new DataTable();
                string dateAdded = DateTime.Now.ToString("M/d/yyy");
                string myrem = "";

                if (mdrRemTxt.Text.Equals(string.Empty))
                {
                    myrem = "NO REMARKS";
                }
                else
                {
                    myrem = mdrRemTxt.Text;
                }

                dupcheck.Fill(dubpchecktbl);
                if(mdrFacilTxt.Text == "" || mdrPartnumTxt.Text == "" || mdrQuantTxt.Text == "" ||  mdrPullTktTxt.Text == "" || mdrLineTxt.Text == "")
                {
                    MessageBox.Show("Fill out the required field!", "WARNING!");
                }
                else if (dubpchecktbl.Rows.Count > 0)
                {
                    MessageBox.Show("Part number or Line Already Exist!, Try another One(DJ KHALID!!!!)", "Warning!");
                }
                else
                {
                    string insrec = "insert into MANUAL_DR (FACILITY, PARTNUMBER, PULL_QTY, PULL_TICKET_NUMBER, LINE, REMARKS, DATE_ADDED) values ('"+mdrFacilTxt.Text.ToString().ToUpper() + "','"+ mdrPartnumTxt.Text.ToString().ToUpper() + "','" +mdrQuantTxt.Text.ToString().ToUpper() + "','"+mdrPullTktTxt.Text.ToString().ToUpper() + "','" +mdrLineTxt.Text.ToString().ToUpper() + "','" + myrem.ToString().ToUpper() + "','" +dateAdded+"')";
                    OracleCommand insrecadd = new OracleCommand(insrec,con);
                    insrecadd.ExecuteNonQuery();


                    MessageBox.Show("Added Succefully","Success!");
                    manualDRrecord();
                    clearCheckedManualDrData();
                    manualdrdatagrid.Location = new Point(2, 66);
                    btnadd_dr.Enabled = true;
                    btnupdate_dr.Enabled=true;
                    btndel_dr.Enabled = true;
                    pnl_manualdr_add.Width = 0;

                    mdrFacilTxt.Text = "";
                    mdrPartnumTxt.Text = "";
                    mdrQuantTxt.Text = "";
                    mdrPullTktTxt.Text = "";
                    mdrLineTxt.Text = "";
                    mdrRemTxt.Text = "";
                }



            }
            else if (btnAddMdr.Text == "UPDATE")
            {
                
                if (mdrFacilTxt.Text == "" || mdrPartnumTxt.Text == "" || mdrQuantTxt.Text == "")
                {
                    MessageBox.Show("Fill out all the fields!","Error!");
                }
                else
                {
                    var dupcheck = new OracleDataAdapter("select * from MANUAL_DR where PULL_TICKET_NUMBER = '"+ mdrPullTktTxt.Text.ToString().ToUpper() + "' and LINE = '"+ mdrLineTxt.Text.ToString().ToUpper() +"'", con);
                    DataTable dupchecktbl = new DataTable();
                    dupcheck.Fill(dupchecktbl);
                    dupcheck.Dispose();

                    string insRec = "UPDATE MANUAL_DR SET FACILITY = '" + mdrFacilTxt.Text+ "', PARTNUMBER = '" +mdrPartnumTxt.Text + "', PULL_QTY = '" +mdrQuantTxt.Text + "', REMARKS = '" +mdrRemTxt.Text + "' WHERE PULL_TICKET_NUMBER ='" +mdrPullTktTxt.Text + "' and LINE = '" +mdrLineTxt.Text + "'";
                    var insRecup = new OracleCommand(insRec,con);
                    insRecup.ExecuteNonQuery();

                    manualDRrecord();
                    MessageBox.Show("Updated Successfully !");
                    clearCheckedManualDrData();
                    pnl_manualdr_add.Width = 0;
                    manualdrdatagrid.Location = new Point(2, 66);
                    btnadd_dr.Enabled = true;
                    btnupdate_dr.Enabled = true;
                    btndel_dr.Enabled = true;
                }


            }
        }

        private void btndel_dr_Click(object sender, EventArgs e)
        {
            if(checkedmanualDrData.Count <= 0)
            {
                MessageBox.Show("No data selected!", "Warning!");
            }
            else if(checkedmanualDrData.Count > 1)
            {
                MessageBox.Show("Too many items selected. . . .");
            }else if (checkedmanualDrData.Count == 1)
            {
                mdrPullTktTxt.Text = checkedmanualDrData[0].Pull_Ticker_Number;
                mdrLineTxt.Text = checkedmanualDrData[0].Line;

                DialogResult Dr = MessageBox.Show("Are you sure you want to delete this data?", "Are you sure?", MessageBoxButtons.YesNo);
                switch(Dr)
                {
                    case DialogResult.Yes:
                        string mdrdelete = "delete from MANUAL_DR where PULL_TICKET_NUMBER  = '"+ checkedmanualDrData[0].Pull_Ticker_Number + "' and LINE = '" + checkedmanualDrData[0].Line + "'";
                        var mdrdeletecmd = new OracleCommand(mdrdelete, con);
                        mdrdeletecmd.ExecuteNonQuery();

                        MessageBox.Show("Deleted Successfully");
                        clearCheckedManualDrData();
                        manualDRrecord();
                        break;
                    case DialogResult.No:
                        //MessageBox.Show("edi don't!");
                        break;
                }
            }
        }

        private void btndownloadptk_Click(object sender, EventArgs e)
        {
            if (recorddatagrid.Columns.Count == 0 || recorddatagrid.Rows.Count == 0)
            {
                MessageBox.Show("There is no data to export!");
                return;
            }

            ExportRec(); // calls the method for creating excel files
        }

        private void ExportRec()
        {

            //prompts the user to select the location where the file will be saved
            string dateTimeFile = DateTime.Now.ToString("yyyy_MM_dd (hhmmtt)");

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "Excel Workbook|*.xls,*.xlsx";
            saveFileDialog1.Title = "Save Excel File";
            saveFileDialog1.FileName = $"PULLTICKET_RECORD_{dateTimeFile}.xls";
            saveFileDialog1.InitialDirectory = @"C:\";

            string strFileName;
            bool blnFileOpen;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(saveFileDialog1.FileName))
                {
                    try
                    {
                        using (FileStream fs = (FileStream)saveFileDialog1.OpenFile())
                        {
                            // No need to close the file stream as it will be closed by the using statement
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("File Not Found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    strFileName = saveFileDialog1.FileName;
                    blnFileOpen = false;

                    try
                    {
                        using (FileStream fileTemp = File.OpenWrite(strFileName))
                        {
                            // No need to close the file stream as it will be closed by the using statement
                        }
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
            }
            else
            {
                return;
            }

            // Show the form that tells the status of the application
            Form2 f2 = new Form2();
            f2.Show();
            f2.label2.Text = "Creating Excel File...";
            f2.Refresh();

            // Get the current date and time
            dateTimeFile = DateTime.Now.ToString("yyyy_MM_dd (hhmmtt)");

            // Create a new Excel application/file
            DataSet dset = new DataSet();
            dset.Tables.Add();
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wBook = xlApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet wSheet = (Microsoft.Office.Interop.Excel.Worksheet)wBook.Sheets[1];


            // selects the pull ticket data in the database
            OracleDataAdapter StrSQL = new OracleDataAdapter("SELECT * FROM PULL_TICKET_RECORD", con);
            DataSet DR2 = new DataSet();
            StrSQL.Fill(DR2);
            StrSQL.Dispose();
            System.Data.DataTable usertable = DR2.Tables[0];

            DataTable dt = dset.Tables[0];
            DataColumn dc;
            int colIndex = 0; // for excel column index
            int rowIndex = 0; // for excel row index
            int nextRowIndex = 1;

            object[,] arr; // creates an array object for storing data
            arr = new object[usertable.Rows.Count, usertable.Columns.Count]; // sets the length of the array object

            colIndex = colIndex + 1;

            for (int i = 0; i < usertable.Columns.Count; i++)
            {
                xlApp.Cells[rowIndex + 1, colIndex + i] = usertable.Columns[i].ColumnName.Replace("_", " ").Replace("BLANK", "");
            }

            //loops through the datatable and stores the value in the array object
            for (int r = 0; r < usertable.Rows.Count; r++)
            {
                DataRow dr = usertable.Rows[r];
                for (int c = 0; c < usertable.Columns.Count; c++)
                {
                    arr[r, c] = dr[c];
                }
            }


            Range c1 = wSheet.Cells[2, 1];
            Range c2 = wSheet.Cells[2 + usertable.Rows.Count - 1, usertable.Columns.Count];
            Range range = wSheet.Range[c1, c2]; // sets the excel range for the data

            range.Value = arr;//inserts the data in the specified excel range


            // formats the excel cells
            Range formatRange2 = wSheet.UsedRange;
            Range cell = formatRange2.Range[wSheet.Cells[rowIndex + 1, colIndex], wSheet.Cells[rowIndex + 1, colIndex + 21]];

            Microsoft.Office.Interop.Excel.Borders border = cell.Borders;

            border.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;
            cell.EntireRow.Font.Bold = true;
            cell.Interior.ColorIndex = 20;

            int rowRange = nextRowIndex + 1;
            Range formatRange3 = wSheet.UsedRange;
            Range cell2 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[usertable.Rows.Count + 1, 22]];

            Microsoft.Office.Interop.Excel.Borders border2 = cell.Borders;
            border2 = cell2.Borders;
            border2.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border2.Weight = 2.0;

            cell2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            cell2.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            wSheet.Columns.AutoFit();

            f2.Close();
            strFileName = saveFileDialog1.FileName;
            wBook.SaveAs(strFileName);

            MessageBox.Show("Excel file created!");
            wBook = xlApp.Workbooks.Open(strFileName); // opens the excel file
            xlApp.Visible = true;

            // releases/disposes the objects
            try
            {
                Marshal.ReleaseComObject(wSheet);
                Marshal.ReleaseComObject(wBook);
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                // handle the exception
            }
        }
        private void copyAllBacklog()
        {
            backlogdatagrid.SelectAll();
            DataObject dataObj = backlogdatagrid.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void copyAllRevised()
        {
            revisiondatagrid.SelectAll();
            DataObject dataObj = revisiondatagrid.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void copyAllMDR()
        {
            manualdrdatagrid.SelectAll();
            DataObject dataObj = manualdrdatagrid.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void btndownloadbcklog_Click(object sender, EventArgs e)
        {
            if (backlogdatagrid.Columns.Count == 0 || backlogdatagrid.Rows.Count == 0)
            {
                MessageBox.Show("There is no data to print!");
                return;
            }
            ExportBacklog();
        }

        private void ExportBacklog()
        {
            //prompts the user to select the location where the file will be saved
            string dateTimeFile = DateTime.Now.ToString("yyyy_MM_dd (hhmmtt)");

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "Excel Workbook|*.xls,*.xlsx";
            saveFileDialog1.Title = "Save Excel File";
            saveFileDialog1.FileName = $"BACKLOG_{dateTimeFile}.xls";
            saveFileDialog1.InitialDirectory = @"C:\";

            string strFileName;
            bool blnFileOpen;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(saveFileDialog1.FileName))
                {
                    try
                    {
                        using (FileStream fs = (FileStream)saveFileDialog1.OpenFile())
                        {
                            // No need to close the file stream as it will be closed by the using statement
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("File Not Found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    strFileName = saveFileDialog1.FileName;
                    blnFileOpen = false;

                    try
                    {
                        using (FileStream fileTemp = File.OpenWrite(strFileName))
                        {
                            // No need to close the file stream as it will be closed by the using statement
                        }
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
            }
            else
            {
                return;
            }

            // Show the form that tells the status of the application
            Form2 f2 = new Form2();
            f2.Show();
            f2.label2.Text = "Creating Excel File...";
            f2.Refresh();

            // Get the current date and time
            dateTimeFile = DateTime.Now.ToString("yyyy_MM_dd (hhmmtt)");

            // Create a new Excel application/file
            DataSet dset = new DataSet();
            dset.Tables.Add();
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wBook = xlApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet wSheet = (Microsoft.Office.Interop.Excel.Worksheet)wBook.Sheets[1];

            OracleDataAdapter StrSQL = new OracleDataAdapter("select DEL_DATE, DEL_TIME, JOB_NO, FACILITY, PARTNUMBER, PULL_TICKET_NUMBER, LINE, ORIGINAL_PULL, QTY_DEL, BACKLOGTYPE, HISTORY, BALANCE, GO_NUMBER from BACKLOG", con);
            DataSet DR2 = new DataSet();
            StrSQL.Fill(DR2);
            StrSQL.Dispose();
            DataTable usertable = DR2.Tables[0];

            object[,] arr = new object[usertable.Rows.Count, usertable.Columns.Count];

            int colIndex = 0; // for the column index in excel
            int rowIndex = 0; // for the row index in excel
            int nextRowIndex = 1;

            // adds values to excel cell
            colIndex = colIndex + 1;
            xlApp.Cells[rowIndex + 1, colIndex] = "DEL DATE";
            xlApp.Cells[rowIndex + 1, colIndex + 1] = "DEL TIME";
            xlApp.Cells[rowIndex + 1, colIndex + 2] = "JOB NO";
            xlApp.Cells[rowIndex + 1, colIndex + 3] = "FACILITY";
            xlApp.Cells[rowIndex + 1, colIndex + 4] = "PARTNUMBER";
            xlApp.Cells[rowIndex + 1, colIndex + 5] = "PULLTICKET NUMBER";
            xlApp.Cells[rowIndex + 1, colIndex + 6] = "LINE";
            xlApp.Cells[rowIndex + 1, colIndex + 7] = "PULL QTY";
            xlApp.Cells[rowIndex + 1, colIndex + 8] = "QTY DELIVERED";
            xlApp.Cells[rowIndex + 1, colIndex + 9] = "BACKLOG TYPE";
            xlApp.Cells[rowIndex + 1, colIndex + 10] = "HISTORY";
            xlApp.Cells[rowIndex + 1, colIndex + 11] = "BALANCE";
            xlApp.Cells[rowIndex + 1, colIndex + 12] = "GO NUMBER";

            // loops through the datatable and saves it to the array object
            for (int r = 0; r < usertable.Rows.Count; r++)
            {
                DataRow dr = usertable.Rows[r];
                for (int c = 0; c < usertable.Columns.Count; c++)
                {
                    arr[r, c] = dr[c];
                }
            }

            Range c1 = (Range)wSheet.Cells[2, 1];
            Range c2 = (Range)wSheet.Cells[2 + usertable.Rows.Count - 1, usertable.Columns.Count];
            Range range = wSheet.Range[c1, c2]; // sets the excel range for the data

            range.Value = arr;//inserts the data in the specified excel range

            // formats the excel cells
            Range formatRange2 = wSheet.UsedRange;
            Range cell = formatRange2.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 13]];

            Microsoft.Office.Interop.Excel.Borders border = cell.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2.0;

            cell.EntireRow.Font.Bold = true;
            cell.Interior.ColorIndex = 20;

            // formats the excel cells
            Range formatRange3 = wSheet.UsedRange;
            Range cell2 = formatRange3.Range[wSheet.Cells[1, 1], wSheet.Cells[usertable.Rows.Count + 1, 13]];

            // sets the border style
            Microsoft.Office.Interop.Excel.Borders border2 = cell2.Borders;
            border2.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border2.Weight = 2.0;

            cell2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            cell2.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            wSheet.Columns.AutoFit();

            f2.Close();
            strFileName = saveFileDialog1.FileName;
            wBook.SaveAs(strFileName); //saves the excel file

            MessageBox.Show("Excel file created!");
            wBook = xlApp.Workbooks.Open(strFileName); // opens the excel file
            xlApp.Visible = true;

            // releases/disposes the objects
            try
            {
                Marshal.ReleaseComObject(wSheet);
                Marshal.ReleaseComObject(wBook);
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                // handle the exception
            }
        }

        private void btndownloadrvsn_Click(object sender, EventArgs e)
        {
            if (revisiondatagrid.Columns.Count == 0 || revisiondatagrid.Rows.Count == 0)
            {
                // checks if the DataGridView has no data
                MessageBox.Show("There is no data to export!");
                return;
            }

            ExportRev(); // calls the method for creating Excel file
        }
        private void ExportRev()
        {
            // prompts the user to select the location where the file will be saved
            string dateTimeFile = DateTime.Now.ToString("yyyy_MM_dd (hhmmtt)");
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel Workbook|*.xls|Excel Workbook 2011|*.xlsx";
            saveFileDialog1.Title = "Save Excel File";
            saveFileDialog1.FileName = "PULLTICKET_REVISION_" + dateTimeFile + ".xls";
            saveFileDialog1.InitialDirectory = "C:/";

            string strFileName;
            bool blnFileOpen;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (saveFileDialog1.FileName != "")
                {
                    try
                    {
                        System.IO.FileStream fs = (System.IO.FileStream)saveFileDialog1.OpenFile();
                        fs.Close();
                    }
                    catch (Exception a)
                    {
                        MessageBox.Show("File Not Found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    strFileName = saveFileDialog1.FileName;
                    blnFileOpen = false;

                    try
                    {
                        System.IO.FileStream fileTemp = System.IO.File.OpenWrite(strFileName);
                        fileTemp.Close();
                    }
                    catch (Exception ex)
                    {
                        blnFileOpen = false;
                        return;
                    }

                    if (System.IO.File.Exists(strFileName))
                    {
                        System.IO.File.Delete(strFileName);
                    }
                }
            }
            else
            {
                return;
            }

            // shows the form telling the user the status of the application
            Form2 f2 = new Form2();
            f2.Show();
            f2.label2.Text = "Creating Excel File...";
            f2.Refresh();

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wBook = xlApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet wSheet = (Microsoft.Office.Interop.Excel.Worksheet)wBook.Sheets[1];

            // selects pull ticket revision data in the database
            OracleDataAdapter StrSQL = new OracleDataAdapter("select * from PULL_QUANTITY_REV", con);
            DataSet DR2 = new DataSet();
            StrSQL.Fill(DR2);
            StrSQL.Dispose();
            DataTable usertable = DR2.Tables[0];

            // creates an array object
            object[,] arr;

            // sets the length of the array object
            arr = new object[usertable.Rows.Count, usertable.Columns.Count];

            // for excel column index
            int colIndex = 0;

            // for excel row index
            int rowIndex = 0;

            // for excel next row index
            int nextRowIndex = 1;

            colIndex++;
            for (int i = 0; i < usertable.Columns.Count; i++)
            {
                xlApp.Cells[rowIndex + 1, colIndex + i] = usertable.Columns[i].ColumnName.Replace("_", " ").Replace("BLANK", "");
            }

            // loops through the datatable and stores the data in the array object
            for (int r = 0; r < usertable.Rows.Count; r++)
            {
                DataRow dr = usertable.Rows[r];
                for (int c = 0; c < usertable.Columns.Count; c++)
                {
                    arr[r, c] = dr[c];
                }
            }

            Range c1 = wSheet.Cells[2, 1];
            Range c2 = wSheet.Cells[2 + usertable.Rows.Count - 1, usertable.Columns.Count];
            Range range = wSheet.Range[c1, c2]; // sets the excel range for the data

            range.Value = arr;

            // formats the excel cells
            Range formatRange2 = wSheet.UsedRange;
            Range cell = formatRange2.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
            Microsoft.Office.Interop.Excel.Borders border = cell.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2.0;
            cell.EntireRow.Font.Bold = true;
            cell.Interior.ColorIndex = 20;

            // formats the excel cells
            Range formatRange3 = wSheet.UsedRange;
            Range cell2 = formatRange3.Range[wSheet.Cells[1, 1], wSheet.Cells[usertable.Rows.Count + 1, 8]];
            Microsoft.Office.Interop.Excel.Borders border2 = cell2.Borders;
            border2.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border2.Weight = 2.0;
            cell2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            cell2.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            wSheet.Columns.AutoFit();

            f2.Close();
            strFileName = saveFileDialog1.FileName;
            wBook.SaveAs(strFileName); //saves the excel file

            MessageBox.Show("Excel file created!");
            wBook = xlApp.Workbooks.Open(strFileName); // opens the excel file
            xlApp.Visible = true;

            // releases/disposes the objects
            try
            {
                Marshal.ReleaseComObject(wSheet);
                Marshal.ReleaseComObject(wBook);
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                // handle the exception
            }

        }

        private void backlogdatagrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == backlogdatagrid.Columns["PARTNUMBER"].Index)
            {
                e.CellStyle.Font = new System.Drawing.Font(e.CellStyle.Font, FontStyle.Bold);
            }

            if (e.ColumnIndex == backlogdatagrid.Columns["FACILITY"].Index)
            {
                e.CellStyle.Font = new System.Drawing.Font(e.CellStyle.Font, FontStyle.Bold);
            }

        }

        private void backlogdatagrid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            foreach (DataGridViewRow row in backlogdatagrid.Rows)
            {
                row.Cells["SELECTED"].Value = false;
            }
            checkedBacklogData.Clear(); // clear the checkedBacklogData list
        }

        //adding data to backlog
        private void backlogpanel_btnadd_Click(object sender, EventArgs e)
        {
            if(backlogpanel_btnadd.Text == "ADD")
            {
                var dupcheck = new OracleDataAdapter("Select * from BACKLOG where PULL_TICKET_NUMBER = '" + pnlpullticketnumber.Text + "' and LINE = '" + pnlline.Text + "'", con); //dataapater(for SQL syntax) , and connection
                DataTable dupchecktbl = new DataTable(); //temporary datatable used for checking data if there is duplicate.
                string newtime = "", backtype = ""; //temporary holders for DEL_TIME and BACKLOGTYPE


                
                /// condition statement for backlogtypedata
                if (pnlBacklogtype.Checked == true)
                {
                    backtype = "AS PER ADVISE";
                }
                else if (pnlBacklogtype.Checked != true)
                {
                    backtype = "";
                }
                /// /// /// /// /// /// /// ///
                dupcheck.Fill(dupchecktbl); //throw data to datatable.
                // // // // // Condition statements for adding data into backlog
                if (pnlDeldate.Text == "" || pnlDeltime.Text == "" || pnlfacility.Text == "" || asppartnumber.Text == "" || pnlpullqty.Text == "" || pnldelivered.Text == "" || pnlpullticketnumber.Text == "" || pnlline.Text == "") //condtion for null or empty field
                {
                    MessageBox.Show("Please fill up the required field!!!"); // messagebox show if there is an empty field
                }
                else if (int.Parse(pnlpullqty.Text) < int.Parse(pnldelivered.Text)) //condtion for pull quantity and delivered.
                {
                    MessageBox.Show("Delivered quantity should not be greater than the pull quantity!!"); // messagebox show if the delivered quantity inputed is higher than pull quantity
                }
                else if (dupchecktbl.Rows.Count > 0) // condition for duplicate data
                {
                    MessageBox.Show("Pull ticket number and line already exist!, please try another one!"); // messagebox show if the inputed pull ticket number and line inputed is already exist
                }
                else
                {
                    if (amCheckbox.Checked == true) //condition for deltime
                    {
                        if (pnlDeltime.SelectedIndex == -1) // condition for dropdownbox 
                        {
                            newtime = "10:00:00 AM";// set deltime
                        }
                        else
                        {
                            newtime = pnlDeltime.Text + ":00:00 AM"; // set deltime
                        }
                    }
                    else if (pmCheckbox.Checked == true) //condition for deltime
                    {
                        if (pnlDeltime.SelectedIndex == -1)
                        {
                            newtime = "10:00:00 PM";
                        }
                        else
                        {
                            newtime = pnlDeltime.Text + ":00:00 PM";
                        }
                    }
                    //temporary holder for insterting/adding data to sql
                    string insRec = "insert into BACKLOG (DEL_DATE, DEL_TIME, FACILITY, PARTNUMBER, PULL_TICKET_NUMBER, LINE, QTY_DEL, ORIGINAL_PULL, BACKLOGTYPE, BALANCE) values ('" + pnlDeldate.Value.ToString("MM/dd/yyyy") + "','" + newtime + "','" + pnlfacility.Text.ToUpper() + "','" + asppartnumber.Text.ToUpper() + "','" + pnlpullticketnumber.Text.ToUpper() + "','" + pnlline.Text.ToUpper() + "','" + pnldelivered.Text.ToUpper() + "','" + pnlpullqty.Text.ToUpper() + "','" + backtype + "','" + (int.Parse(pnlpullqty.Text) - int.Parse(pnldelivered.Text)) + "')";
                    OracleCommand insRecAd = new OracleCommand(insRec, con); //command for inserting and openning connection of sql
                    insRecAd.ExecuteNonQuery();
                    //backlogdatagrid.DataSource = dtbacklog; //throw new sqldata to datagridview
                    MessageBox.Show("Added Successfully !"); //messagebox show if added successfully
                    pnl_asperadvice.Width = 0; // hide panel
                    pnlBacklogtype.Checked = false; // clear field
                    pnlDeltime.Text = string.Empty; // clear field
                    pnlfacility.Text = string.Empty; // clear field
                    asppartnumber.Text = string.Empty; // clear field
                    pnlpullqty.Text = string.Empty; // clear field
                    pnldelivered.Text = string.Empty; // clear field
                    pnlpullticketnumber.Text = string.Empty; // clear field
                    pnlline.Text = string.Empty; // clear field
                    amCheckbox.Checked = false; // clear field
                    pmCheckbox.Checked = false; // clear field
                    backlogdatagrid.Location = new Point(0, 86); // restore datagridview location
                    searchby.Visible = true; // show label
                    backlogBox.Visible = true; //show dropdown box used to specify what to search
                    backlogTxt.Visible = true; // show textbox used to search data
                    selectallchk.Visible = true; // show checkbox used to filter data
                    btnadd.Enabled = true; //enabling add button
                    btnupdate.Enabled = true; // enabling update button
                    btndlt.Enabled = true; // enabling delete button
                    showbacklogdata(); // show sql data
                    clearCheckedBacklogData();
                    totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;
                }
            }
            else if(backlogpanel_btnadd.Text == "UPDATE")
            {

                if(checkedBacklogData.Count == 1)
                {
                    asppartnumber.Enabled = false; pnlpullticketnumber.Enabled = false; pnlline.Enabled = false; 
                    if (pnlpullqty.Text == "" || pnldelivered.Text == "")
                    {
                        MessageBox.Show("Please Fill out the required field");
                    }
                    else if(int.Parse(pnldelivered.Text.ToString()) > int.Parse(pnlpullqty.Text.ToString()))
                    {
                        MessageBox.Show("Delivered quantity should not be greater than the pull quantity");
                    }
                    else
                    {
                        var dupcheck = new OracleDataAdapter("select * from BACKLOG where PULL_TICKET_NUMBER = '" + pnlpullticketnumber.Text.ToString().ToUpper() + "' and LINE = '" + pnlline.Text.ToString().ToUpper() + "'", con);
                        DataTable dupchecktbl = new DataTable();
                        dupcheck.Fill(dupchecktbl);
                        dupcheck.Dispose();
                        string backtype = string.Empty;
                        if (pnlBacklogtype.Checked == true)
                        {
                            backtype = "AS PER ADVISE";
                        }
                        else
                        {
                            backtype = "";
                        }
                        string insRec = "update BACKLOG set FACILITY = '" + pnlfacility.Text.ToString().ToUpper() + "', QTY_DEL = " + pnldelivered.Text.ToString().ToUpper() + ", ORIGINAL_PULL= " + pnlpullqty.Text.ToString().ToUpper() + ", PULL_TICKET_NUMBER = '" + pnlpullticketnumber.Text.ToString().ToUpper() + "', LINE = '" + pnlline.Text.ToString().ToUpper() + "', BACKLOGTYPE = '" + backtype + "', BALANCE = " + (int.Parse(pnlpullqty.Text) - int.Parse(pnldelivered.Text)) + " WHERE PULL_TICKET_NUMBER ='" + pnlpullticketnumber.Text + "'AND LINE ='" + pnlline.Text + "'";
                        var insrecadd = new OracleCommand(insRec, con);
                        insrecadd.ExecuteNonQuery();
                        MessageBox.Show("Updated Successfully !"); //messagebox show if added successfully

                        pnl_asperadvice.Width = 0; // hide panel
                        pnlBacklogtype.Checked = false; // clear field
                        pnlDeltime.Text = string.Empty; // clear field
                        pnlfacility.Text = string.Empty; // clear field
                        asppartnumber.Text = string.Empty; // clear field
                        pnlpullqty.Text = string.Empty; // clear field
                        pnldelivered.Text = string.Empty; // clear field
                        pnlpullticketnumber.Text = string.Empty; // clear field
                        pnlline.Text = string.Empty; // clear field
                        amCheckbox.Checked = false; // clear field
                        pmCheckbox.Checked = false; // clear field
                        backlogdatagrid.Location = new Point(0, 86); // restore datagridview location
                        searchby.Visible = true; // show label
                        backlogBox.Visible = true; //show dropdown box used to specify what to search
                        backlogTxt.Visible = true; // show textbox used to search data
                        selectallchk.Visible = true; // show checkbox used to filter data
                        btnadd.Enabled = true; //enabling add button
                        btnupdate.Enabled = true; // enabling update button
                        btndlt.Enabled = true; // enabling delete button
                        showbacklogdata(); // show sql data
                        clearCheckedBacklogData();
                        totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;
                    }
                    

                }
                else if(checkedBacklogData.Count > 1)
                {
                    pnlfacility.Enabled = false; asppartnumber.Enabled = false; pnlpullticketnumber.Enabled = false; pnlline.Enabled = false; pnlDeldate.Enabled = false;
                    pnlpullqty.Enabled = false; pnldelivered.Enabled = false; pnlDeltime.Enabled = false; amCheckbox.Enabled = false; pmCheckbox.Enabled = false;

                    //pnlBacklogtype.Checked = true;
                    string backtype = string.Empty;
                    if (pnlBacklogtype.Checked == true)
                    {
                        backtype = "AS PER ADVISE";
                    }
                    else
                    {
                        backtype = "";
                    }
                    foreach(var item in checkedBacklogData)
                    {
                        string insrecM = "update BACKLOG set BACKLOGTYPE = '"+ backtype + "' WHERE PULL_TICKET_NUMBER = '" + item.PullTicketNumber + "' AND LINE = '" + item.Line + "'";
                        try
                        {
                            var insrecMadd = new OracleCommand(insrecM, con);
                            insrecMadd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error updating item: " + ex.Message);
                        }
                    }
                    MessageBox.Show("Items Updated Successfully !");

                    pnl_asperadvice.Width = 0; // hide panel
                    pnlBacklogtype.Checked = false; // clear field
                    pnlDeltime.Text = string.Empty; // clear field
                    pnlfacility.Text = string.Empty; // clear field
                    asppartnumber.Text = string.Empty; // clear field
                    pnlpullqty.Text = string.Empty; // clear field
                    pnldelivered.Text = string.Empty; // clear field
                    pnlpullticketnumber.Text = string.Empty; // clear field
                    pnlline.Text = string.Empty; // clear field
                    amCheckbox.Checked = false; // clear field
                    pmCheckbox.Checked = false; // clear field
                    backlogdatagrid.Location = new Point(0, 86); // restore datagridview location
                    searchby.Visible = true; // show label
                    backlogBox.Visible = true; //show dropdown box used to specify what to search
                    backlogTxt.Visible = true; // show textbox used to search data
                    selectallchk.Visible = true; // show checkbox used to filter data
                    btnadd.Enabled = true; //enabling add button
                    btnupdate.Enabled = true; // enabling update button
                    btndlt.Enabled = true; // enabling delete button
                    showbacklogdata(); // show sql data
                    clearCheckedBacklogData();
                    totalbacklog.Text = "TOTAL DATA: " + backlogdatagrid.Rows.Count;

                }
                else
                {
                    MessageBox.Show("No Data Entry!");
                }
                

            }
        }
        //amcheckbox method
        private void amCheckbox_Click(object sender, EventArgs e)
        {
            amCheckbox.Checked = true;
        }
        //pmcheckbox method
        private void pmCheckbox_Click(object sender, EventArgs e)
        {
            pmCheckbox.Checked=true;
        }
        //button add method for mdr
        private void btnAddMdr_Click(object sender, EventArgs e)
        {
        }
        ////Select/deselect data method//
        private void selectallchk_CheckedChanged(object sender, EventArgs e)
        {
            if (selectallchk.Checked)
            {
                selectallchk.Text = "Deselect All";
            }
            else
            {
                selectallchk.Text = "Select All";
            }
            //loop used to set all checkbox unselected by default
            foreach (DataGridViewRow row in backlogdatagrid.Rows)
            {
                DataGridViewCheckBoxCell chk = row.Cells[0] as DataGridViewCheckBoxCell;
                if (chk != null)
                {
                    chk.Value = selectallchk.Checked;
                }
                else
                {
                    chk.Value = !selectallchk.Checked;
                }
            }

        }
        // filter checkbox for pull ticket record
        private void recordfilterchk_CheckedChanged(object sender, EventArgs e)
        {
            DataTable dttemp = new DataTable(); //temporary datatable
            DataView dv = new DataView(dt); //temporary dataview

            if (facilbox.Text == "ALL") //condition if combobox selected to all
            {
                recorddatagrid.DataSource = dt;
                if (partnumbox.Text == "")//condition if partnumber search textbox in null/empty 
                {
                    //filter data with zero PUL_QTY//
                    if (recordfilterchk.Checked == true)
                    {
                        dt.DefaultView.RowFilter = "[PULL_QTY] > 0";

                    }
                    else
                    {
                        dt.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }
                }
                else
                {
                    dv.RowFilter = "PARTNUMBER LIKE '%" + partnumbox.Text.ToUpper() + "%'";
                    dttemp = dv.ToTable();
                    recorddatagrid.DataSource = dttemp;
                    //filter data with zero PUL_QTY//
                    if (recordfilterchk.Checked == true)
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0";
                    }
                    else
                    {
                        dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0";
                    }
                }
            }
            else
            {

                switch (facilbox.Text)
                {
                    case "CAV3":
                        dv.RowFilter = "FACILITY = 'CAV3'";
                        break;
                    case "DKP":
                        dv.RowFilter = "FACILITY = 'DKP'";
                        break;
                    case "DANAM T":
                        dv.RowFilter = "FACILITY = 'DANAM T'";
                        break;
                    case "DANAM":
                        dv.RowFilter = "FACILITY = 'DANAM'";
                        break;
                    case "MACRO":
                        dv.RowFilter = "FACILITY = 'MACRO'";
                        break;
                    case "CAV2":
                        dv.RowFilter = "FACILITY = 'CAV2'";
                        break;
                    case "CAV5-CPPK":
                        dv.RowFilter = "FACILITY = 'CAV5-CPPK'";
                        break;
                    case "CLP/CAV5":
                        dv.RowFilter = "FACILITY = 'CLP/CAV5'";
                        break;
                    case "CLP":
                        dv.RowFilter = "FACILITY = 'CLP'";
                        break;             

                }
                dttemp = dv.ToTable();
                
                if (partnumbox.Text == "")
                {
                    dv.RowFilter = "FACILITY = '" + facilbox.Text + "'";
                    dttemp = dv.ToTable();
                    recorddatagrid.DataSource = dttemp; 
                }
                else
                {
                    dv.RowFilter = "FACILITY = '" + facilbox.Text + "' and PARTNUMBER LIKE '%" + partnumbox.Text.ToUpper() + "%'";
                    dttemp = dv.ToTable();
                    recorddatagrid.DataSource = dttemp;
                }
                if (recordfilterchk.Checked)
                {
                    dttemp.DefaultView.RowFilter = "[PULL_QTY] > 0"; // Line 1
                }
                else
                {
                    dttemp.DefaultView.RowFilter = "[PULL_QTY] >= 0"; // Line 2
                }
            }
        }
    }
}