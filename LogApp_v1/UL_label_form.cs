using System;
using System.Drawing;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.IO;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
namespace LogApp_v1
{
    public partial class UL_label_form : Form
    {
        public OpenFileDialog openFD = new OpenFileDialog();
        public Worksheet xlWorkSheet;
        public Workbook xlWorkBook;
        public Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        public Microsoft.Office.Interop.Excel.Application oExcel2 = new Microsoft.Office.Interop.Excel.Application();
        private string connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.50.40)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User ID=mec;Password=mec2024"; //oracle host source
        private OracleConnection con = new OracleConnection();
        public UL_label_form()
        {
            InitializeComponent();
            this.BackColor = System.Drawing.Color.LimeGreen;
            this.TransparencyKey = System.Drawing.Color.LimeGreen;
            ul_label_panel.Location = new Point(680, 90);
            connectdata();
        }
        private void connectdata()
        {
            con.ConnectionString = connectionString; //declaring connection from host
            con.Open(); //to open database connection
        }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            this.Close();
        }
        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);
            this.Close();
        }
        private void closeImgBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void updateLblBtn_Click(object sender, EventArgs e)
        {
            //openFD = new OpenFileDialog();
            openFD.InitialDirectory = "C:";
            openFD.Filter = "All Files (*.*)|*.*";
            openFD.Title = "Choose a File";
            openFD.FilterIndex = 2;
            openFD.RestoreDirectory = true;


            if (openFD.ShowDialog().Equals(DialogResult.OK))
            {
                if (File.Exists(openFD.FileName))
                {
                    Form2 f2 = new Form2();
                    string deLUL = "delete from UL_LABEL";
                    OracleCommand delULAd = new OracleCommand(deLUL, con);
                    delULAd.ExecuteNonQuery();

                    string ulLabelFile;
                    f2.Show();
                    f2.label2.Text = "READING DATA. . . . .";
                    f2.Refresh();
                    ulLabelFile = openFD.FileName;

                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlApp.DisplayAlerts = false;
                    xlWorkBook = xlApp.Workbooks.Open(ulLabelFile, Microsoft.Office.Interop.Excel.XlFileAccess.xlReadOnly);


                    DataColumn column = new DataColumn();
                    int sheetRows = 0;
                    int rangeCount = 0;
                    string[] PARTNUMBER = new string[0], UL_LABEL = new string[0], REMARKS = new string[0];


                    for (int a = 1; a <= xlWorkBook.Sheets.Count; a++)
                    {
                        xlWorkSheet = xlWorkBook.Sheets[a];
                        int lrow = xlWorkSheet.Range["A" + xlWorkSheet.Rows.Count].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;
                        Microsoft.Office.Interop.Excel.Range range = xlWorkSheet.Range["A4:G" + lrow];
                        object[,] data = (object[,])range.Value;

                        if (a == 1)
                        {
                            rangeCount = rangeCount + (range.Rows.Count - 1);
                        }
                        else
                        {
                            rangeCount = rangeCount + (range.Rows.Count);
                        }
                        Array.Resize(ref PARTNUMBER, rangeCount);
                        Array.Resize(ref UL_LABEL, rangeCount);
                        Array.Resize(ref REMARKS, rangeCount);

                        f2.progressBar1.Maximum = xlWorkBook.Sheets.Count;
                        f2.progressBar1.Value = a;
                        f2.label2.Text = "IMPORTING DATA. . .  .";
                        f2.Refresh();


                        for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                        {
                            f2.progressBar1.Maximum = range.Columns.Count;
                            f2.progressBar1.Value = cCnt;
                            f2.label2.Text = "READING DATA. . .  .";
                            f2.Refresh();

                            for (int rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                            {
                                string cellVal = string.Empty;
                                cellVal = (string)data[rCnt, cCnt];

                                DataRow[] Row;


                                if (cCnt == 3)
                                {
                                    if (cellVal == "")
                                    {
                                        PARTNUMBER[rCnt - 1 + sheetRows] = cellVal;
                                    }
                                    else
                                    {
                                        PARTNUMBER[rCnt - 1 + sheetRows] = cellVal.ToString().ToUpper();
                                    }

                                }
                                else if (cCnt == 5)
                                {
                                    if (cellVal == "")
                                    {
                                        UL_LABEL[rCnt - 1 + sheetRows] = cellVal;
                                    }
                                    else
                                    {
                                        UL_LABEL[rCnt - 1 + sheetRows] = cellVal.ToString().ToUpper();
                                    }
                                }
                                else if (cCnt == 7)
                                {
                                    if (cellVal == "")
                                    {
                                        REMARKS[rCnt - 1 + sheetRows] = cellVal;
                                    }
                                    else
                                    {
                                        REMARKS[rCnt - 1 + sheetRows] = cellVal.ToString().ToUpper();
                                    }
                                }

                            }


                        }


                        sheetRows = range.Rows.Count + sheetRows;

                    }

                    OracleParameter p_PARTNUMBER = new OracleParameter();
                    p_PARTNUMBER.OracleDbType = OracleDbType.Varchar2;
                    p_PARTNUMBER.Value = PARTNUMBER;
                    OracleParameter p_UL_LABEL = new OracleParameter();
                    p_UL_LABEL.OracleDbType = OracleDbType.Varchar2;
                    p_UL_LABEL.Value = UL_LABEL;
                    OracleParameter p_REMARKS = new OracleParameter();
                    p_REMARKS.OracleDbType = OracleDbType.Varchar2;
                    p_REMARKS.Value = REMARKS;

                    OracleCommand cmd = new OracleCommand();
                    cmd = con.CreateCommand();
                    cmd.CommandText = "insert into UL_LABEL values(:3, :5, :7)";


                    cmd.ExecuteNonQuery();

                    f2.label2.Text = "IMPORTING COMPLETE.";
                    f2.Close();
                    xlWorkBook.Close();
                    Form1 f1 = new Form1();

                    try
                    {
                        xlApp.Quit();
                    }
                    catch (Exception ex)
                    {
                    }
                    f1.ReleaseObject(xlWorkSheet);
                    f1.ReleaseObject(xlWorkBook);
                    f1.ReleaseObject(xlApp);

                    this.Close();
                }
                else
                {
                    MessageBox.Show("File not Found!", "warning" ,MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
    }
}
