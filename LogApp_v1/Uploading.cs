using System;
using System.Data;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using DateTime = System.DateTime;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


namespace LogApp_v1
{
    public partial class Uploading : Form
    {
        public OpenFileDialog openFD = new OpenFileDialog();
        bool uploadOne, uploadTwo;
        DataTable uploadOneTbl = new DataTable(), 
            newData = new DataTable(),
            uploadTwoTbl = new DataTable();
        public Worksheet xlWorksheet;
        public Workbook xlWorkbook;
        public Range xlRange;
        public Excel.Application oExcel2 = new Excel.Application();
        public Excel.Application xlApp = new Excel.Application();
        private object ptkFileName;
        Form1 f1 = new Form1();
        Form2 f2 = new Form2();
        private string connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.50.40)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User ID=mec;Password=mec2024"; //oracle host source
        private OracleConnection con = new OracleConnection();

        public Uploading()
        {
            InitializeComponent();
            connectdata(); //calling this metohd to open the sql connection
            uploadOne = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (uploadOne) // checks if the user already uploaded the file
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
                f1.ptkFileName.Text = openFD.FileName;
                if (!string.IsNullOrEmpty(FileSystem.Dir(openFD.FileName)))
                {
                    if (!openFD.SafeFileName.ToString().ToUpper().Contains("DELIVERY"))
                    {
                        Interaction.MsgBox("The file you are trying to import is named " + openFD.SafeFileName + Microsoft.VisualBasic.Constants.vbCrLf + Microsoft.VisualBasic.Constants.vbCrLf + "Make sure you are importing the correct file!");
                        return;
                    }
                    else
                    {
                        if (strFileName != "")
                        {
                            //f2.Show();
                            f2.label2.Text = "Reading Data . . .";
                            f2.Refresh();
                            try
                            {
                                uploadSave(strFileName, openFD.SafeFileName);
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("Please close the excel application then try importing again!");
                                MessageBox.Show(ex.Message);
                                f2.Close();
                            }
                            
                        }
                        else
                        {
                            MessageBox.Show("FILE NOT FOUND!");
                        }
                    }
                }
            }

        }
        private void uploadSave(string myfiledirect, string filename)
        {
            uploadOneTbl = new DataTable(); // creates a new instance of the datatable for storing data
            Form2 f2 = new Form2();
            f2.Show();
            oExcel2 = new Excel.Application();
            oExcel2.DisplayAlerts = false;
            xlWorkbook = oExcel2.Workbooks.Open(myfiledirect, XlFileAccess.xlReadOnly);
            DataColumn Column = new DataColumn();
            int sheetRows = 0;



            for (int a = 1; a <= xlWorkbook.Sheets.Count; a++) //loops through all the worksheets available in the excel file
            {
                f2.label2.Text = "Reading Data . . .";
                f2.Refresh();

                xlWorksheet = xlWorkbook.Worksheets[a];
                int lRow = xlWorksheet.Range["E" + xlWorksheet.Rows.Count].End[XlDirection.xlUp].Row;
                xlRange = xlWorksheet.Range["A2:K" + lRow]; ; //sets the range of the data in the excel file
                object[,] data = (object[,])xlRange.Value; ; //stores the value in an array

                //Create new Column in DataTable
                for (int cCnt = 1; cCnt <= xlRange.Columns.Count + 1; cCnt++)
                {
                    f2.label2.Text = "Extracting Data . . .";
                    f2.Refresh();
                    if (a == 1)
                    {
                        Column = new DataColumn();
                        Column.DataType = typeof(string);
                        Column.ColumnName = cCnt.ToString();
                        uploadOneTbl.Columns.Add(Column);
                    }
                    else
                    {
                        Column = uploadOneTbl.Columns["" + cCnt + ""];
                    }

                    //Create row for Data Table
                    for (int rCnt = 3; rCnt <= xlRange.Rows.Count; rCnt++)
                    {
                        f2.progressBar1.Maximum = xlRange.Rows.Count;
                        f2.progressBar1.Value = rCnt;
                        f2.label2.Text = "Importing Data . . .";
                        f2.Refresh();

                        string CellVal = string.Empty;

                        if (cCnt != xlRange.Columns.Count + 1)
                        {
                            CellVal = Convert.ToString(data[rCnt, cCnt]);
                        }

                        DataRow Row;

                        //Adds row to the DataTable
                        if (cCnt == 1)
                        {
                            Row = uploadOneTbl.NewRow();
                            Row[Column.ColumnName.ToString()] = CellVal;
                            uploadOneTbl.Rows.Add(Row);
                        }
                        else if (cCnt == xlRange.Columns.Count + 1)
                        {
                            Row = uploadOneTbl.Rows[(rCnt - 3) + sheetRows];
                            Row[Column.ColumnName.ToString()] = data[1, 1].ToString().Replace("FACILITY: ", "").ToUpper();
                        }
                        else
                        {
                            Row = uploadOneTbl.Rows[(rCnt - 3) + sheetRows];
                            Row[Column.ColumnName.ToString()] = CellVal;
                        }
                    }
                }
                sheetRows = xlRange.Rows.Count - 2 + sheetRows;
            }
            uploadOneTbl.AcceptChanges();

            //f2.label2.Text = "IMPORTING COMPLETE!";
            MessageBox.Show("Import Successfully", "Success");
            xlWorkbook.Close();
            f2.Close();
            uploadOne = true;
            btnDel.BackColor = System.Drawing.Color.Navy;
            btnDel.Enabled = false;
            try
            {
                xlApp.Quit();
            }
            catch (Exception ex)
            {

            }
            ReleaseObject(xlWorksheet);
            ReleaseObject(xlWorkbook);
            ReleaseObject(oExcel2);

            string delUpLog = "DELETE FROM UPLOAD_LOG";
            OracleCommand delUpLogAd = new OracleCommand(delUpLog, con);
            delUpLogAd.ExecuteNonQuery();

            string delDate, delTime, delDatePrev = null, delTimePrev = null;
            for (int c = 0; c < uploadOneTbl.Rows.Count; c++)
            {
                if (uploadOneTbl.Rows[c][0] is DBNull)
                {
                    delDate = delDatePrev;
                }
                else
                {
                    delDate = uploadOneTbl.Rows[c][0].ToString();
                    delDatePrev = uploadOneTbl.Rows[c][0].ToString();
                }

                if (uploadOneTbl.Rows[c][1] is DBNull)
                {
                    delTime = delTimePrev;
                }
                else
                {
                    delTime = uploadOneTbl.Rows[c][1].ToString();
                    delTimePrev = uploadOneTbl.Rows[c][1].ToString();
                }

                string insPullTick = "insert into UPLOAD_LOG values('" + delDate + "','" + delTime + "','" + uploadOneTbl.Rows[c][2] + "','" + uploadOneTbl.Rows[c][3] + "','" + uploadOneTbl.Rows[c][4] + "','" + uploadOneTbl.Rows[c][5] + "','" + uploadOneTbl.Rows[c][6] + "','" + uploadOneTbl.Rows[c][7] + "','" + uploadOneTbl.Rows[c][8] + "','" + uploadOneTbl.Rows[c][9] + "','" + uploadOneTbl.Rows[c][10] + "','" + uploadOneTbl.Rows[c][11] + "')";
                OracleCommand insPullTickAd = new OracleCommand(insPullTick, con);
                insPullTickAd.ExecuteNonQuery();

            }
            newData = new DataTable();
            string backCheck = @"SELECT PARTNUMBER, GO_NUM, GO_LINE, COUNT(*) AS COUNT, SUM(DELIVERY) As TOTAL 
                     FROM UPLOAD_LOG 
                     GROUP BY PARTNUMBER, GO_NUM, GO_LINE, FACILITY";
            OracleDataAdapter adapter2 = new OracleDataAdapter(backCheck, con); //selects data in the database
            adapter2.Fill(newData);
            //stores the data in the datatable
        }

        private void btnDL_Click(object sender, EventArgs e)
        {
            if (!uploadOne || !uploadTwo) //checks if the two files haven't been imported
            {
                MessageBox.Show("Please upload the files before printing!");
                return;
            }
            //opens a dialog for the user to select the location of the file to be saved
            string dateTimeFile = DateTime.Now.ToString("yyyy_MM_dd (hhmmtt)");
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel Workbook|*.xls|Excel Workbook 2011|*.xlsx";
            saveFileDialog1.Title = "Save Excel File";
            saveFileDialog1.FileName = "UPLOADING" + dateTimeFile + ".xls";
            saveFileDialog1.InitialDirectory = "C:/";

            string strFileName;
            bool blnFileOpen;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK) // checks if the user selects "OK" in the dialog
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

            DataTable uploadFinalTbl = new DataTable();

            // adds columns to the datatable
            uploadFinalTbl.Columns.Add("DEL_DATE", typeof(string));
            uploadFinalTbl.Columns.Add("DEL_TIME", typeof(string));
            uploadFinalTbl.Columns.Add("PARTNUMBER", typeof(string));
            uploadFinalTbl.Columns.Add("PULL", typeof(string));
            uploadFinalTbl.Columns.Add("PULL_TICKET_NUMBER", typeof(string));
            uploadFinalTbl.Columns.Add("LINE", typeof(string));
            uploadFinalTbl.Columns.Add("QTY", typeof(string));
            uploadFinalTbl.Columns.Add("DATE", typeof(string));
            uploadFinalTbl.Columns.Add("APC_DR", typeof(string));

            DataRow[] z2;
            DataRow[] z;
            int totalDel = 0;
            string delDate, delTime, delTimePrev = null, delDatePrev = null;
            string apcDr;

            for (int a = 0; a < uploadOneTbl.Rows.Count; a++)
            {
                string expr = $"PARTNUMBER = '{uploadOneTbl.Rows[a][2].ToString().ToUpper()}' AND GO_NUM = '{uploadOneTbl.Rows[a][9].ToString().ToUpper()}' AND GO_LINE = '{uploadOneTbl.Rows[a][10].ToString().ToUpper()}'";
                z = newData.Select(expr);

                totalDel = int.Parse(z[0][4].ToString());
                string expr2 = $"([6] LIKE '{uploadOneTbl.Rows[a][2].ToString().ToUpper()}%' or [6] = '{uploadOneTbl.Rows[a][2].ToString().ToUpper()}') AND [5] = '{uploadOneTbl.Rows[a][9].ToString().ToUpper()}' AND [13] = '{totalDel}'";
                z2 = uploadTwoTbl.Select(expr2);

                if (uploadOneTbl.Rows[a][0] is DBNull)
                {
                    delDate = delDatePrev;
                }
                else
                {
                    delDate = uploadOneTbl.Rows[a][0].ToString();
                    delDatePrev = uploadOneTbl.Rows[a][0].ToString();
                }

                if (uploadOneTbl.Rows[a][1] is DBNull)
                {
                    delTime = delTimePrev;
                }
                else
                {
                    delTime = uploadOneTbl.Rows[a][1].ToString();
                    delTimePrev = uploadOneTbl.Rows[a][1].ToString();
                }

                DataRow pq = uploadFinalTbl.NewRow();
                pq["DEL_DATE"] = delDate;
                pq["DEL_TIME"] = delTime;
                pq["PARTNUMBER"] = uploadOneTbl.Rows[a][2].ToString();
                pq["PULL"] = uploadOneTbl.Rows[a][3].ToString();
                pq["PULL_TICKET_NUMBER"] = uploadOneTbl.Rows[a][4].ToString();
                pq["LINE"] = uploadOneTbl.Rows[a][5].ToString();
                pq["QTY"] = uploadOneTbl.Rows[a][6].ToString();
                pq["DATE"] = uploadOneTbl.Rows[a][7].ToString();
                if (z2.Length == 0)
                {
                    apcDr = "";
                }
                else
                {
                    apcDr = z2[0][1].ToString();
                }

                pq["APC_DR"] = apcDr;
                uploadFinalTbl.Rows.Add(pq);
                uploadFinalTbl.AcceptChanges();
            }
            //f2.Show();
            f2.label2.Text = "Creating Excel File . . .";
            f2.Refresh();

            //creates a new excel workbook
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wBook = excel.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet wSheet = wBook.ActiveSheet;
            wSheet.Name = "WITH DR";


            // Iterate through the rows in the DataTable
            for (int i = 1; i < uploadFinalTbl.Rows.Count; i++)
            {
                // Check if the current row's delivery date is null or empty
                if (uploadFinalTbl.Rows[i]["DEL_DATE"] == null || string.IsNullOrEmpty(uploadFinalTbl.Rows[i]["DEL_DATE"].ToString()))
                {
                    // Get the previous row's delivery date
                    if (i > 0 && uploadFinalTbl.Rows[i - 1]["DEL_DATE"] != null)
                    {
                        string prevDeliveryDate = uploadFinalTbl.Rows[i - 1]["DEL_DATE"].ToString();
                        // Update the current row's delivery date with the previous row's delivery date
                        uploadFinalTbl.Rows[i]["DEL_DATE"] = prevDeliveryDate;
                    }
                }
            }

            // Iterate through the rows in the DataTable
            for (int i = 1; i < uploadFinalTbl.Rows.Count; i++)
            {
                // Check if the current row's delivery time is null or empty
                if (uploadFinalTbl.Rows[i]["DEL_TIME"] == null || string.IsNullOrEmpty(uploadFinalTbl.Rows[i]["DEL_TIME"].ToString()))
                {
                    // Get the previous row's delivery date
                    if (i > 0 && uploadFinalTbl.Rows[i - 1]["DEL_TIME"] != null)
                    {
                        string prevDeliveryDate = uploadFinalTbl.Rows[i - 1]["DEL_TIME"].ToString();
                        // Update the current row's delivery time with the previous row's delivery time
                        uploadFinalTbl.Rows[i]["DEL_TIME"] = prevDeliveryDate;
                    }
                }
            }


            // Clone the DataTable to remove null values
            DataTable uploadFinalTblNotNull = uploadFinalTbl.Clone();
            foreach (DataRow row in uploadFinalTbl.Rows)
            {
                if (!string.IsNullOrEmpty(row["APC_DR"].ToString()))
                {
                    uploadFinalTblNotNull.ImportRow(row);
                }
            }

            object[,] arr = new object[uploadFinalTblNotNull.Rows.Count, uploadFinalTblNotNull.Columns.Count];

            int colIndex = 0; //column index in excel
            int rowIndex = 0; //row index in excel
            int nextRowIndex = 1;

            //adds values to the excel cells
            colIndex++;
            excel.Cells[rowIndex + 1, colIndex] = "DEL Date";
            excel.Cells[rowIndex + 1, colIndex + 1] = "DEL Time";
            excel.Cells[rowIndex + 1, colIndex + 2] = "Part Number";
            excel.Cells[rowIndex + 1, colIndex + 3] = "Pull Qty";
            excel.Cells[rowIndex + 1, colIndex + 4] = "Pull Ticket No";
            excel.Cells[rowIndex + 1, colIndex + 5] = "Line";
            excel.Cells[rowIndex + 1, colIndex + 6] = "QTY Deliverd";
            excel.Cells[rowIndex + 1, colIndex + 7] = "Date";
            excel.Cells[rowIndex + 1, colIndex + 8] = "APC DR#";
            excel.Cells[rowIndex + 1, colIndex + 9] = "Remarks";
            excel.Cells[rowIndex + 1, colIndex + 9].ColumnWidth = 25;



            //insert data to the array
            for (int r = 0; r < uploadFinalTblNotNull.Rows.Count; r++)
            {
                DataRow dr = uploadFinalTblNotNull.Rows[r];
                for (int c = 0; c < uploadFinalTblNotNull.Columns.Count; c++)
                {
                    arr[r, c] = dr[c];
                }
            }


            //specify excel range
            Range c1 = wSheet.Cells[2, 1];
            Range c2 = wSheet.Cells[2 + uploadFinalTblNotNull.Rows.Count - 1, uploadFinalTblNotNull.Columns.Count];
            Range range = wSheet.Range[c1, c2];

            range.Value = arr;//inserts the values to the excel range

            Range formatRange2 = wSheet.UsedRange;
            Range cell = formatRange2.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 10]];
            Microsoft.Office.Interop.Excel.Borders border = cell.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2.0;
            cell.EntireRow.Font.Bold = true;
            cell.Interior.ColorIndex = 20;

            Range formatRange3 = wSheet.UsedRange;
            Range cell2 = formatRange3.Range[wSheet.Cells[1, 1], wSheet.Cells[uploadFinalTblNotNull.Rows.Count + 1, 10]];
            Microsoft.Office.Interop.Excel.Borders border2 = cell2.Borders;
            border2.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border2.Weight = 2.0;
            cell2.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            cell2.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            cell2 = wSheet.Range[wSheet.Cells[2, 2], wSheet.Cells[uploadFinalTblNotNull.Rows.Count + 1, 2]];
            cell2.NumberFormat = "h:mm:ss AM/PM";

            /////////////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////////////////////////////////////////////////

            //Create a new worksheet
            Microsoft.Office.Interop.Excel.Worksheet newSheet = wBook.Worksheets.Add();
            newSheet.Name = "NO DR";

            // Copy the column headers to the new sheet
            for (int c = 1; c <= uploadFinalTbl.Columns.Count; c++)
            {
                newSheet.Cells[1, c] = wSheet.Cells[1, c].Value;
            }

            // Clone the DataTable to remove non-null values
            DataTable uploadFinalTblNull = uploadFinalTbl.Clone();
            foreach (DataRow row in uploadFinalTbl.Rows)
            {
                if (string.IsNullOrEmpty(row["APC_DR"].ToString()))
                {
                    uploadFinalTblNull.ImportRow(row);
                }
            }

            // Get the last column index
            int lastColumnIndex = uploadFinalTblNull.Columns.Count;

            // Add a new column header
            newSheet.Cells[1, lastColumnIndex + 1].Value = "Remarks";
            // Format the new column header
            newSheet.Cells[1, lastColumnIndex + 1].Font.Bold = true;
            newSheet.Cells[1, lastColumnIndex + 1].Interior.ColorIndex = 20;
            newSheet.Cells[1, lastColumnIndex + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            newSheet.Cells[1, lastColumnIndex + 1].VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            // Set the border for the new column header
            Microsoft.Office.Interop.Excel.Borders newBorder = newSheet.Cells[1, lastColumnIndex + 1].Borders;
            newBorder.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            newBorder.Weight = 2.0;

            // Adjust the column width for the new column
            newSheet.Columns[lastColumnIndex + 1].ColumnWidth = 7;

            // Format the new worksheet
            Range newFormatRange2 = newSheet.UsedRange;
            Range newCell = newFormatRange2.Range[newSheet.Cells[1, 1], newSheet.Cells[1, uploadFinalTblNull.Columns.Count]];
            newBorder = newCell.Borders;
            newBorder.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            newBorder.Weight = 2.0;
            newCell.EntireRow.Font.Bold = true;
            newCell.Interior.ColorIndex = 20;

            // Set column width for Column A
            newSheet.Columns[1].ColumnWidth = 13;

            // Set column width for Column C
            newSheet.Columns[3].ColumnWidth = 11;

            // Set column width for Column E
            newSheet.Columns[5].ColumnWidth = 74;

            // Set column width for Column G
            newSheet.Columns[7].ColumnWidth = 11;

            // Set column width for Column H
            newSheet.Columns[8].ColumnWidth = 13;

            // Set column width for Column I
            newSheet.Columns[9].ColumnWidth = 11;


            // Iterate through the rows and columns of uploadFinalTblNull
            int newRow = 2;
            foreach (DataRow row in uploadFinalTblNull.Rows)
            {
                for (int c = 1; c <= uploadFinalTblNull.Columns.Count; c++)
                {
                    newSheet.Cells[newRow, c] = row[c - 1];
                }
                newRow++;
            }

            Range newFormatRange3 = newSheet.UsedRange;
            Range newCell2 = newFormatRange3.Range[newSheet.Cells[1, 1], newSheet.Cells[newSheet.UsedRange.Rows.Count, uploadFinalTblNull.Columns.Count]];
            Microsoft.Office.Interop.Excel.Borders newBorder2 = newCell2.Borders;
            newBorder2.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            newBorder2.Weight = 2.0;
            newCell2.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            newCell2.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            newCell2 = newSheet.Range[newSheet.Cells[2, 2], newSheet.Cells[uploadFinalTblNull.Rows.Count + 1, 2]];
            newCell2.NumberFormat = "h:mm:ss AM/PM";

            wSheet.Columns.AutoFit();
            f2.Close();
            strFileName = saveFileDialog1.FileName;
            wBook.SaveAs(strFileName); //saves the excel file

            //resets the value to their default value
            uploadOne = false;
            uploadTwo = false;
            uploadOneTbl = new DataTable();
            uploadTwoTbl = new DataTable();
            newData = new DataTable();


            //prompts the user that the excel file was created and then open it
            MessageBox.Show("Excel file created!");
            excel.Workbooks.Open(strFileName);
            excel.Visible = true;
            btnDel.BackColor = System.Drawing.Color.SteelBlue;
            btnAxmrUpl.BackColor = System.Drawing.Color.SteelBlue;
            btnDel.Enabled = true;
            btnAxmrUpl.Enabled = true;

            try
            {
                xlApp.Quit();
            }
            catch (Exception ex)
            {

            }
            ReleaseObject(xlWorksheet);
            ReleaseObject(xlWorkbook);
            ReleaseObject(oExcel2);
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
        private void connectdata()
        {
            con.ConnectionString = connectionString; //declaring connection from host
            con.Open(); //to open database connection
        }

        private void btnAxmrUpl_Click(object sender, EventArgs e)
        {
            if (!uploadOne) // checks if the user has imported the first excel file
            {
                MessageBox.Show("Please import the first file!");
                return;
            }

            if (uploadTwo) // checks if the user has already clicked the button
            {
                return;
            }

            //displays a dialog where user can select the file to open
            string strFileName;

            openFD = new OpenFileDialog();
            openFD.InitialDirectory = "'C:\'";
            openFD.Filter = "Excel Office | *.xlsx; *.xls";
            openFD.Title = "Choose a File";
            openFD.FilterIndex = 2;
            openFD.RestoreDirectory = true;

            if (openFD.ShowDialog().Equals(DialogResult.OK))
            {
                strFileName = openFD.FileName;
                f1.ptkFileName.Text = openFD.FileName;
                if (!string.IsNullOrEmpty(FileSystem.Dir(openFD.FileName)))
                {
                    if (!openFD.SafeFileName.ToString().ToUpper().Contains("AXMR620"))
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
                                uploadTwoSave(strFileName, openFD.SafeFileName);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Please close the excel application then try importing again!");
                                f2.Close();

                            }
                            f2.label2.Text = "IMPORTING COMPLETE!";
                            f2.Close();
                        }
                        else
                        {
                            MessageBox.Show("FILE NOT FOUND!");
                        }
                    }
                }
            }
        }

        private void uploadTwoSave(string myfiledirect, string filename)
        {
            uploadTwoTbl = new DataTable(); //creates a new instance of the datatable
            Form2 f2 = new Form2();
            oExcel2 = new Excel.Application();
            oExcel2.DisplayAlerts = false;
            xlWorkbook = oExcel2.Workbooks.Open(myfiledirect, XlFileAccess.xlReadOnly); ;
            xlWorksheet = xlWorkbook.Worksheets[1];
            DataColumn Column = new DataColumn();
            int lRow = xlWorksheet.Range["M" + xlWorksheet.Rows.Count.ToString()].End[XlDirection.xlUp].Row;
            xlRange = xlWorksheet.Range["A7:M" + lRow.ToString()];
            object[,] data = (object[,])xlRange.Value;
            string apcDRexcel = "";


            for (int cCnt = 1; cCnt <= xlRange.Columns.Count; cCnt++)
            {
                f2.Show();
                f2.progressBar1.Maximum = xlRange.Columns.Count;
                f2.progressBar1.Value = cCnt;
                f2.label2.Text = "Importing Data . . .";
                f2.Refresh();

                // Creates new Column in DataTable
                Column = new DataColumn();
                Column.DataType = typeof(string);
                Column.ColumnName = cCnt.ToString();
                uploadTwoTbl.Columns.Add(Column);

                // Creates new row for Data Table
                for (int rCnt = 2; rCnt <= xlRange.Rows.Count; rCnt++)
                {
                    string rCnt1 = Convert.ToString(rCnt);
                    string cCnt1 = Convert.ToString(cCnt);
                    string CellVal = string.Empty;
                    CellVal = Convert.ToString(data[rCnt, cCnt]);

                    DataRow Row;

                    //Adds new row to the DataTable
                    if (cCnt == 1)
                    {
                        Row = uploadTwoTbl.NewRow();
                        Row[Column.ColumnName.ToString()] = CellVal;
                        uploadTwoTbl.Rows.Add(Row);
                    }
                    else
                    {
                        if (cCnt == 2)
                        {
                            if (string.IsNullOrEmpty(CellVal))
                            {
                                CellVal = apcDRexcel;
                                CellVal = CellVal.Replace("PC", "APC").Replace("AAPC", "APC");
                            }
                            else
                            {
                                CellVal = CellVal.Replace("PC", "APC").Replace("AAPC", "APC");
                                if (!CellVal.Contains("CDR"))
                                {
                                    apcDRexcel = CellVal;
                                }
                            }
                        }
                        Row = uploadTwoTbl.Rows[rCnt - 2];
                        Row[Column.ColumnName.ToString()] = CellVal;
                    }
                }
            }

            DataRow[] prt = uploadTwoTbl.Select("[13] is null or [13] LIKE '%Shipped Qty%'");

            foreach (DataRow prt2 in prt)
            {
                prt2.Delete();
            }

            uploadTwoTbl.AcceptChanges();

            xlWorkbook.Close();
            f2.Close();
            uploadTwo = true;
            MessageBox.Show("import Successfully", "Success");
            btnAxmrUpl.BackColor = System.Drawing.Color.Navy;
            btnAxmrUpl.Enabled = false;
            try
            {
                oExcel2.Quit();
            }
            catch (Exception ex)
            {
                // Handle the exception
            }

            ReleaseObject(lRow);
            ReleaseObject(xlWorksheet);
            ReleaseObject(xlWorkbook);
            ReleaseObject(oExcel2);
        }
    }
}
