using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Data.OleDb;
using System.Globalization;
using System.Diagnostics;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data;

namespace WorkReport
{
    public partial class Form1 : Form
    {
        public string FromDate = string.Empty;
        public string ToDate =   string.Empty;
        
        public Form1()
        {
            InitializeComponent();
            progrssbar.Visible = false;
            txtfile.Enabled = false;
            lblprgrss.Visible = false;
            lblmsg.Visible = false;
            txtmsg.Visible = false;
            btnrunreport.Enabled = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
         
            frmdate.Format = DateTimePickerFormat.Short;
            frmdate.Value = DateTime.Today;

            todate.Format = DateTimePickerFormat.Short;
            todate.Value = DateTime.Today;
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog odf = new OpenFileDialog();
            odf.Filter = "Excel|*.xls|Xlsx|*.xlsx";
            odf.Multiselect = false;
            if (odf.ShowDialog() == DialogResult.OK)
            {
                txtfile.Text = odf.FileName;
                btnrunreport.Enabled = true;
            }
        }

        private void btnrunreport_Click(object sender, EventArgs e)
        {
            progrssbar.Visible = true;
            lblmsg.Visible = true;
            txtmsg.Visible = true;
            lblmsg.Text = "Generating...";
            genratereport.Enabled = false;
            myBGWorker.RunWorkerAsync();
        }

        public DataSet GetExcelData(string servicefilepath)
        {
            DataSet ds = new DataSet();
       
            string strConn = string.Empty;
            if (Path.GetExtension(servicefilepath) == ".xls")
            {
                strConn = "provider=Microsoft.Jet.OLEDB.4.0;" + @"data source=" + servicefilepath + ";" + "Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
            }
            if (Path.GetExtension(servicefilepath) == ".xlsx")
            {
                strConn = "provider=Microsoft.ACE.OLEDB.12.0;" + @"data source=" + servicefilepath + ";" + "Extended Properties='Excel 12.0;HDR=NO;IMEX=1';";
            }

            try
            {
                using (OleDbConnection conn = new OleDbConnection(strConn))
                {
                    conn.Open();

                    DataTable Sheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    string SheetName = "";

                    for (int i = 0; i < Sheets.Rows.Count; i++)
                    {
                        string SheetNameInner = Convert.ToString(Sheets.Rows[i]["TABLE_NAME"]);

                        if (SheetNameInner.ToLower().Equals("'page 1$'"))
                        {
                            SheetName = SheetNameInner;
                        }
                    }

                    string command = string.Format("SELECT * FROM [{0}]", SheetName);

                    OleDbCommand cmd = new OleDbCommand(command, conn);

                    // Create new OleDbDataAdapter

                    oleda.SelectCommand = cmd;

                    // Fill the DataSet from the data extracted from the worksheet.

                    // Fill the DataSet from the data extracted from the worksheet.
                    oleda.Fill(ds);
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }

            return ds;
        }

        public string GetTechnologyandManager(string name, string groupname, DataTable dtmasterdata)
        {
            var rows = dtmasterdata.AsEnumerable().Where(a => Convert.ToString(a[3]) == groupname && Convert.ToString(a[1]) == name);

            if (rows.Any())
            {
                var dt = rows.CopyToDataTable();
               if(dt.Rows.Count>0)

                return dt.Rows[0][2].ToString() + ":" + dt.Rows[0][4].ToString();
                
               else
                   return string.Empty;

            }
            else
            {
                return string.Empty;
            }
        }

        private void myBGWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                string servicenowfilepath = Path.GetDirectoryName(Application.ExecutablePath).Replace(@"bin\debug\", string.Empty) + "\\ServiceNowData-Input\\" + Path.GetFileNameWithoutExtension(txtfile.Text) + DateTime.Now.ToString("dd-MMM-yyyy").ToString()+"-Auto" + Path.GetExtension(txtfile.Text);
                string masterdatafilepath = Path.GetDirectoryName(Application.ExecutablePath).Replace(@"bin\debug\", string.Empty) + "\\MasterData\\MasterData.xlsx";
                string statusfilepath = Path.GetDirectoryName(Application.ExecutablePath).Replace(@"bin\debug\", string.Empty) + "\\MasterData\\ServiceNowStatus.xlsx";

                File.Copy(txtfile.Text, servicenowfilepath, true);

                myBGWorker.ReportProgress(2);

                DataTable dtservicenow = new DataTable();
                dtservicenow = GetExcelData(servicenowfilepath).Tables[0];

                myBGWorker.ReportProgress(5);

                DataTable dtmasterdata = new DataTable();
                dtmasterdata = GetExcelData(masterdatafilepath).Tables[0];

                DataTable dtservicenowstatus = new DataTable();
                dtservicenowstatus = GetExcelData(statusfilepath).Tables[0];

                myBGWorker.ReportProgress(10);

                dtservicenow.Columns.Add("Technology", typeof(string));
                dtservicenow.Columns.Add("Status", typeof(string));
                dtservicenow.Columns.Add("Ticket Type", typeof(string));
                dtservicenow.Columns.Add("FromDate", typeof(string));
                dtservicenow.Columns.Add("ToDate", typeof(string));
                dtservicenow.Columns.Add("Month", typeof(string));
                dtservicenow.Columns.Add("IsCurrentWeek", typeof(string));

                foreach (DataRow dr in dtservicenow.Rows)
                {
                    if (dr[0].ToString() == "Task")
                    {
                        dr[13] = "Technology";
                        dr[14] = "Status";
                        dr[15] = "Ticket Type";
                        dr[16] = "FromDate";
                        dr[17] = "ToDate";
                        dr[18] = "Month";
                        dr[19] = "IsCurrentWeek";
                    }
                    else
                    {
                        string technmnager = GetTechnologyandManager(dr[5].ToString(), dr[4].ToString(), dtmasterdata);
                        string[] details = technmnager.Split(':');
                        dr[13] = details[0].ToString();

                        var rows = dtservicenowstatus.AsEnumerable().Where(a => Convert.ToString(a[1]).ToLower() == dr[7].ToString().ToLower());
                        if (rows.Any())
                        {
                            var filtereddt = rows.CopyToDataTable();
                            foreach (DataRow row in filtereddt.Rows)
                            {
                                dr[14] = row[2].ToString();
                            }

                        }
                        else
                        {
                            dr[14] = "NA";
                        }

                        if (dr[2].ToString().Contains("4"))
                            dr[15] = "P4";
                        if (dr[2].ToString().Contains("3"))
                            dr[15] = "P3";
                        if (dr[2].ToString().Contains("2"))
                            dr[15] = "P2";
                        if (dr[2].ToString().Contains("1"))
                            dr[15] = "P1";
                        
                        DateTime date = DateTime.ParseExact(todate.Text.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        dr[16]= frmdate.Text.ToString();
                        dr[17] = todate.Text.ToString();
                        dr[18]=date.ToString("MMMM ,yyyy");
                        dr[19]=isdisplayweek.Checked?"Yes":"No";
                    }

                }

                myBGWorker.ReportProgress(20);

                dtservicenow.Rows.Remove(dtservicenow.Rows[0]);
                string taskcolumnname = dtservicenow.Columns[0].ColumnName;

                dtservicenow.DefaultView.Sort = dtservicenow.Columns[0].ColumnName + " ASC," + dtservicenow.Columns[11].ColumnName + " DESC";
                dtservicenow = dtservicenow.DefaultView.ToTable();
                dtservicenow = dtservicenow.AsEnumerable().GroupBy(x => x.Field<string>(taskcolumnname)).Select(y => y.First()).CopyToDataTable();

                //using (var command = new SqlCommand("InsertTable") { CommandType = CommandType.StoredProcedure })
                //{
                //    var dt = new DataTable(); //create your own data table
                //    command.Parameters.Add(new SqlParameter("@myTableType", dt));
                //    SqlHelper.exe(command);
                //}
                CreatingDataforReport(ref dtservicenow, ref dtmasterdata);
            }
            catch (Exception ex)
            {

                throw;
                //MessageBox.Show("There must be some issue with the source file. The error is- "+ex.Message.ToString());
            }
        }

        void myBGWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progrssbar.Value = e.ProgressPercentage;
            if (e.ProgressPercentage == 100)
            {
                //System.Threading.Thread.Sleep(100);
                txtmsg.Text = string.Empty;
                txtmsg.Text = "The file has been successfully generated.";
                lblmsg.Text = "Click here to go to the file..";
                lblmsg.Links.Clear();
                LinkLabel.Link link = new LinkLabel.Link();
                link.LinkData = Path.GetDirectoryName(Application.ExecutablePath).Replace(@"bin\debug\", string.Empty) + "\\Report-Output\\";
                lblmsg.Links.Add(link);
            }
        }

        void myBGWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                txtmsg.Text = string.Empty;
                txtmsg.Text = "There must be some issue with the source file.\nThe error is- " + e.Error.ToString();
                lblmsg.Text = string.Empty;
                lblmsg.Visible = false;
            }
            else {
                genratereport.Enabled = true;
            }
        }

        public void CreatingDataforReport(ref DataTable dtservicenow, ref DataTable dtmasterdata)
        {
            DataTable dttechnology = new DataTable();
            dttechnology = dtmasterdata.DefaultView.ToTable(true, dtmasterdata.Columns[2].ColumnName);
            dttechnology.Rows.Remove(dttechnology.Rows[0]);

            DataTable dtreport = new DataTable();
            dtreport.Columns.Add("Ticket Count", typeof(int));
            dtreport.Columns.Add("Ticket Type", typeof(string));
            dtreport.Columns.Add("Stage", typeof(string));
            dtreport.Columns.Add("Technology", typeof(string));
            dtreport.Columns.Add("Month", typeof(string));
            dtreport.Columns.Add("From Date", typeof(string));
            dtreport.Columns.Add("To Date", typeof(string));
            dtreport.Columns.Add("Display Data For Current Week", typeof(string));
            int progresscount = 20;
            foreach (DataRow dr in dttechnology.Rows)
            {
                AppendRowsforReport(dr[0].ToString(), ref dtreport, dtservicenow);
                progresscount += 10;
                myBGWorker.ReportProgress(progresscount);
            }

            RenderReport(ref dtreport);
        }

        public void AppendRowsforReport(string tech, ref DataTable dtreport, DataTable dtservicenow)
        {
            for (int i = 1; i <= 12; i++)
            {
                DataRow dr = dtreport.NewRow();
                if (i <= 4)
                {
                    dr[1] = "P" + i;
                    dr[2] = "Resolved";
                    dr[3] = tech;
                    DateTime date = DateTime.ParseExact(todate.Text.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    dr[4] = date.ToString("MMMM ,yyyy");
                    dr[5] = frmdate.Text.ToString();
                    dr[6] = todate.Text.ToString();
                    FromDate = DateTime.ParseExact(dr[5].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("ddMMMyyyy");
                    ToDate = DateTime.ParseExact(dr[6].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("ddMMMyyyy");
                    dr[7]= isdisplayweek.Checked?"Yes":"No";
                    var rows = dtservicenow.AsEnumerable().Where(a => Convert.ToString(a[13]) == tech && Convert.ToString(a[14]) == "Resolved" && Convert.ToString(a[15]) == dr[1].ToString());
                    if (rows.Any())
                    {
                        var dt = rows.CopyToDataTable();
                        dr[0] = dt.Rows.Count.ToString();

                    }
                    else
                    {
                        dr[0] = "0";
                    }
                    dtreport.Rows.Add(dr);
                }
                if (i > 4 && i <= 8)
                {
                    int count = i - 4;
                    dr[1] = "P" + count;
                    dr[2] = "In Progress";
                    dr[3] = tech;
                    DateTime date = DateTime.ParseExact(todate.Text.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    dr[4] = date.ToString("MMMM ,yyyy");
                    dr[5] = frmdate.Text.ToString();
                    dr[6] = todate.Text.ToString();
                    FromDate = DateTime.ParseExact(dr[5].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("ddMMMyyyy");
                    ToDate = DateTime.ParseExact(dr[6].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("ddMMMyyyy");
                    dr[7] = isdisplayweek.Checked ? "Yes" : "No";
                    var rows = dtservicenow.AsEnumerable().Where(a => Convert.ToString(a[13]) == tech && Convert.ToString(a[14]) == "In Progress" && Convert.ToString(a[15]) == dr[1].ToString());
                    if (rows.Any())
                    {
                        var dt = rows.CopyToDataTable();
                        dr[0] = dt.Rows.Count.ToString();

                    }
                    else
                    {
                        dr[0] = "0";
                    }

                    dtreport.Rows.Add(dr);
                }

                if (i > 8 && i <= 12)
                {
                    int count = i - 8;
                    dr[1] = "P" + count;
                    dr[2] = "Breached";
                    dr[3] = tech;
                    DateTime date = DateTime.ParseExact(todate.Text.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    dr[4] = date.ToString("MMMM ,yyyy");
                    dr[5] = frmdate.Text.ToString();
                    dr[6] = todate.Text.ToString();
                    FromDate = DateTime.ParseExact(dr[5].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("ddMMMyyyy");
                    ToDate = DateTime.ParseExact(dr[6].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("ddMMMyyyy");
                    dr[7] = isdisplayweek.Checked ? "Yes" : "No";
                    var rows = dtservicenow.AsEnumerable().Where(a => Convert.ToString(a[13]) == tech && Convert.ToString(a[14]) == "Breached" && Convert.ToString(a[15]) == dr[1].ToString());
                    if (rows.Any())
                    {
                        var dt = rows.CopyToDataTable();
                        dr[0] = dt.Rows.Count.ToString();

                    }
                    else
                    {
                        dr[0] = "0";
                    }

                    dtreport.Rows.Add(dr);
                }
            }

            for (int i = 1; i <= 3; i++)
            {
                DataRow dr = dtreport.NewRow();
                dr[0] = "0";
                dr[1] = "Enhancements";
                switch (i)
                {
                    case 1: dr[2] = "Resolved";
                        break;
                    case 2: dr[2] = "In Progress";
                        break;
                    case 3: dr[2] = "Breached";
                        break;
                }
                dr[3] = tech;
                DateTime date = DateTime.ParseExact(todate.Text.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                dr[4] = date.ToString("MMMM ,yyyy");
                dr[5] = frmdate.Text.ToString();
                dr[6] = todate.Text.ToString();
                FromDate = DateTime.ParseExact(dr[5].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("ddMMMyyyy");
                ToDate = DateTime.ParseExact(dr[6].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("ddMMMyyyy");
                dr[7] = isdisplayweek.Checked ? "Yes" : "No";
                dtreport.Rows.Add(dr);
            }

        }

        public void RenderReport(ref DataTable dtreport)
        {

            dtreport.DefaultView.Sort = dtreport.Columns[3].ColumnName + " ASC," + dtreport.Columns[1].ColumnName + " DESC," + dtreport.Columns[2].ColumnName + " DESC";
            dtreport = dtreport.DefaultView.ToTable();

            string reportpath = Path.GetDirectoryName(Application.ExecutablePath).Replace(@"bin\debug\", string.Empty) + "\\Report-Output\\" + "Resolution-" + FromDate +" - " + ToDate + "_Output-Auto" +".xlsx";

            ExcelPackage package = new ExcelPackage();
            ExcelWorksheet ws;
            ws = package.Workbook.Worksheets.Add("Ticket BreakUp");
            ws.View.ShowGridLines = true;
            ws.Cells["A1"].LoadFromDataTable(dtreport, true);

            ws.Cells["A1:H1"].Style.Font.Bold = true;
            ws.Cells["A1:H1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Color myColor = ColorTranslator.FromHtml("#FFFFCC");
            ws.Cells["A1:H1"].Style.Fill.BackgroundColor.SetColor(myColor);

            ws.Column(1).Width = 20;
            ws.Column(2).Width = 20;
            ws.Column(3).Width = 20;
            ws.Column(4).Width = 20;
            ws.Column(5).Width = 20;
            ws.Column(6).Width = 20;
            ws.Column(7).Width = 20;
            ws.Column(8).Width = 30;
            int dtrowscount = dtreport.Rows.Count + 1;

            ExcelRange range = ws.Cells["A1:H" + dtrowscount.ToString()];

            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.White);
            range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


            ws = package.Workbook.Worksheets.Add("Track & Status");
            ws.View.ShowGridLines = true;
            

            ws.Cells["A1:E1"].Style.Font.Bold = true;
            ws.Cells["A1:E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Color myColor1 = ColorTranslator.FromHtml("#FFFFCC");
            ws.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(myColor1);
            ws.Cells[1, 1].Value = "S.No.";
            ws.Cells[1, 2].Value = "Workstream";
            ws.Cells[1, 3].Value = "Category";
            ws.Cells[1, 4].Value = "Stage";
            ws.Cells[1, 5].Value = "Current Status";
            ws.Column(1).Width = 10;
            ws.Column(2).Width = 30;
            ws.Column(3).Width = 20;
            ws.Column(4).Width = 20;
            ws.Column(5).Width = 100;
            ws.Column(5).Style.WrapText = true;
            
            ExcelRange range1 = ws.Cells["A1:E100"];

            range1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range1.Style.Fill.BackgroundColor.SetColor(Color.White);
            range1.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            range1.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range1.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range1.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;



            FileInfo newFile = new FileInfo(reportpath);
            package.SaveAs(newFile);

            myBGWorker.ReportProgress(100);
        }

        private void lblmsg_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(e.Link.LinkData as string);
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

    }
}
