using CrystalDecisions.ReportAppServer.ClientDoc;
using CrystalDecisions.ReportAppServer.DataDefModel;
using System;
using System.IO;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Threading;
using Snowflake.Data.Client;
using System.Data;
using CrystalDecisions.ReportAppServer.Controllers;
using CrystalDecisions.Shared;

namespace CR_COMMAND_EXT_snow
{
    public partial class ExtractCMD : Form
    {
        public ExtractCMD()
        {
            InitializeComponent();
        }

        //Button to Navigate to browse to Directory path folder.
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        // Create a progress bar.
        private Task ProcessData(List<string> list, IProgress<ProgressReport> progressBar1)
        {
            int index = 1;
            int totalProcess = list.Count;
            var progressReport = new ProgressReport();
            return Task.Run(() =>
            {
                for (int i = 0; i < totalProcess; i++)
                {
                    progressReport.PercentComplete = index++ * 100 / totalProcess;
                    progressBar1.Report(progressReport);
                    Thread.Sleep(10);//used to simulate lenght of operation
                }
            });
        }
        //Button to Process the files in the Directory
        private async void button2_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            for (int i = 0; i < 1000; i++)
                list.Add(i.ToString());
            label2.Text = "Working...";
            var progress = new Progress<ProgressReport>();
            progress.ProgressChanged += (o, report) =>
            {
                label2.Text = string.Format("Processing...{0}%", report.PercentComplete);
                progressBar1.Value = report.PercentComplete;
                progressBar1.Update();
            };

            //For each file in Directory find .rpt. Get Files with the .rpt extension -> Load the file into Crystal Reports Engine
            CrystalDecisions.CrystalReports.Engine.ReportDocument doc = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            string filename; // Add this new variable to store the filename.
            string Subfilename;

            foreach (string files in Directory.GetFiles(textBox1.Text, "*.rpt", SearchOption.AllDirectories))
            {
                try
                {
                    Console.WriteLine(String.Format("Processing {0}...", files));
                    doc.Load(files);
                }
                catch (CrystalReportsException ex)
                {
                    //Save the exception to a Log file
                    string error = ex.ToString();
                    string exfilepath = @"3 - ErrorLog\\log.txt";

                    //string FilePath = Path.Combine(savepath + "\\" + filename) + ".sql";
                    //Append the error details to the text file
                    File.AppendAllText(exfilepath, error);

                }
                filename = Path.GetFileName(files);
                Subfilename = Path.GetFileName(files);

                {
                    {
                        try
                        {
                            {
                                //For each file with .rpt find the file with ClassName = CrystalReports. CommandTable
                                //MAIN REPORT = COMMAND - no SUBREPORT  
                                foreach (dynamic mntable in doc.ReportClientDocument.DatabaseController.Database.Tables)
                                {
                                    if (mntable.ClassName == "CrystalReports.CommandTable")
                                    {
                                        ISCDReportClientDocument mnrptClientDoc;
                                        mnrptClientDoc = doc.ReportClientDocument;

                                        foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table mntmpTbl in mnrptClientDoc.DatabaseController.Database.Tables)
                                            if (mntmpTbl is CommandTable)
                                            {
                                                CommandTable mncmdTbl = (CommandTable)mntmpTbl;
                                                string commandsql = mncmdTbl.CommandText;
                                            }
                                        //For each file with .rpt find the file with ClassName = CrystalReports. CommandTable
                                        //MAIN REPORT = COMMAND - SUBREPORT = COMMAND               
                                        //For each file with .rpt find the file with ClassName = CrystalReports. CommandTable

                                        foreach (dynamic Subtable in doc.ReportClientDocument.DatabaseController.Database.Tables)
                                        {
                                            if (!doc.IsSubreport && Subtable.ClassName == "CrystalReports.CommandTable")
                                            {
                                                ISCDReportClientDocument rptSubClientDoc;
                                                rptSubClientDoc = doc.ReportClientDocument;

                                                foreach (dynamic subName in rptSubClientDoc.SubreportController.GetSubreportNames())
                                                    foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table tmpSubTbl in rptSubClientDoc.DatabaseController.Database.Tables)

                                                        if (tmpSubTbl is CommandTable)
                                                        {
                                                            CommandTable cmdSubTbl = (CommandTable)tmpSubTbl;
                                                            //string commandSubSql = cmdSubTbl.CommandText;
                                                        }

                                                //Set variables
                                                string srcfilePath = doc.FileName;
                                                string commandSql = mntable.CommandText;
                                                string commandSubSql = Subtable.CommandText;


                                                // Create a SqlConnection object and connect to the database.
                                                using (IDbConnection conn = new SnowflakeDbConnection())
                                                {
                                                    conn.ConnectionString = "account=psjh_prod.west-us-2.azure;user=debra.quarles@providence.org;authenticator=externalbrowser;warehouse=ACOE_WH_S;db=ACOE_SANDBOX;schema=P418549";

                                                    //Open the SqlConnection object
                                                    conn.Open();

                                                    // Create a SqlCommand object and specify the SQL INSERT statement.
                                                    IDbCommand cmd = conn.CreateCommand();
                                                    cmd.CommandText = "INSERT INTO P418549.CR_Command_Extract(CRRPT_NM,CRRPT_PATH,CRRPT_SQL_CD,CRRPT_SUBSQL_CD) VALUES (:filename,:srcfilePath,:commandSql,:commandSubSql)";

                                                    // Set the parameters of the SqlCommand object with the values from the report document.
                                                    var p1 = cmd.CreateParameter();
                                                    p1.ParameterName = "filename";
                                                    p1.Value = filename;
                                                    p1.DbType = DbType.String;
                                                    cmd.Parameters.Add(p1);

                                                    var p2 = cmd.CreateParameter();
                                                    p2.ParameterName = "srcfilePath";
                                                    p2.Value = srcfilePath;
                                                    p2.DbType = DbType.String;
                                                    cmd.Parameters.Add(p2);

                                                    var p3 = cmd.CreateParameter();
                                                    p3.ParameterName = "commandSql";
                                                    p3.Value = commandSql;
                                                    p3.DbType = DbType.String;
                                                    cmd.Parameters.Add(p3);

                                                    var p4 = cmd.CreateParameter();
                                                    p4.ParameterName = "commandSubSql";
                                                    p4.Value = commandSubSql;
                                                    p4.DbType = DbType.String;
                                                    cmd.Parameters.Add(p4);

                                                    var count = cmd.ExecuteNonQuery();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        catch (CrystalReportsException ex)
                        {
                            //Save the exception to a Log file
                            string error = ex.ToString();
                            string exfilepath = @"3 - ErrorLog\\log.txt";

                            //string FilePath = Path.Combine(savepath + "\\" + filename) + ".sql";
                            //Append the error details to the text file
                            File.AppendAllText(exfilepath, error);

                        }
                    }
                }

                try
                {
                    CrystalDecisions.CrystalReports.Engine.ReportObjects crReportObjects;
                    CrystalDecisions.CrystalReports.Engine.SubreportObject crSubreportObject;
                    CrystalDecisions.CrystalReports.Engine.ReportDocument crSubreportDocument;

                    ISCDReportClientDocument rptSubsubClientDoc;
                    rptSubsubClientDoc = doc.ReportClientDocument;
                    try
                    {
                        //set the crSections object to the current report's sections
                        CrystalDecisions.CrystalReports.Engine.Sections crSections = doc.ReportDefinition.Sections;

                        //loop through all the sections to find all the report objects
                        foreach (CrystalDecisions.CrystalReports.Engine.Section crSection in crSections)
                        {
                            crReportObjects = crSection.ReportObjects;
                            //loop through all the report objects to find all the subreports
                            foreach (CrystalDecisions.CrystalReports.Engine.ReportObject crReportObject in crReportObjects)
                            {
                                if (crReportObject.Kind == ReportObjectKind.SubreportObject)
                                {
                                    //you will need to typecast the reportobject to a subreport 
                                    //object once you find it
                                    crSubreportObject = (CrystalDecisions.CrystalReports.Engine.SubreportObject)crReportObject;
                                    crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);
                                    SubreportClientDocument subRCD = rptSubsubClientDoc.SubreportController.GetSubreport(crSubreportObject.SubreportName);
                                    string mysubname = crSubreportObject.SubreportName.ToString();

                                    if (subRCD.DatabaseController.Database.Tables.Count != 0)
                                    {
                                        foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table crTable in subRCD.DatabaseController.Database.Tables)
                                        {
                                            try
                                            {
                                                // Subreport is using a Command so use RAS to get the SQL
                                                if (((dynamic)crTable.Name) == "Command")
                                                {
                                                    CrystalDecisions.ReportAppServer.Controllers.DatabaseController databaseController = subRCD.DatabaseController;
                                                    CommandTable SubsubTable = (CommandTable)databaseController.Database.Tables[0];
                                                    ((dynamic)SubsubTable).CommandText.ToString();

                                                    //Set variables        
                                                    string commandSubsubSql = SubsubTable.CommandText;
                                                    string SubsrcfilePath = doc.FileName;



                                                    // Create a SqlConnection object and connect to the database.
                                                    using (IDbConnection conn = new SnowflakeDbConnection())
                                                    {
                                                        conn.ConnectionString = "account=psjh_prod.west-us-2.azure;user=debra.quarles@providence.org;authenticator=externalbrowser;warehouse=ACOE_WH_S;db=ACOE_SANDBOX;schema=P418549";

                                                        //Open the SqlConnection object
                                                        conn.Open();
                                                        // Create a SqlCommand object and specify the SQL INSERT statement.
                                                        IDbCommand cmdsub = conn.CreateCommand();
                                                        cmdsub.CommandText = "INSERT INTO P418549.CR_Command_Extract(CRRPT_NM,CRRPT_PATH,CRRPT_SUBSUBSQL_CD) VALUES (:filename,:srcfilePath,:commandSubsubSql)";

                                                        var p5 = cmdsub.CreateParameter();
                                                        p5.ParameterName = "filename";
                                                        p5.Value = Subfilename + "-sub";
                                                        p5.DbType = DbType.String;
                                                        cmdsub.Parameters.Add(p5);

                                                        var p6 = cmdsub.CreateParameter();
                                                        p6.ParameterName = "srcfilePath";
                                                        p6.Value = SubsrcfilePath;
                                                        p6.DbType = DbType.String;
                                                        cmdsub.Parameters.Add(p6);

                                                        var p7 = cmdsub.CreateParameter();
                                                        p7.ParameterName = "commandSubsubSql";
                                                        p7.Value = commandSubsubSql;
                                                        p7.DbType = DbType.String;
                                                        cmdsub.Parameters.Add(p7);

                                                        // Execute the SqlCommand object.
                                                        var countsub = cmdsub.ExecuteNonQuery();

                                                        // Close the SqlConnection object.
                                                        conn.Close();
                                                    }
                                                }
                                            }
                                            catch (InvalidCastException ex)
                                            {
                                                //Save the exception to a Log file
                                                string error = ex.ToString();
                                                string exfilepath = @"3 - ErrorLog\\log.txt";

                                                //string FilePath = Path.Combine(savepath + "\\" + filename) + ".sql";
                                                //Append the error details to the text file
                                                File.AppendAllText(exfilepath, error);

                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    catch (CrystalReportsException ex)
                    {
                        //Save the exception to a Log file
                        string error = ex.ToString();
                        string exfilepath = @"3 - ErrorLog\\log.txt";

                        //string FilePath = Path.Combine(savepath + "\\" + filename) + ".sql";
                        //Append the error details to the text file
                        File.AppendAllText(exfilepath, error);
                    }
                }
                catch (CrystalDecisions.CrystalReports.Engine.LoadSaveReportException ex)
                {
                    //Save the exception to a Log file
                    string error = ex.ToString();
                    string exfilepath = @"3 - ErrorLog\\log.txt";

                    //string FilePath = Path.Combine(savepath + "\\" + filename) + ".sql";
                    //Append the error details to the text file
                    File.AppendAllText(exfilepath, error);
                }

            }
                                    
                //Begin the Green progress bar - When ALL files in the chosen folder are processed show "Done"
                await ProcessData(list, progress);
                label2.Text = "Done !";
            }

        //Exit Button to exit the form and close
        private void button3_Click(object sender, EventArgs e)
        {
            {
                this.Close();
            }
        }
        //Label to instruct to Navigate to Directory Path
            private void label1_Click(object sender, EventArgs e)
        {

        }
        //Textbox in which to display the source folder name
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        //Label to Display "Processing progress
        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
