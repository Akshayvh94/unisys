using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using unisys.Models;
using static unisys.Models.WorkItemFetchResponse;

namespace unisys
{
    internal class Program
    {
        public static List<string> iterationList = new List<string>();
        public static List<string> UserList = new List<string>();
        public static string logFile = string.Empty;
        public static string path = string.Empty;
        public static string sheetName = string.Empty;
        private static void Main(string[] args)
        {
            try
            {
                path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
                path = path.Replace("file:\\", "");
                Console.WriteLine("Consolidated report from Azure DevOps and Time Sheet");
                //Console.WriteLine(path);
                Console.WriteLine();
                Console.WriteLine();
                if (!Directory.Exists(path + "\\Logs"))
                {
                    Directory.CreateDirectory(path + "\\Logs");
                }
                logFile = path + "\\Logs\\GetReport-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
                Console.WriteLine("Enter Organization Name");
                string OrgName = Console.ReadLine();

                Console.WriteLine("Enter Project Name");
                string projectName = Console.ReadLine();

                Console.WriteLine("Enter PAT");
                string pat = Console.ReadLine();

                Console.WriteLine("Initializing...");
                string basePat = string.Empty;

                if (string.IsNullOrEmpty(pat) || string.IsNullOrEmpty(OrgName) || string.IsNullOrEmpty(projectName))
                {
                    Console.WriteLine("Please enter all the details");
                    WriteFileToDisk("", "Please enter all the details");
                    Console.ReadLine();
                    return;
                }
                else
                {
                    basePat = Convert.ToBase64String(Encoding.ASCII.GetBytes(string.Format("{0}:{1}", "", pat)));
                }

                string Filepath = string.Empty;

                do
                {
                    Console.WriteLine("Enter the file path : Should end with .xlsx or .xls");
                    Filepath = Console.ReadLine();
                } while (!File.Exists(Filepath));
                do
                {
                    Console.WriteLine("Enter Sheet Name");
                    sheetName = Console.ReadLine();
                } while (string.IsNullOrEmpty(sheetName));
                //string pat = "rahfqiseaujjsyj5qv7o25yghyicy3cgtrqu3cse267be52lu2na";
                //string pat = "vmi4rbndghwzyea7camsnkj5jmb6u4za23vyt7xahg2texklfiwa"; // PAT from your ORG


                UrlParameters parameters = new UrlParameters();
                parameters.Project = projectName;// Project name 
                parameters.Account = OrgName; // Organization Name
                parameters.UriString = "https://dev.azure.com/" + parameters.Account + "/" + parameters.Project;
                parameters.PatBase = basePat;
                parameters.Pat = pat;


                WriteFileToDisk("Entered Details", "\t Organization: " + OrgName + "\t Project Name: " + projectName);

                List<WorkItemWithIteration> workItemIteration_cWork = new List<WorkItemWithIteration>();
                Console.WriteLine($"Fetching work item details from https://dev.azure.com/{parameters.Project}");

                WriteFileToDisk("Entered Details", $"Fetching work item details from https://dev.azure.com/{parameters.Project}");

                List<WorkItems> workItemsList = GetWorkItemsfromSource("Task", parameters);
                Console.WriteLine("Taking summation of completed work");
                WriteFileToDisk("", "Taking summation of completed work");

                if (workItemsList.Count > 0)
                {
                    List<string> users = new List<string>();
                    foreach (var ItemList in workItemsList)
                    {
                        foreach (var workItem in ItemList.value)
                        {
                            var element = workItemIteration_cWork.Find(e => e.UserName == workItem.fields.SystemAssignedTo.uniqueName && e.IterationPath == workItem.fields.SystemIterationPath);
                            if (element != null)
                            {
                                element.Value = element.Value + workItem.fields.MicrosoftVSTSSchedulingCompletedWork;
                            }
                            else
                            {
                                workItemIteration_cWork.Add(new WorkItemWithIteration
                                {
                                    UserName = workItem.fields.SystemAssignedTo.uniqueName,
                                    IterationPath = workItem.fields.SystemIterationPath,
                                    Value = workItem.fields.MicrosoftVSTSSchedulingCompletedWork
                                });
                            }
                        }
                    }
                }
                else
                {
                    WriteFileToDisk("", "Couldn't get the work items. Please check the Organization details and try again");
                    Console.WriteLine("Couldn't get the work items. Please check the Organization details and try again");
                    return;
                }
                DataTable dtworkItem = ExportToDataTable(workItemIteration_cWork);


                //var workIterationFromTimeSheet = ReadDataFromExcel(Filepath);
                Console.WriteLine($"Reading Time sheet data from the file path {Filepath}");
                WriteFileToDisk("", $"Reading Time sheet data from the file path {Filepath}");

                DataSet ds = ReadExcel(Filepath, sheetName);
                readUserAndIterationFromTimeSheet(ds.Tables[0]);

                Console.WriteLine("Comparing TimeSheet data with Azure DevOps data. Please wait...");
                WriteFileToDisk("", "Comparing TimeSheet data with Azure DevOps data. Please wait...");
                DataTable compareData = Comparetable(dtworkItem, ds.Tables[0]);

                ExportDataToExcel(compareData);
                //var workIterationFromDataTable = ReadDataFromDataTable(ds.Tables[0]);
                Console.WriteLine("Completed sucessfully, press any key to exit..");
                WriteFileToDisk("", "Comparing TimeSheet data with Azure DevOps data. Please wait...");

                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Sorry, we ran into some issue. Please check the Debug logs {logFile}");
                Console.ReadLine();
                WriteFileToDisk("Main ", ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        public static List<WorkItems> GetWorkItemsfromSource(string workItemType, UrlParameters parameters)
        {
            GetWorkItemsResponse.Results viewModel = new GetWorkItemsResponse.Results();
            List<WorkItemFetchResponse.WorkItems> fetchedWIs;
            try
            {
                // create wiql object
                Object wiql = new
                {
                    //select [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.IterationPath], [Microsoft.VSTS.Scheduling.CompletedWork] from WorkItems where [System.TeamProject] = @project and [System.WorkItemType] = 'Task' and [System.State] <> '' and [System.IterationPath] = 'unisys\\Iteration 1'
                    query = "select [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.IterationPath], [Microsoft.VSTS.Scheduling.CompletedWork] from WorkItems where [System.TeamProject] = '" + parameters.Project + "' and [System.WorkItemType] = 'Task' and [System.State] <> '' and [System.IterationPath] under '" + parameters.Project + "'"
                };
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", parameters.PatBase);
                    var postValue = new StringContent(JsonConvert.SerializeObject(wiql), Encoding.UTF8, "application/json"); // mediaType needs to be application/json-patch+json for a patch call

                    // set the httpmethod to Patch
                    var method = new HttpMethod("POST");

                    // send the request               
                    var request = new HttpRequestMessage(method, parameters.UriString + "/_apis/wit/wiql?api-version=5.0") { Content = postValue };
                    var response = client.SendAsync(request).Result;

                    if (response.IsSuccessStatusCode && response.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        string res = response.Content.ReadAsStringAsync().Result;
                        viewModel = JsonConvert.DeserializeObject<GetWorkItemsResponse.Results>(res);
                    }
                    else
                    {
                        var errorMessage = response.Content.ReadAsStringAsync();
                    }
                    List<string> witIds = new List<string>();
                    string workitemIDstoFetch = ""; int WICtr = 0;
                    foreach (GetWorkItemsResponse.Workitem WI in viewModel.workItems)
                    {
                        workitemIDstoFetch = WI.id + "," + workitemIDstoFetch;
                        WICtr++;
                        if (WICtr >= 199)
                        {
                            witIds.Add(workitemIDstoFetch = workitemIDstoFetch.Trim(','));
                            workitemIDstoFetch = "";
                        }
                    }
                    if (witIds.Count == 0)
                    {
                        workitemIDstoFetch = workitemIDstoFetch.TrimEnd(',');
                        witIds.Add(workitemIDstoFetch);
                    }
                    fetchedWIs = GetWorkItemsDetailInBatch(witIds, parameters);
                    return fetchedWIs;
                }
            }
            catch (Exception ex)
            {
                string error = ex.Message;
                WriteFileToDisk("GetWorkItemsfromSource", ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return new List<WorkItemFetchResponse.WorkItems>();
        }

        public static List<WorkItems> GetWorkItemsDetailInBatch(List<string> witIDsList, UrlParameters parameters)
        {
            List<WorkItemFetchResponse.WorkItems> viewModelList = new List<WorkItemFetchResponse.WorkItems>();
            try
            {
                if (witIDsList.Count > 0)
                {
                    foreach (var workitemstoFetch in witIDsList)
                    {
                        WorkItemFetchResponse.WorkItems viewModel = new WorkItemFetchResponse.WorkItems();

                        using (var client = new HttpClient())
                        {
                            client.DefaultRequestHeaders.Clear();
                            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", parameters.PatBase);
                            HttpResponseMessage response = client.GetAsync(parameters.UriString + "/_apis/wit/workitems?api-version=5.0&ids=" + workitemstoFetch + "&$expand=relations").Result;
                            if (response.IsSuccessStatusCode && response.StatusCode == System.Net.HttpStatusCode.OK)
                            {
                                string res = response.Content.ReadAsStringAsync().Result;
                                viewModel = JsonConvert.DeserializeObject<WorkItemFetchResponse.WorkItems>(res);
                                viewModelList.Add(viewModel);
                            }
                            else
                            {
                                var errorMessage = response.Content.ReadAsStringAsync();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string error = ex.Message;
                WriteFileToDisk("GetWorkItemsDetailInBatch", ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return viewModelList;
        }

        // read excel data into datatable
        public static DataSet ReadExcel(string excelFilePath, string workSheetName)
        {
            DataSet dsWorkbook = new DataSet();
            try
            {

                string connectionString = string.Empty;

                switch (Path.GetExtension(excelFilePath).ToUpperInvariant())
                {
                    case ".XLS":
                        connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0}; Extended Properties=Excel 8.0;", excelFilePath);
                        break;

                    case ".XLSX":
                        connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties=Excel 12.0;", excelFilePath);
                        break;

                }

                if (!String.IsNullOrEmpty(connectionString))
                {
                    string selectStatement = string.Format("SELECT * FROM [{0}$]", workSheetName);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectStatement, connectionString))
                    {
                        adapter.Fill(dsWorkbook, workSheetName);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteFileToDisk("ReadExcel", ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return dsWorkbook;
        }

        //Read excel data by looping each record
        public static List<WorkItemWithIteration> ReadDataFromExcel(string excelFilePath)
        {
            List<WorkItemWithIteration> workItemIteration = new List<WorkItemWithIteration>();
            try
            {
                Microsoft.Office.Interop.Excel.Application appExl = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = appExl.Workbooks.Open(excelFilePath);
                Microsoft.Office.Interop.Excel._Worksheet NwSheet = workbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range ShtRange = NwSheet.UsedRange;

                int rowCount = ShtRange.Rows.Count;
                int colCount = ShtRange.Columns.Count;

                int rowindex = 0;
                DataTable dt = new DataTable();
                for (int i = 5; i <= rowCount; i++)
                {
                    if (ShtRange.Cells[i, 1] != null && ShtRange.Cells[i, 1].Value2 != null)
                    {
                        if (ShtRange.Cells[i, 1].Value2.ToString() == "Row Labels")
                        {
                            rowindex = i;
                        }
                    }
                    if (i > rowindex)
                    {
                        for (int j = 2; j <= colCount - 1; j++)
                        {
                            if (ShtRange.Cells[i, j] != null && ShtRange.Cells[i, j].Value2 != null)
                            {
                                WorkItemWithIteration wrkitem = new WorkItemWithIteration();
                                if (ShtRange.Cells[i, 1] != null && ShtRange.Cells[i, 1].Value2 != null)
                                {
                                    wrkitem.UserName = ShtRange.Cells[i, 1].Value2.ToString();
                                }

                                if (ShtRange.Cells[rowindex, 1] != null && ShtRange.Cells[rowindex, 1].Value2 != null)
                                {
                                    wrkitem.IterationPath = ShtRange.Cells[rowindex, j].Value2.ToString();
                                }

                                if (ShtRange.Cells[i, j] != null && ShtRange.Cells[i, j].Value2 != null)
                                {
                                    wrkitem.Value = Convert.ToDouble(ShtRange.Cells[i, j].Value2.ToString() == "" ? 0 : ShtRange.Cells[i, j].Value2.ToString());
                                    workItemIteration.Add(wrkitem);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteFileToDisk("ReadDataFromExcel", ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return workItemIteration;
        }

        //Read data by looping DataTable record
        public static List<WorkItemWithIteration> ReadDataFromDataTable(DataTable dt)
        {
            List<WorkItemWithIteration> workItemIteration = new List<WorkItemWithIteration>();
            try
            {
                int rowindex = 0;
                for (int i = 5; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i][1].ToString() == "Row Labels")
                    {
                        rowindex = i;
                    }
                    if (i > rowindex)
                    {
                        for (int j = 0; j < dt.Columns.Count - 1; j++)
                        {
                            WorkItemWithIteration wrkitem = new WorkItemWithIteration();
                            if (dt.Rows[i][0] != null)
                            {
                                wrkitem.UserName = dt.Rows[i][0].ToString();
                            }

                            if (dt.Rows[rowindex][0] != null)
                            {
                                wrkitem.IterationPath = dt.Rows[rowindex][j].ToString();
                            }

                            if (dt.Rows[i][j] != null)
                            {
                                wrkitem.Value = Convert.ToDouble(dt.Rows[i][j].ToString() == "" ? 0 : dt.Rows[i][j]);
                                workItemIteration.Add(wrkitem);
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                WriteFileToDisk("ReadDataFromDataTable", ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return workItemIteration;
        }

        public static void ExportDataToExcel(DataTable dtWorklist)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "Result";

                worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                worKsheeT.Cells[1, 1] = "Cosolidated compared data between Azure Devops and TimeSheet Data";
                worKsheeT.Cells.Font.Size = 10;


                int rowcount = 2;

                foreach (DataRow datarow in dtWorklist.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= dtWorklist.Columns.Count; i++)
                    {

                        if (rowcount == 3)
                        {
                            worKsheeT.Cells[2, i] = dtWorklist.Columns[i - 1].ColumnName;
                            worKsheeT.Cells.Font.Color = System.Drawing.Color.Black;
                        }

                        worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == dtWorklist.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, dtWorklist.Columns.Count]];
                                }

                            }
                        }

                    }

                }
                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, dtWorklist.Columns.Count]];
                celLrangE.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, dtWorklist.Columns.Count]];
                if (!Directory.Exists(path + "\\Reports"))
                {
                    Directory.CreateDirectory(path + "\\Reports");
                }
                string reportPath = path + "\\Reports\\UnisysReport-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                Console.WriteLine($"Writing result to {reportPath}");
                worKbooK.SaveAs(reportPath);
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                WriteFileToDisk("ExportDataToExcel", ex.Message + Environment.NewLine + ex.StackTrace);
                Console.ReadLine();
            }

        }

        public static DataTable ExportToDataTable(List<WorkItemWithIteration> workItem)
        {
            DataTable dt = new DataTable();
            try
            {
                foreach (var witem in workItem)
                {
                    if (!iterationList.Contains(witem.IterationPath))
                    {
                        iterationList.Add(witem.IterationPath);
                    }

                    if (!UserList.Contains(witem.UserName))
                    {
                        UserList.Add(witem.UserName);
                    }
                }
                if (!dt.Columns.Contains("Row Labels"))
                {
                    dt.Columns.Add("Row Labels", typeof(string));
                }
                foreach (string str in iterationList)
                {
                    if (!dt.Columns.Contains(str))
                    {
                        dt.Columns.Add(str, typeof(string));
                    }
                }
                foreach (string struser in UserList)
                {
                    dt.Rows.Add();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["Row Labels"] = UserList[i];
                    foreach (DataColumn dc in dt.Columns)
                    {
                        string column = dc.ColumnName;
                        var element = workItem.Find(e => e.UserName == UserList[i] && e.IterationPath == column);
                        if (element != null)
                        {
                            dt.Rows[i][column] = element.Value.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteFileToDisk("ExportToDataTable", ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return dt;
        }

        public static DataTable Comparetable(DataTable dtVsts, DataTable dtTimesheet)
        {
            DataTable dt = new DataTable();
            try
            {
                for (int j = 0; j < dtTimesheet.Columns.Count; j++)
                {

                    if (j == 0)
                    {
                        dt.Columns.Add(dtTimesheet.Columns[j].ColumnName);
                    }
                    else if (dtTimesheet.Columns[j].ColumnName.Trim() != "Grand Total")
                    {
                        dt.Columns.Add(dtTimesheet.Columns[j].ColumnName.Replace(@"URBIS\", "").Trim() + "-VSTS");
                        dt.Columns.Add(dtTimesheet.Columns[j].ColumnName.Replace(@"URBIS\", "").Trim() + "-TS");
                        dt.Columns.Add(dtTimesheet.Columns[j].ColumnName.Replace(@"URBIS\", "").Trim() + "-Diff");
                    }
                }
                foreach (var user in UserList)
                {
                    dt.Rows.Add();
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["Row Labels"] = UserList[i];
                    for (int j = 0; j < dtTimesheet.Columns.Count; j++)
                    {
                        if (j > 0 && dtTimesheet.Columns[j].ColumnName.Trim() != "Grand Total")
                        {

                            string Column = dtTimesheet.Columns[j].ColumnName;
                            var TimesheetVal = dtTimesheet.Rows[i][Column].ToString() == "" ? 0 : Convert.ToDouble(dtTimesheet.Rows[i][Column].ToString());
                            double vstsVal = 0;
                            if (dtVsts.Columns.Contains(Column) && i < dtVsts.Rows.Count)
                            {
                                int rowindex = 0;
                                for (int k = 0; k < dtVsts.Rows.Count; k++)
                                {
                                    string value = dtVsts.Rows[k]["Row Labels"].ToString();
                                    if (UserList[i].ToLower() == value.ToLower())
                                    {
                                        rowindex = k;
                                        break;
                                    }

                                }
                                if (UserList[i].ToLower() == dtVsts.Rows[rowindex]["Row Labels"].ToString().ToLower())
                                {
                                    vstsVal = dtVsts.Rows[rowindex][Column].ToString() == "" ? 0 : Convert.ToDouble(dtVsts.Rows[rowindex][Column].ToString());
                                }
                            }

                            var difference = vstsVal - TimesheetVal;

                            dt.Rows[i][dtTimesheet.Columns[j].ColumnName.Replace(@"URBIS\", "").Trim() + "-TS"] = TimesheetVal.ToString();
                            dt.Rows[i][dtTimesheet.Columns[j].ColumnName.Replace(@"URBIS\", "").Trim() + "-VSTS"] = vstsVal.ToString();
                            dt.Rows[i][dtTimesheet.Columns[j].ColumnName.Replace(@"URBIS\", "").Trim() + "-Diff"] = difference.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteFileToDisk("Comparetable", ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return dt;
        }

        public static void readUserAndIterationFromTimeSheet(DataTable dt)
        {
            try
            {
                UserList = new List<string>();
                iterationList = new List<string>();
                foreach (DataRow dr in dt.Rows)
                {
                    if (!UserList.Contains(dr["Row Labels"].ToString()))
                    {
                        UserList.Add(dr["Row Labels"].ToString());
                    }
                }
                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    if (!iterationList.Contains(dt.Columns[i].ColumnName))
                    {
                        iterationList.Add(dt.Columns[i].ColumnName);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteFileToDisk("readUserAndIterationFromTimeSheet", ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        private static void WriteFileToDisk(string label, string dataToWriteFile)
        {
            File.AppendAllText(logFile, DateTime.Now.ToString("yyyy:MM:dd:HH:MM:ss") + "\t" + label + "\t" + dataToWriteFile);
        }
    }
}

public class UrlParameters
{
    public string UriString { get; set; }
    public string Account { get; set; }
    public string Project { get; set; }
    public string Pat { get; set; }
    public string PatBase { get; set; }
}
