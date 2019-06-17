using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using unisys.Models;
using static unisys.Models.WorkItemFetchResponse;

namespace unisys
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Replace("file:\\", "");
            Console.WriteLine(path);
            string pat = "rahfqiseaujjsyj5qv7o25yghyicy3cgtrqu3cse267be52lu2na";
            string basePat = Convert.ToBase64String(Encoding.ASCII.GetBytes(string.Format("{0}:{1}", "", pat)));

            UrlParameters parameters = new UrlParameters();
            parameters.Project = "unisys";
            parameters.Account = "devteamtests";
            parameters.UriString = "https://dev.azure.com/" + parameters.Account + "/" + parameters.Project;
            parameters.PatBase = basePat;
            parameters.Pat = pat;
            //List<Tuple<string, string, double>> workItemIteration_cWork = new List<Tuple<string, string, double>>();
            List<WorkItemWithIteration> workItemIteration_cWork = new List<WorkItemWithIteration>();
            List<WorkItemFetchResponse.WorkItems> workItemsList = GetWorkItemsfromSource("Task", parameters);
            if (workItemsList.Count > 0)
            {
                List<string> users = new List<string>();
                foreach (var ItemList in workItemsList)
                {
                    foreach (var workItem in ItemList.value)
                    {
                        var element = workItemIteration_cWork.Find(e => e.UserName == workItem.fields.SystemAssignedTo.displayName && e.IterationPath == workItem.fields.SystemIterationPath);
                        if (element != null)
                            element.Value = element.Value + workItem.fields.MicrosoftVSTSSchedulingCompletedWork;
                        else
                            workItemIteration_cWork.Add(new WorkItemWithIteration
                            {
                                UserName = workItem.fields.SystemAssignedTo.displayName,
                                IterationPath = workItem.fields.SystemIterationPath,
                                Value = workItem.fields.MicrosoftVSTSSchedulingCompletedWork
                            });

                        Console.WriteLine("User: " + workItem.fields.SystemAssignedTo.displayName + "\n Complated Work: " + workItem.fields.MicrosoftVSTSSchedulingCompletedWork + "\n Original Estimate " + workItem.fields.MicrosoftVSTSSchedulingOriginalEstimate + Environment.NewLine);
                        Console.WriteLine();
                    }
                }
            }

            string Filepath = @"D:\Unisys\Anwar Report.xlsx";            
            var workIterationFromTimeSheet= ReadDataFromExcel(Filepath);

            DataSet ds = ReadExcel(Filepath, "Sheet1");
            var workIterationFromDataTable = ReadDataFromDataTable(ds.Tables[0]);


            Console.ReadLine();
        }


        public static List<WorkItemFetchResponse.WorkItems> GetWorkItemsfromSource(string workItemType, UrlParameters parameters)
        {
            GetWorkItemsResponse.Results viewModel = new GetWorkItemsResponse.Results();
            List<WorkItemFetchResponse.WorkItems> fetchedWIs;
            try
            {
                // create wiql object
                Object wiql = new
                {
                    //select [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.IterationPath], [Microsoft.VSTS.Scheduling.CompletedWork] from WorkItems where [System.TeamProject] = @project and [System.WorkItemType] = 'Task' and [System.State] <> '' and [System.IterationPath] = 'unisys\\Iteration 1'
                    query = "select [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.IterationPath], [Microsoft.VSTS.Scheduling.CompletedWork] from WorkItems where [System.TeamProject] = '" + parameters.Project + "' and [System.WorkItemType] = 'Task' and [System.State] <> '' and [System.IterationPath] = 'unisys\\Iteration 1'"
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
                        Console.Write(errorMessage);
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
            }
            return new List<WorkItemFetchResponse.WorkItems>();
        }

        public static List<WorkItemFetchResponse.WorkItems> GetWorkItemsDetailInBatch(List<string> witIDsList, UrlParameters parameters)
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
            }
            return viewModelList;
        }

        // read excel data into datatable
        public static DataSet ReadExcel(string excelFilePath, string workSheetName)
        {
            DataSet dsWorkbook = new DataSet();

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

            return dsWorkbook;
        }

        //Read excel data by looping each record
        public static List<WorkItemWithIteration> ReadDataFromExcel(string excelFilePath)
        {
            List<WorkItemWithIteration> workItemIteration = new List<WorkItemWithIteration>();

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
                if(i>rowindex)
                {
                    for (int j = 2; j <= colCount-1; j++)
                    {
                        if (ShtRange.Cells[i, j] != null && ShtRange.Cells[i, j].Value2 != null)
                        {
                            WorkItemWithIteration wrkitem = new WorkItemWithIteration();
                            if (ShtRange.Cells[i, 1] != null && ShtRange.Cells[i, 1].Value2 != null)
                                wrkitem.UserName = ShtRange.Cells[i, 1].Value2.ToString();
                            if (ShtRange.Cells[rowindex, 1] != null && ShtRange.Cells[rowindex, 1].Value2 != null)
                                wrkitem.IterationPath = ShtRange.Cells[rowindex, j].Value2.ToString();
                            if (ShtRange.Cells[i, j] != null && ShtRange.Cells[i, j].Value2 != null)
                            {
                                wrkitem.Value = Convert.ToDouble(ShtRange.Cells[i, j].Value2.ToString() == "" ? 0 : ShtRange.Cells[i, j].Value2.ToString());
                                workItemIteration.Add(wrkitem);
                            }
                        }
                    }
                }
                
                Console.WriteLine();
            }
            return workItemIteration;
        }

        //Read data by looping DataTable record
        public static List<WorkItemWithIteration> ReadDataFromDataTable(DataTable dt)
        {
            List<WorkItemWithIteration> workItemIteration = new List<WorkItemWithIteration>();
            int rowindex = 0;
            for(int i=5; i < dt.Rows.Count; i++)
            {
                if(dt.Rows[i][1].ToString()== "Row Labels")
                {
                    rowindex = i;
                }
                if (i > rowindex)
                {
                    for (int j = 0; j < dt.Columns.Count-1; j++)
                    {
                        WorkItemWithIteration wrkitem = new WorkItemWithIteration();
                        if (dt.Rows[i][0] != null)
                            wrkitem.UserName = dt.Rows[i][0].ToString();
                        if (dt.Rows[rowindex][0] != null)
                            wrkitem.IterationPath = dt.Rows[rowindex][j].ToString();
                        if (dt.Rows[i][j] != null)
                        {
                            wrkitem.Value = Convert.ToDouble(dt.Rows[i][j].ToString() == "" ? 0 : dt.Rows[i][j]);
                            workItemIteration.Add(wrkitem);
                        }
                    }
                }
                
            }
            return workItemIteration;
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
