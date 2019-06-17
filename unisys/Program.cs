using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using unisys.Models;

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
            List<Tuple<string, string, double>> workItemIteration_cWork = new List<Tuple<string, string, double>>();
            List<WorkItemFetchResponse.WorkItems> workItemsList = GetWorkItemsfromSource("Task", parameters);
            if (workItemsList.Count > 0)
            {
                List<string> users = new List<string>();
                foreach (var ItemList in workItemsList)
                {
                    foreach (var workItem in ItemList.value)
                    {
                        if (!users.Contains(workItem.fields.SystemAssignedTo.displayName))
                        {
                            users.Add(workItem.fields.SystemAssignedTo.displayName);
                        }
                        workItemIteration_cWork.Add(Tuple.Create(workItem.fields.SystemAssignedTo.displayName, workItem.fields.SystemIterationPath, workItem.fields.MicrosoftVSTSSchedulingCompletedWork));
                        Console.WriteLine("User: " + workItem.fields.SystemAssignedTo.displayName + "\n Complated Work: " + workItem.fields.MicrosoftVSTSSchedulingCompletedWork + "\n Original Estimate " + workItem.fields.MicrosoftVSTSSchedulingOriginalEstimate + Environment.NewLine);
                        Console.WriteLine();
                    }
                }
            }
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
