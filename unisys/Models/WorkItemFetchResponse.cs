using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace unisys.Models
{
    public class WorkItemFetchResponse
    {
        public class WorkItems
        {
            public int count { get; set; }
            public IList<Value> value { get; set; }
        }
        public class Value
        {

            public int id { get; set; }

            public int rev { get; set; }

            public Fields fields { get; set; }

            public string url { get; set; }
        }

        public class Fields
        {

            [JsonProperty("System.AreaPath")]
            public string SystemAreaPath { get; set; }

            [JsonProperty("System.TeamProject")]
            public string SystemTeamProject { get; set; }

            [JsonProperty("System.IterationPath")]
            public string SystemIterationPath { get; set; }

            [JsonProperty("System.WorkItemType")]
            public string SystemWorkItemType { get; set; }

            [JsonProperty("System.State")]
            public string SystemState { get; set; }

            [JsonProperty("System.AssignedTo")]
            public SystemAssignedTo SystemAssignedTo { get; set; }

            [JsonProperty("System.Title")]
            public string SystemTitle { get; set; }

            [JsonProperty("Microsoft.VSTS.Scheduling.OriginalEstimate")]
            public double MicrosoftVSTSSchedulingOriginalEstimate { get; set; }

            [JsonProperty("Microsoft.VSTS.Scheduling.RemainingWork")]
            public double MicrosoftVSTSSchedulingRemainingWork { get; set; }
            [JsonProperty("Microsoft.VSTS.Scheduling.CompletedWork")]
            public double MicrosoftVSTSSchedulingCompletedWork { get; set; }

        }
        public class SystemAssignedTo
        {

            [JsonProperty("displayName")]
            public string displayName { get; set; }

            [JsonProperty("url")]
            public string url { get; set; }

            [JsonProperty("id")]
            public string id { get; set; }

            [JsonProperty("uniqueName")]
            public string uniqueName { get; set; }

            [JsonProperty("imageUrl")]
            public string imageUrl { get; set; }

            [JsonProperty("descriptor")]
            public string descriptor { get; set; }
        }

        public class WorkItemWithIteration
        {
            public string UserName { get; set; }
            public string IterationPath { get; set; }
            public double Value { get; set; }
        }
    }
}
