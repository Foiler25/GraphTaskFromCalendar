using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace GraphTaskFromCalendar
{
    public class PlannerHelper
    {
        private readonly GraphServiceClient _graphClient;
        public PlannerHelper(GraphServiceClient graphClient)
        {
            _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        }

        public async Task PlannerHelperCall()
        {
            //Group Id is pulled from our planner url
            var groupId = "groupid"; 

            if (groupId != null)
            {
                //planId is pulled from our planner url
                var planId = "planid";
                //bucketId is pulled from inspect element and finding the li tag then Id of the bucket
                var bucketId = "bucketid";

                var date = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm");
                var date2 = DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-ddTHH:mm");

                var queryOptions = new List<QueryOption>
                {
                    new QueryOption("startDateTime", date),  
                    new QueryOption("endDateTime", date2)
                
                };

                var calEvent = await _graphClient.Me.CalendarView
	                .Request(queryOptions)
	                .Select(e => new  
                    {
                        e.Subject,
                        e.BodyPreview,
                        e.End,
                        e.Start
                    })
                    .OrderBy("start/DateTime")
                    //.Top(5)
	                .GetAsync();

                foreach (var spevent in calEvent)
                {
                    await CreatePlannerTask(groupId, planId, bucketId, spevent.Subject, spevent.BodyPreview, spevent.End.DateTime);
                }
            }
        }
        private async Task CreatePlannerTask(string groupId, string planId, string bucketId, string subject, string body, string end)
        {
            // Preparing the assignment for the task
            var assignments = new PlannerAssignments();

            // Creating a task within the bucket
            var createdTask = await _graphClient.Planner.Tasks.Request().AddAsync(
                new PlannerTask
                {
                    DueDateTime = DateTime.Parse(end),
                    Title = subject,
                    Details = new PlannerTaskDetails
                    {
                        Description = body
                    },
                    Assignments = assignments,
                    PlanId = planId,
                    BucketId = bucketId
                }
            );
            Console.WriteLine($"Added new task {createdTask.Title} to bucket");
        }
    }
}