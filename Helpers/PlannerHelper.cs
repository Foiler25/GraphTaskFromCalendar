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
        /*public async Task CalendarEventCall()
        {


            //await CreatePlannerTask(subject, body, endDate);
        }*/

        public async Task PlannerHelperCall()
        {
            //Group Id is pulled from our planner url
            var groupId = "215f7acd-ce36-4eda-83b1-9f6c23092195"; //Getting the first group we can find to create a plan //(await _graphClient.Me.GetMemberGroups(false).Request().PostAsync()).FirstOrDefault();

            if (groupId != null)
            {
                //planId is pulled from our planner url
                var planId = "yi7xKSBrMEKiq0zFS0mwi2UAHg9l";
                //bucketId is pulled from inspect element and finding the li tag then Id of the bucket
                var bucketId = "ay5DdfT-MUK8mLW1evZGqmUAIPpc";

                //var calEvent = await _graphClient.Me.Events[""]
	            //    .Request()
	            //    .Select("subject,bodypreview,end")
	            //    .GetAsync();

                //var calSearch = await _graphClient.Search
                //    .Query(null)
                //    .Request()
                //    .PostAsync();

                var date1 = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm");
                var date2 = DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-ddTHH:mm");

                var queryOptions = new List<QueryOption>
                {
                    new QueryOption("startDateTime", date1),  
                    new QueryOption("endDateTime", date2)
                
                };

                var calEvent = await _graphClient.Me.CalendarView
	                .Request(queryOptions)
                    //.Filter("Start/DateTime ge '" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm") + "'and contains(subject,'Monthly Technical')")
	                .Select(e => new  
                    {
                        e.Subject,
                        e.BodyPreview,
                        e.End,
                        e.Start
                    })
                    .OrderBy("start/DateTime")
	                .GetAsync();

                foreach (var spevent in calEvent)
                {
                    await CreatePlannerTask(groupId, planId, bucketId, spevent.Subject, spevent.BodyPreview, spevent.End.DateTime);
                }

                //var getcal = calEvent.CurrentPage.ToList();
//
                //var gotevent = getcal.Select(g => new getcalevent{
                //    Subject = g.Subject,
                //    BodyPreview = g.BodyPreview,
                //    End = g.End
                //});
//
                //var subject = (from s in gotevent
                //                select s.Subject).First();
//
                //var body = (from b in gotevent
                //                select b.BodyPreview).First();
//
                //var end = (from d in gotevent
                //                select d.End.DateTime).First();
//
                //await CreatePlannerTask(groupId, planId, bucketId, subject, body, end);
            }
        }
        private async Task<string> GetAndListCurrentPlans(string groupId)
        {
            //Querying plans in current group
            var plans = await _graphClient.Groups[groupId].Planner.Plans.Request(new List<QueryOption>
            {
                new QueryOption("$orderby", "Title asc")
            }).GetAsync();
            if (plans.Any())
            {
                Console.WriteLine($"Number of plans in current tenant: {plans.Count}");
                Console.WriteLine(plans.Select(x => $"-- {x.Title}").Aggregate((x, y) => $"{x}\n{y}"));
                return plans.First().Id;
            }
            else
            {
                Console.WriteLine("No existing plan");
                return null;
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