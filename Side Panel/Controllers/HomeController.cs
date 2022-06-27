using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using AdaptiveCards;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using SidePanel.Models;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using System.Collections.Concurrent;
using System.Threading;
using System.IO;
using Microsoft.Bot.Schema.Teams;
using System.Web;

namespace SidePanel.Controllers
{
    public class HomeController : Controller
    {
        public static string conversationId;
        public static string serviceUrl;
        public static List<TaskInfo> taskInfoData = new List<TaskInfo>();

        private readonly IConfiguration _configuration;
        private readonly AppCredentials botCredentials;
        private readonly HttpClient httpClient;
        private readonly IBotFrameworkHttpAdapter _adapter;
        private readonly string[] _cards =
        {
            Path.Combine(".", "Resources", "FlightItineraryCard.json")
        };

        public HomeController(IBotFrameworkHttpAdapter adapter,  IConfiguration configuration, IHttpClientFactory httpClientFactory, AppCredentials botCredentials)
        {
            _adapter = adapter;
            _configuration = configuration;
            this.botCredentials = botCredentials;
            this.httpClient = httpClientFactory.CreateClient();
        }

        //Configure call from Manifest
        [Route("/Home/Configure")]
        public ActionResult Configure()
        {
            return View("Configure");
        }

        //SidePanel Call from Configure
        [Route("/Home/SidePanel")]
        public ActionResult SidePanel()
        {
            List<TaskInfo> model = this.SidePanelDefaultAgendaList();
            return PartialView("SidePanel", model);
        }

        //Add Default Agenda to the List
        private List<TaskInfo> SidePanelDefaultAgendaList()
        {
            if (taskInfoData.Count == 0)
            {
                var tData1 = new TaskInfo
                {
                    Title = "Approve 5% dividend payment to shareholders."
                };
                taskInfoData.Add(tData1);
                var tData2 = new TaskInfo
                {
                    Title = "Increase research budget by 10%."
                };
                taskInfoData.Add(tData2);
                var tData3 = new TaskInfo
                {
                    Title = "Continue with WFH for next 3 months."
                };
                taskInfoData.Add(tData3);
            }
            return taskInfoData;
        }

        //Add New Agenda Point to the Agenda List
        [Route("/Home/AddNewAgendaPoint")]
        public List<TaskInfo> AddNewAgendaPoint(TaskInfo taskInfo)
        {
            var tData = new TaskInfo
            {
                Title = taskInfo.Title
            };
            taskInfoData.Add(tData);
            return taskInfoData;
        }

        //Senda Agenda List to the Meeting Chat
        [Route("/Home/SendAgenda")]
        public void SendAgenda()
        {
            string appId = _configuration["MicrosoftAppId"];
            string appSecret = _configuration["MicrosoftAppPassword"];
            using var connector = new ConnectorClient(new Uri(serviceUrl), appId, appSecret);
            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl, DateTime.MaxValue);
            var replyActivity = new Activity();
            replyActivity.Type = "message";
            replyActivity.Conversation = new ConversationAccount(id: conversationId);
            var adaptiveAttachment = AgendaAdaptiveList();
            replyActivity.Attachments = new List<Attachment> { adaptiveAttachment };
            var response = connector.Conversations.SendToConversationAsync(conversationId, replyActivity).Result;
        }

        //Create Adaptive Card with the Agenda List
        private Attachment AgendaAdaptiveList()
        {
            AdaptiveCard adaptiveCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0));
            adaptiveCard.Body = new List<AdaptiveElement>()
            {
                new AdaptiveTextBlock(){Text="Here is the Agenda for Today", Weight=AdaptiveTextWeight.Bolder}
            };

            foreach (var agendaPoint in taskInfoData)
            {
                var textBlock = new AdaptiveTextBlock() { Text = "- " + agendaPoint.Title + " \r" };
                adaptiveCard.Body.Add(textBlock);
            }

            return new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard
            };
        }

        //Check if the Participant Role is Organizer
        [Route("/Home/IsOrganizer")]
        public async Task<ActionResult<bool>> IsOrganizer(string userId, string meetingId, string tenantId)
        {
            var newReference = new ConversationReference()
            {
                Bot = new ChannelAccount()
                {
                    Id = _configuration["MicrosoftAppId"]
                },
                Conversation = new ConversationAccount()
                {
                    Id = tenantId
                },
                ServiceUrl = "https://smba.trafficmanager.net/in/",
            };

            await ((BotAdapter)_adapter).ContinueConversationAsync(_configuration["MicrosoftAppId"], newReference, BotCallback, default(CancellationToken));
            return true;

          /*  var response = await GetMeetingRoleAsync(meetingId, userId, tenantid);
            if (response.meeting.role == "Organizer")
                return true;
            else
                return false;*/
        }

        private async Task BotCallback(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            Random r = new Random();
            var cardAttachment = CreateAdaptiveCardAttachment(_cards[0]);

            var activity = MessageFactory.Attachment(cardAttachment);
            activity.ChannelData = new TeamsChannelData
            {
                Notification = new NotificationInfo()
                {
                    AlertInMeeting = true,
                    ExternalResourceUrl = $"https://teams.microsoft.com/l/bubble/e4bad6ab-52d6-4646-8fcf-4dd03fd10ee5?url=" +
                                                  HttpUtility.UrlEncode($"https://d1e4-2604-3d09-2881-8f00-6845-7ed8-6f3f-577d.ngrok.io?topic=Topic Title") +
                                                  $"&height=270&width=250&title=ContentBubble&completionBotId=e4bad6ab-52d6-4646-8fcf-4dd03fd10ee5"
                }
            };
            await turnContext.SendActivityAsync(activity);

        }

        private static Attachment CreateAdaptiveCardAttachment(string filePath)
        {
            var adaptiveCardJson = System.IO.File.ReadAllText(filePath);
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = JsonConvert.DeserializeObject(adaptiveCardJson),
            };
            return adaptiveCardAttachment;
        }

        public async Task<UserMeetingRoleServiceResponse> GetMeetingRoleAsync(string meetingId, string userId, string tenantId)
        {
            if (serviceUrl == null)
            {
                throw new InvalidOperationException("Service URL is not avaiable for tenant ID " + tenantId);
            }

            using var getRoleRequest = new HttpRequestMessage(HttpMethod.Get, new Uri(new Uri(serviceUrl), string.Format("v1/meetings/{0}/participants/{1}?tenantId={2}", meetingId, userId, tenantId)));
            getRoleRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await this.botCredentials.GetTokenAsync());

            using var getRoleResponse = await this.httpClient.SendAsync(getRoleRequest);
            getRoleResponse.EnsureSuccessStatusCode();

            var response = JsonConvert.DeserializeObject<UserMeetingRoleServiceResponse>(await getRoleResponse.Content.ReadAsStringAsync());
            return response;
        }
    }
}
