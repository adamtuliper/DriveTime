using OfficeAddInServerAuth.Helpers;
using OfficeAddInServerAuth.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using Newtonsoft.Json;
using System.Text;
using System.IO;

namespace OfficeAddInServerAuth.Controllers
{
    public class MessageController : Controller
    {
        public ActionResult Index(SendMessageResponse sendMessageResponse, UserInfo userInfo)
        {
            EnsureUser(ref userInfo);

            ViewBag.UserInfo = userInfo;
            ViewBag.MessageResponse = sendMessageResponse;

            return View();
        }

        public ActionResult Facebook(SendMessageResponse sendMessageResponse, UserInfo userInfo)
        {
            EnsureUser(ref userInfo);

            ViewBag.UserInfo = userInfo;
            ViewBag.MessageResponse = sendMessageResponse;

            return View();
        }

        private Task<dynamic> GetModelAsync()
        {
            var path = Server.MapPath("~/Models/CreateEvent.json");
            var json = System.IO.File.ReadAllText(path);
            dynamic data = JsonConvert.DeserializeObject(json);
            return data;
        }

        public async Task Create(DateTime start, DateTime end, string subject)
        {
            var url = "http://graph.microsoft.com/v1.0/me/events";
            dynamic appointment = await GetModelAsync();
            appointment.end.dateTime = "HELLO WORLD";
            var json = JsonConvert.SerializeObject(appointment);
            var token = Data.GetUserSessionTokenAny(Settings.GetUserAuthStateId(ControllerContext.HttpContext)).AccessToken;

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                    var r = await client.SendAsync(request);
                }
            }
        }

        public async Task SendMessageSubmit(UserInfo userInfo)
        {
            var url = "http://graph.microsoft.com/v1.0/me/events";
            var path = Server.MapPath("~/Models/CreateEvent.json");
            var json = System.IO.File.ReadAllText(path);
            dynamic data = JsonConvert.DeserializeObject(json);
            data.end.dateTime = "HELLO WORLD";
            json = JsonConvert.SerializeObject(data);
            var token = Data.GetUserSessionTokenAny(Settings.GetUserAuthStateId(ControllerContext.HttpContext)).AccessToken;

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                    var r = await client.SendAsync(request);
                }
            }

            //// After Index method renders the View, user clicks Send Mail, which comes in here.
            //EnsureUser(ref userInfo);
            //SendMessageResponse sendMessageResult = new SendMessageResponse();
            //// Send email using the Microsoft Graph API.
            //var token = Data.GetUserSessionTokenAny(Settings.GetUserAuthStateId(ControllerContext.HttpContext));

            //if (token.Provider == Settings.AzureADAuthority || token.Provider == Settings.AzureAD2Authority)
            //{
            //    sendMessageResult = await GraphApiHelper.SendMessageAsync(
            //        token.AccessToken,
            //        GenerateEmail(userInfo));
            //}
            //else if (token.Provider == Settings.GoogleAuthority)
            //{
            //    sendMessageResult = await GoogleApiHelper.SendMessageAsync(token.AccessToken, GenerateEmail(userInfo), token.Username);
            //}
            //// Reuse the Index view for messages (sent, not sent, fail) .
            //// Redirect to tell the browser to call the app back via the Index method.
            //return RedirectToAction(nameof(Index), new RouteValueDictionary(new Dictionary<string, object>{
            //    { "Status", sendMessageResult.Status },
            //    { "StatusMessage", sendMessageResult.StatusMessage },
            //    { "Address", userInfo.Address },
            //}));
        }


        public async Task<ActionResult> FacebookSendMessageSubmit(UserInfo userInfo)
        {
            // After Index method renders the View, user clicks Send Mail, which comes in here.
            EnsureUser(ref userInfo);
            SendMessageResponse sendMessageResult = new SendMessageResponse();
            // Send email using the Microsoft Graph API.
            var token = Data.GetUserSessionTokenAny(Settings.GetUserAuthStateId(ControllerContext.HttpContext));

            if (token.Provider == Settings.FacebookAuthority)
            {
                sendMessageResult =
                    await FacebookApiHelper.PostMessageAsync(token.AccessToken, token.Username, Settings.MessageSubject);
            }
            // Reuse the Index view for messages (sent, not sent, fail) .
            // Redirect to tell the browser to call the app back via the Index method.
            return RedirectToAction(nameof(Facebook), new RouteValueDictionary(new Dictionary<string, object>{
                { "Status", sendMessageResult.Status },
                { "StatusMessage", sendMessageResult.StatusMessage },
                { "Address", userInfo.Address },
            }));
        }



        // Use the login user name or recipient email address if no user name.
        void EnsureUser(ref UserInfo userInfo)
        {
            var token = Data.GetUserSessionTokenAny(Settings.GetUserAuthStateId(ControllerContext.HttpContext));
            var currentUser = new UserInfo() { Name = token.Username, Address = token.Username };


            if (string.IsNullOrEmpty(userInfo?.Address))
            {
                userInfo = currentUser;
            }
            else if (userInfo.Address.Equals(currentUser.Address, StringComparison.OrdinalIgnoreCase))
            {
                userInfo = currentUser;
            }
            else
            {
                userInfo.Name = userInfo.Address;
            }
        }

        // Create email with predefine body and subject.
        SendMessageRequest GenerateEmail(UserInfo to)
        {
            return CreateEmailObject(
                to: to,
                subject: Settings.MessageSubject,
                body: string.Format(Settings.MessageBody, to.Name)
            );
        }

        // Create email object in the required request format/data contract.
        private SendMessageRequest CreateEmailObject(UserInfo to, string subject, string body)
        {
            return new SendMessageRequest
            {
                Message = new Message
                {
                    Subject = subject,
                    Body = new MessageBody
                    {
                        ContentType = "Html",
                        Content = body
                    },
                    ToRecipients = new List<Recipient>
                    {
                        new Recipient
                        {
                            EmailAddress = new UserInfo
                            {
                                 Name =  to.Name,
                                 Address = to.Address
                            }
                        }
                    }
                },
                SaveToSentItems = true
            };
        }

        public async Task<ActionResult> DropBox()
        {
            var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.DropBoxAuthority);
            var usage = await DropBoxApiHelper.GetDropBoxSpaceUsage(token);
            return View(usage);
        }
    }
}