using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace O365MSGraphConsole
{
   public class GraphService
   {
      static string m_FilterUnreadAttachment = "isRead ne true and hasAttachments eq true";

      public GraphServiceClient InitializeClient()
      {
         try
         {
            var app = ConfidentialClientApplicationBuilder
                                           .Create(Properties.Settings.Default.ClientId)
                                           .WithTenantId(Properties.Settings.Default.TenantId)
                                           .WithClientSecret(Properties.Settings.Default.ClientPwd)
                                           .Build();

            var accounts = app.GetAccountsAsync().Result;
            foreach (var user in accounts)
            {
               app.RemoveAsync(user);
            }

            var authProvider = new ClientCredentialProvider(app);

            GraphServiceClient gsc = new GraphServiceClient(authProvider);

            return gsc;
         }
         catch (Exception e)
         {
            Console.WriteLine(e);
            throw;
         }
      }

      public async Task<List<string>> GetUnreadMailSubjectFromInbox(GraphServiceClient client, string userId)
      {
         List<string> rs = new List<string>();

         var user = client.Users[userId];
         var inbox = await user.MailFolders.Inbox.Request().GetAsync();
         var unreadCount = inbox.UnreadItemCount;
         Console.WriteLine("Unread email count in Inbox:" + unreadCount);

         var messages = await user.MailFolders.Inbox.Messages.Request()
                                  .Filter(m_FilterUnreadAttachment)
                                  .Top(unreadCount.GetValueOrDefault())
                                  .GetAsync();

         if (messages?.Count > 0)
         {
            foreach (Message message in messages)
            {
               var attachments = user.Messages[message.Id].Attachments.Request().GetAsync();

               string msg = $"Subject: {message.Subject} - Attachment count: {attachments.Result.Count}";
               Console.WriteLine("Processing:" + message.Id);
               rs.Add(msg);
               message.IsRead = true;

               // Mark read
               await user.Messages[message.Id].Request().Select("IsRead").UpdateAsync(new Message {IsRead = true});
            }

         }

        

         return rs;
      }
   }
}
