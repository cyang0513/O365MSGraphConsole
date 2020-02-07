using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365MSGraphConsole
{
   class Program
   {
      static void Main(string[] args)
      {
         string userId = Properties.Settings.Default.UserId;

         GraphService gs = new GraphService();

         Console.WriteLine("Initialize client...");
         var client = gs.InitializeClient();

         var msgSubjects = gs.GetUnreadMailSubjectFromInbox(client, userId);

         Console.WriteLine("Printing subjects...");
         foreach (var subject in msgSubjects.Result)
         {
            Console.WriteLine(subject);
         }

         Console.WriteLine("Done!");
         Console.ReadKey();

      }
   }
}
