using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace CSOMDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ProShow();
            Console.Read();
        }

        public static ClientContext GetClientContext()
        {
            ClientContext clientContext = new ClientContext("https://bigapp.sharepoint.com/sites/demo");

            SecureString ss = new SecureString();
            "$RFV5tgb^YHN".ToCharArray().ToList().ForEach(ss.AppendChar);
            clientContext.Credentials = new SharePointOnlineCredentials("gamma.chen@baron.space", ss);
            return clientContext;
        }

        public static void FirstConn()
        {
            using (ClientContext clientContext=new ClientContext("https://bigapp.sharepoint.com"))
            {
                SecureString password = new SecureString();
                "$RFV5tgb^YHN".ToCharArray().ToList().ForEach(password.AppendChar);
                clientContext.Credentials = new SharePointOnlineCredentials("gamma.chen@baron.space", password);
                var testsSite = clientContext.Site;
                clientContext.Load(testsSite);
                Console.WriteLine("Sitecollection info:" + testsSite.Id);

                clientContext.ExecuteQuery();
                
            }
        }

        public static void ExceptionTestError(string listName)
        {
            ClientContext clientContext = GetClientContext();
            Web web = clientContext.Web;

            ExceptionHandlingScope ehs = new ExceptionHandlingScope(clientContext);
            using (ehs.StartScope())
            {
                using (ehs.StartTry())
                {
                    List list = web.Lists.GetByTitle(listName);
                    list.Description = "List get";
                    list.Update();
                }
                using (ehs.StartCatch())
                {
                    ListCreationInformation create = new ListCreationInformation();
                    create.Title = listName;
                    create.TemplateType = (int)ListTemplateType.DocumentLibrary;
                    create.Description = "List create";
                    web.Lists.Add(create);
                }
            }
            List result = web.Lists.GetByTitle(listName);
            clientContext.Load(result);
            clientContext.ExecuteQuery();//执行查询,不会出异常
            Console.WriteLine("Exception" + ehs.HasException);
            Console.WriteLine("Message" + ehs.ErrorMessage);
        }
        public static void ConditionalScope()
        {
            ClientContext cc = GetClientContext();

            var file = cc.Web.GetFileByServerRelativeUrl("Site/Demo/Shared%20Documents/th.jpg");


            ConditionalScope cs = new ConditionalScope(cc,()=>file.Exists,true);
            //cc.Load(file);
            //Console.WriteLine(file.Exists);
            using (cs.StartScope())
            {
                file.ListItemAllFields["Test"] = "StartIfTrue";
                file.ListItemAllFields.Update();
            }
            cc.ExecuteQuery();
            if (cs.TestResult.HasValue)
            {
                Console.WriteLine(cs.TestResult.Value);
            }
        }

        public static void ProShow()
        {
            //ClientContext cc = GetClientContext();
            //Web web = cc.Web;
            //web.Title = "aaa";
            //web.Description = "This is a test";
            //web.Update();
            //cc.ExecuteQuery();
			string testName="This is a test";
            Console.WriteLine("ok");

        }

    }
}
