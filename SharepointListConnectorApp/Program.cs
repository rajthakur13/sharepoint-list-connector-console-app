using System;
using Microsoft.SharePoint.Client;
using System.Security;

namespace SharepointListConnectorApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = "https://x2k47.sharepoint.com/sites/SharepointCustomSite";
            string username = "admin@x2k47.onmicrosoft.com";
            string password = "Peruguprabhakar@7702";

            try
            {
                var securePassword = new SecureString();
                foreach (char c in password) securePassword.AppendChar(c);
                var credentials = new SharePointOnlineCredentials(username, securePassword);

                using (var context = new ClientContext(siteUrl))
                {
                    context.Credentials = credentials;

                    var listCreationInfo = new ListCreationInformation
                    {
                        Title = "Employee List",
                        TemplateType = (int)ListTemplateType.GenericList
                    };
                    var newList = context.Web.Lists.Add(listCreationInfo);

                    context.Load(newList);
                    context.ExecuteQuery();

                    AddColumn(context, newList, "EmployeeID", "Employee ID", "Text");
                    AddColumn(context, newList, "EmployeeName", "Employee Name", "Text");
                    AddColumn(context, newList, "EmployeeEmail", "Employee Email", "Text");
                    AddColumn(context, newList, "EmployeePhoneNumber", "Employee Phone Number", "Text");
                    AddColumn(context, newList, "EmployeeAddress", "Employee Address", "Text");

                    Console.WriteLine("List and columns created successfully!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            Console.ReadLine();
        }

        private static void AddColumn(ClientContext context, List list, string internalName, string displayName, string fieldType)
        {
            try
            {
                var field = list.Fields.GetByInternalNameOrTitle(internalName);
                context.Load(field);
                context.ExecuteQuery();
                Console.WriteLine($"Field '{displayName}' already exists.");
            }
            catch (ServerException)
            {
                string fieldXml = $@"
                    <Field DisplayName='{displayName}' 
                           Type='{fieldType}' 
                           Name='{internalName}' 
                           StaticName='{internalName}' 
                           Group='Custom Columns' />
                ";

                list.Fields.AddFieldAsXml(fieldXml, true, AddFieldOptions.DefaultValue);
                context.ExecuteQuery();
                Console.WriteLine($"Field '{displayName}' created.");
            }
        }
    }
}
