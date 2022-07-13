using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;
using System.Threading.Tasks;

namespace graphconsoleapp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var config = LoadAppSettings();
            if (config == null)
            {
                Console.WriteLine("Invalid appsettings.json file.");
                return;
            }

            var client = GetAuthenticatedGraphClient(config);

            //retrieve users
            // // request 1 - all users
            // var requestAllUsers = client.Users.Request();

            // var results = requestAllUsers.GetAsync().Result;
            // foreach (var user in results)
            // {
            //     Console.WriteLine(user.Id + ": " + user.DisplayName + " <" + user.Mail + ">");
            // }

            // Console.WriteLine("\nGraph Request:");
            // Console.WriteLine(requestAllUsers.GetHttpRequestMessage().RequestUri);

            // // request 2 - current user
            // var requestMeUser = client.Me.Request();

            // var resultMe = requestMeUser.GetAsync().Result;
            // Console.WriteLine(resultMe.Id + ": " + resultMe.DisplayName + " <" + resultMe.Mail + ">");

            // Console.WriteLine("\nGraph Request:");
            // Console.WriteLine(requestMeUser.GetHttpRequestMessage().RequestUri);

            // // request 3 - specific user
            // var requestSpecificUser = client.Users["{110e21b9-8447-4c75-9709-e98aa4530825}"].Request();
            // var resultOtherUser = requestSpecificUser.GetAsync().Result;
            // Console.WriteLine(resultOtherUser.Id + ": " + resultOtherUser.DisplayName + " <" + resultOtherUser.Mail + ">");

            // Console.WriteLine("\nGraph Request:");
            // Console.WriteLine(requestSpecificUser.GetHttpRequestMessage().RequestUri);

            //Bogdan - fecth user photo and manager
            // // request 1 - current user's photo
            // var requestUserPhoto = client.Me.Photo.Request();
            // var resultsUserPhoto = requestUserPhoto.GetAsync().Result;
            // // display photo metadata
            // Console.WriteLine("                Id: " + resultsUserPhoto.Id);
            // Console.WriteLine("media content type: " + resultsUserPhoto.AdditionalData["@odata.mediaContentType"]);
            // Console.WriteLine("        media etag: " + resultsUserPhoto.AdditionalData["@odata.mediaEtag"]);

            // Console.WriteLine("\nGraph Request:");
            // Console.WriteLine(requestUserPhoto.GetHttpRequestMessage().RequestUri);

            // // get actual photo
            // var requestUserPhotoFile = client.Me.Photo.Content.Request();
            // // var requestUserPhotoFile = client.Users["{110e21b9-8447-4c75-9709-e98aa4530825}"].Photo.Content.Request();
            // var resultUserPhotoFile = requestUserPhotoFile.GetAsync().Result;

            // // create the file
            // var profilePhotoPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "profilePhoto_" + resultsUserPhoto.Id + ".jpg");
            // var profilePhotoFile = System.IO.File.Create(profilePhotoPath);
            // resultUserPhotoFile.Seek(0, System.IO.SeekOrigin.Begin);
            // resultUserPhotoFile.CopyTo(profilePhotoFile);
            // Console.WriteLine("Saved file to: " + profilePhotoPath);

            // Console.WriteLine("\nGraph Request:");
            // Console.WriteLine(requestUserPhoto.GetHttpRequestMessage().RequestUri);

            // // request 2 - user's manager
            // var userId = "{3a5e345f-c4a4-41a5-ad99-c0f8c0f64c80}";
            // var requestUserManager = client.Users[userId]
            //                                 .Manager
            //                                 .Request();
            // var resultsUserManager = requestUserManager.GetAsync().Result;
            // Console.WriteLine("   User: " + userId);
            // Console.WriteLine("Manager: " + resultsUserManager.Id);
            // var resultsUserManagerUser = resultsUserManager as Microsoft.Graph.User;
            // if (resultsUserManagerUser != null)
            // {
            //     Console.WriteLine("Manager: " + resultsUserManagerUser.DisplayName);
            //     Console.WriteLine(resultsUserManager.Id + ": " + resultsUserManagerUser.DisplayName + " <" + resultsUserManagerUser.Mail + ">");
            // }

            // Console.WriteLine("\nGraph Request:");
            // Console.WriteLine(requestUserManager.GetHttpRequestMessage().RequestUri);

            var userOptionSelected = ReadUserOption();
            if (userOptionSelected == "1")
            {
                // request 1: create user
                var resultNewUser = CreateUserAsync(client);
                resultNewUser.Wait();
                Console.WriteLine("New user ID: " + resultNewUser.Id);
            }
            else if (userOptionSelected == "2")
            {
                // request 2: update user
                // (1/2) get the user we just created
                var userToUpdate = client.Users.Request()
                                               .Select("id")
                                               .Filter("UserPrincipalName eq 'melissad@8sw0f7.onmicrosoft.com'")
                                               .GetAsync()
                                               .Result[0];
                // (2/2) update the user's phone number
                var resultUpdatedUser = UpdateUserAsync(client, userToUpdate.Id);
                resultUpdatedUser.Wait();
                Console.WriteLine("Updated user ID: " + resultUpdatedUser.Id);
            }
            else if (userOptionSelected == "3")
            {
                var userToUpdate = client.Users.Request()
                                               .Select("id")
                                               .Filter("UserPrincipalName eq 'melissad@8sw0f7.onmicrosoft.com'")
                                               .GetAsync()
                                               .Result[0];
                // request 3: delete user
                var deleteTask = DeleteUserAsync(client, userToUpdate.Id);
                deleteTask.Wait();
            }
            else
            {
                Console.WriteLine("Wrong option selected");
            }
        }

        private static string ReadUserOption()
        {
            string option;
            Console.WriteLine("Select options:\n 1.create \n 2.Update \n 3. Delete");
            option = Console.ReadLine();
            return option;
        }
        private static async Task DeleteUserAsync(GraphServiceClient client, string userIdToDelete)
        {
            await client.Users[userIdToDelete].Request().DeleteAsync();
        }
        private static async Task<Microsoft.Graph.User> UpdateUserAsync(GraphServiceClient client, string userIdToUpdate)
        {
            Microsoft.Graph.User user = new Microsoft.Graph.User()
            {
                MobilePhone = "555-555-1212"
            };
            return await client.Users[userIdToUpdate].Request().UpdateAsync(user);
        }
        private static async Task<Microsoft.Graph.User> CreateUserAsync(GraphServiceClient client)
        {
            Microsoft.Graph.User user = new Microsoft.Graph.User()
            {
                AccountEnabled = true,
                GivenName = "Melissa",
                Surname = "Darrow",
                DisplayName = "Melissa Darrow",
                MailNickname = "MelissaD",
                UserPrincipalName = "melissad@8sw0f7.onmicrosoft.com",
                PasswordProfile = new PasswordProfile()
                {
                    Password = "Password1!",
                    ForceChangePasswordNextSignIn = true
                }
            };
            var requestNewUser = client.Users.Request();
            return await requestNewUser.AddAsync(user);
        }
        private static IConfigurationRoot? LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                                  .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                                  .AddJsonFile("appsettings.json", false, true)
                                  .Build();

                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["tenantId"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }
        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = PublicClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithDefaultRedirectUri()
                                                    .Build();
            return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray());
        }
        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            var graphClient = new GraphServiceClient(authenticationProvider);
            return graphClient;
        }
    }
}

