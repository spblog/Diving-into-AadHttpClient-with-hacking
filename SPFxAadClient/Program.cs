using System;
using System.Threading.Tasks;
using Microsoft.Azure.ActiveDirectory.GraphClient;

namespace SPFxAadClient
{
    class Program
    {
        static void Main(string[] args)
        {
            var client = AuthenticationHelper.GetActiveDirectoryClientAsUser();

            Console.WriteLine("Run operations for signed-in user");
            Console.WriteLine("[a] - add\\update permission Graph permissions \n[d] - deletes Graph permissions \n[c] - deletes api-sso permissions \n[d] - adds api-sso permissions \n[e] - prints all grants added to SharePoint CE web app \nPlease enter your choice:");

            ConsoleKeyInfo key = Console.ReadKey();
            switch (key.KeyChar)
            {
                case 'a':
                    AddSpfxMSGraphPermissions(client, "Calendars.Read,Mail.ReadWrite").Wait();
                    Console.WriteLine("Successfully added permissions for MS Graph!");
                    break;
                case 'b':
                    RemoveSpfxMSGraphPermissions(client).Wait();
                    Console.WriteLine("Successfully deleted permissions for MS Graph!");
                    break;
                case 'c':
                    RemoveThirdPartAPIPermissions(client, "6fc2655e-04cd-437d-a50d-0c1a31383775").Wait();
                    Console.WriteLine("Successfully deleted permissions for api-sso!");
                    break;
                case 'd':
                    AddThirdPartAPIPermissions(client, "6fc2655e-04cd-437d-a50d-0c1a31383775", "user_impersonation").Wait();
                    Console.WriteLine("Successfully added permissions for api-sso!");
                    break;
                case 'e':
                    PrintGrants(client).Wait();
                    break;
            }
        }

        public static async Task AddSpfxMSGraphPermissions(ActiveDirectoryClient client, string scope)
        {
            var spoCeServicePrinicipal = await client.ServicePrincipals
                .Where(sp => sp.AppId == GlobalConstants.SpoCeApplicationId).ExecuteSingleAsync();

            var msGraphServicePrincipal = await client.ServicePrincipals
                .Where(sp => sp.AppId == GlobalConstants.MsGrpaphApplicationId).ExecuteSingleAsync();

            var grants = await client.Oauth2PermissionGrants.ExecuteAsync();
            OAuth2PermissionGrant existingGrant = null;
            foreach (IOAuth2PermissionGrant grant in grants.CurrentPage)
            {
                if (grant.ClientId == spoCeServicePrinicipal.ObjectId &&
                    grant.ResourceId == msGraphServicePrincipal.ObjectId)
                {
                    existingGrant = (OAuth2PermissionGrant) grant;
                }
            }

            if (existingGrant != null)
            {
                existingGrant.Scope = scope;
                await existingGrant.UpdateAsync();
            }
            else
            {
                var auth2PermissionGrant = new OAuth2PermissionGrant
                {
                    ClientId = spoCeServicePrinicipal.ObjectId,
                    ConsentType = "AllPrincipals",
                    PrincipalId = null,
                    ExpiryTime = DateTime.Now.AddYears(10),
                    ResourceId = msGraphServicePrincipal.ObjectId,
                    Scope = scope
                };

                await client.Oauth2PermissionGrants.AddOAuth2PermissionGrantAsync(auth2PermissionGrant);
            }
        }

        public static async Task RemoveSpfxMSGraphPermissions(ActiveDirectoryClient client)
        {
            var spoCeServicePrinicipal = await client.ServicePrincipals
                .Where(sp => sp.AppId == GlobalConstants.SpoCeApplicationId).ExecuteSingleAsync();

            var msGraphServicePrincipal = await client.ServicePrincipals
                .Where(sp => sp.AppId == GlobalConstants.MsGrpaphApplicationId).ExecuteSingleAsync();

            var grants = await client.Oauth2PermissionGrants.ExecuteAsync();
            OAuth2PermissionGrant existingGrant = null;
            foreach (IOAuth2PermissionGrant grant in grants.CurrentPage)
            {
                if (grant.ClientId == spoCeServicePrinicipal.ObjectId &&
                    grant.ResourceId == msGraphServicePrincipal.ObjectId)
                {
                    existingGrant = (OAuth2PermissionGrant)grant;
                }
            }

            if (existingGrant != null)
            {
                await existingGrant.DeleteAsync();
            }
        }

        public static async Task RemoveThirdPartAPIPermissions(ActiveDirectoryClient client, string thirdPartyClientId)
        {
            var spoCeServicePrinicipal = await client.ServicePrincipals
                .Where(sp => sp.AppId == GlobalConstants.SpoCeApplicationId).ExecuteSingleAsync();

            var thirdPartyServicePrincipal = await client.ServicePrincipals
                .Where(sp => sp.AppId == thirdPartyClientId).ExecuteSingleAsync();

            var grants = await client.Oauth2PermissionGrants.ExecuteAsync();
            OAuth2PermissionGrant existingGrant = null;
            foreach (IOAuth2PermissionGrant grant in grants.CurrentPage)
            {
                if (grant.ClientId == spoCeServicePrinicipal.ObjectId &&
                    grant.ResourceId == thirdPartyServicePrincipal.ObjectId)
                {
                    existingGrant = (OAuth2PermissionGrant)grant;
                }
            }

            if (existingGrant != null)
            {
                await existingGrant.DeleteAsync();
            }
        }

        public static async Task AddThirdPartAPIPermissions(ActiveDirectoryClient client, string thirdPartyClientId, string scope)
        {
            var spoCeServicePrinicipal = await client.ServicePrincipals
                .Where(sp => sp.AppId == GlobalConstants.SpoCeApplicationId).ExecuteSingleAsync();

            var thirdPartyServicePrincipal = await client.ServicePrincipals
                .Where(sp => sp.AppId == thirdPartyClientId).ExecuteSingleAsync();

            var grants = await client.Oauth2PermissionGrants.ExecuteAsync();
            OAuth2PermissionGrant existingGrant = null;
            foreach (IOAuth2PermissionGrant grant in grants.CurrentPage)
            {
                if (grant.ClientId == spoCeServicePrinicipal.ObjectId &&
                    grant.ResourceId == thirdPartyServicePrincipal.ObjectId)
                {
                    existingGrant = (OAuth2PermissionGrant)grant;
                }
            }

            if (existingGrant != null)
            {
                existingGrant.Scope = scope;
                await existingGrant.UpdateAsync();
            }
            else
            {
                var auth2PermissionGrant = new OAuth2PermissionGrant
                {
                    ClientId = spoCeServicePrinicipal.ObjectId,
                    ConsentType = "AllPrincipals",
                    PrincipalId = null,
                    ExpiryTime = DateTime.Now.AddYears(10),
                    ResourceId = thirdPartyServicePrincipal.ObjectId,
                    Scope = scope
                };

                await client.Oauth2PermissionGrants.AddOAuth2PermissionGrantAsync(auth2PermissionGrant);
            }
        }

        public static async Task PrintGrants(ActiveDirectoryClient client)
        {
            var spoCeServicePrinicipal = await client.ServicePrincipals
                .Where(sp => sp.AppId == GlobalConstants.SpoCeApplicationId).ExecuteSingleAsync();

            var grants = await client.Oauth2PermissionGrants.ExecuteAsync();
            foreach (IOAuth2PermissionGrant grant in grants.CurrentPage)
            {
                if (grant.ClientId == spoCeServicePrinicipal.ObjectId)
                {
                    var objectId = grant.ResourceId;
                    var servicePrincipal = await client.ServicePrincipals.Where(sp => sp.ObjectId == objectId).ExecuteSingleAsync();

                    Console.WriteLine("*********************");
                    Console.WriteLine("");
                    Console.WriteLine($"AppDisplayName: {servicePrincipal.AppDisplayName}");
                    Console.WriteLine($"ConsentType: {grant.ConsentType}");
                    Console.WriteLine($"Scope: {grant.Scope}");
                    Console.WriteLine($"ClientId: {grant.ClientId}");
                    Console.WriteLine($"ResourceId: {grant.ResourceId}");
                    Console.WriteLine("");
                }
            }
        }
    }
}
