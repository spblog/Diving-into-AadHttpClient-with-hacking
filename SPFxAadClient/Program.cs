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
            Console.WriteLine("[a] - add\\update permission \n[d] - deletes permission\nPlease enter your choice:");

            ConsoleKeyInfo key = Console.ReadKey();
            switch (key.KeyChar)
            {
                case 'a':
                    AddSpfxMSGraphPermissions(client, "Calendars.Read,Mail.ReadWrite").Wait();
                    Console.WriteLine("Successfully added permissions for MS Graph!");
                    break;
                case 'd':
                    RemoveSpfxMSGraphPermissions(client).Wait();
                    Console.WriteLine("Successfully deleted permissions for MS Graph!");
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
    }
}
