namespace SPFxAadClient
{
    internal class UserModeConstants
    {
        public const string AuthString = GlobalConstants.AuthString + "common/";
    }

    internal class GlobalConstants
    {
        public const string AuthString = "https://login.microsoftonline.com/";        
        public const string ResourceUrl = "https://graph.windows.net";
        public const string GraphServiceObjectId = "00000002-0000-0000-c000-000000000000";
        public const string TenantId = "948fd9cc-9adc-40d8-851e-acefa17ab66c"; // <--- change to your TenantId
        public const string ClientId = "dd1155f9-3f40-4190-b5e4-3988a6c18250"; // <--- change to your ClientId
        public const string SpoCeApplicationId = "d02ebf4e-0497-4b18-817d-06ad82d41540"; // <-- SharePoint Online Client Extensibility Web Application Principal Id (might be different in your tenant)
        public const string MsGrpaphApplicationId = "00000003-0000-0000-c000-000000000000";
    }
}
