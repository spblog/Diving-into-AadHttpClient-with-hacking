## Diving into AadHttpClient (with hacking!)

Blog post - [Diving into AadHttpClient (with hacking!)](http://spblog.net/post/2018/08/07/Diving-into-AadHttpClient-(with-hacking!))

### How to run
1. In Azure AD register new application (native)
2. For required permissions add "Windows Azure Active Directory" -> "Access directory as the signed-in user"
3. Go to Constants.cs and change TenantId to your tenant id and ClientId from step 1