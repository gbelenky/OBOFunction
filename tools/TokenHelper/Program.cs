using Microsoft.Identity.Client;

// Interactive (browser) token minter for local testing of GET /api/profile.
//
// Pops a system browser, signs you in, and prints a bearer token whose audience is the
// OBO Function API (api://<clientId>) with the delegated scope access_as_user. This avoids
// device-code flow entirely (Entra-issue friendly) and never needs the client secret.
//
// Prereqs (already configured by setup):
//   - The app registration has a public-client redirect URI http://localhost.
//   - You have access to sign in to the tenant interactively.
//
// Usage (defaults target this project's OBO app):
//   dotnet run --project tools\TokenHelper
//   dotnet run --project tools\TokenHelper -- <tenantId> <clientId>

string tenantId = args.Length > 0 ? args[0] : "cbe03044-c23b-46df-93a5-c018d51915d8";
string clientId = args.Length > 1 ? args[1] : "7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a";

// Delegated scope exposed by the OBO API ("Expose an API" -> access_as_user).
string scope = $"api://{clientId}/access_as_user";

var app = PublicClientApplicationBuilder
    .Create(clientId)
    .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
    .WithRedirectUri("http://localhost")
    .Build();

AuthenticationResult result;
try
{
    result = await app.AcquireTokenInteractive(new[] { scope })
        .WithPrompt(Prompt.SelectAccount)
        .ExecuteAsync();
}
catch (MsalException ex)
{
    Console.Error.WriteLine($"Interactive sign-in failed: {ex.ErrorCode} - {ex.Message}");
    return 1;
}

Console.WriteLine();
Console.WriteLine($"Signed in as : {result.Account.Username}");
Console.WriteLine($"Audience     : api://{clientId}");
Console.WriteLine($"Scope        : access_as_user");
Console.WriteLine($"Expires (UTC): {result.ExpiresOn.UtcDateTime:O}");
Console.WriteLine();
Console.WriteLine("----- ACCESS TOKEN -----");
Console.WriteLine(result.AccessToken);
Console.WriteLine("------------------------");
Console.WriteLine();
Console.WriteLine("Quick test against a locally running Function host (func start):");
Console.WriteLine();
Console.WriteLine("  $t = '<paste-token>'");
Console.WriteLine("  curl http://localhost:7071/api/profile -H \"Authorization: Bearer $t\"");
Console.WriteLine();
return 0;
