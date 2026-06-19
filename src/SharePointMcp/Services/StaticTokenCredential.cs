using Azure.Core;

namespace SharePointMcp.Services;

/// <summary>
/// Wraps a pre-issued bearer token (the user-delegated token injected by Foundry's
/// OAuth connection layer, or a token supplied directly for testing) as a
/// <see cref="TokenCredential"/> so the Graph SDK can consume it unchanged.
/// </summary>
/// <remarks>
/// The wrapped token has a FIXED audience (whatever resource it was minted for —
/// typically Microsoft Graph). It cannot be exchanged for a different audience, so
/// callers must not request a non-matching scope (e.g. a SharePoint-audience token).
/// </remarks>
public sealed class StaticTokenCredential : TokenCredential
{
    private readonly string _token;
    private readonly DateTimeOffset _expiresOn;

    public StaticTokenCredential(string token, DateTimeOffset? expiresOn = null)
    {
        _token = token ?? throw new ArgumentNullException(nameof(token));
        // We don't decode the JWT exp here; assume short validity to force refresh upstream.
        _expiresOn = expiresOn ?? DateTimeOffset.UtcNow.AddMinutes(5);
    }

    public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        => new(_token, _expiresOn);

    public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        => new(new AccessToken(_token, _expiresOn));
}
