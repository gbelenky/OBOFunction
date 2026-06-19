using Microsoft.AspNetCore.Http;

namespace OBOFunction.Auth;

public sealed class UserTokenAccessor : IUserTokenAccessor
{
    public string GetBearerToken(HttpRequest request)
    {
        if (!request.Headers.TryGetValue("Authorization", out var header))
            throw new UnauthorizedAccessException("Missing Authorization header.");

        var value = header.ToString();
        const string prefix = "Bearer ";
        if (!value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            throw new UnauthorizedAccessException("Authorization header must use the Bearer scheme.");

        var token = value[prefix.Length..].Trim();
        if (string.IsNullOrEmpty(token))
            throw new UnauthorizedAccessException("Bearer token is empty.");

        return token;
    }
}
