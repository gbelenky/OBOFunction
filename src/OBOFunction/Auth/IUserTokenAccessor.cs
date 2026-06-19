using Microsoft.AspNetCore.Http;

namespace OBOFunction.Auth;

public interface IUserTokenAccessor
{
    string GetBearerToken(HttpRequest request);
}
