using OBOFunction.Models;

namespace OBOFunction.Services;

public interface IGraphProfileService
{
    Task<UserProfile> GetMyProfileAsync(string userAssertion, CancellationToken ct);
}
