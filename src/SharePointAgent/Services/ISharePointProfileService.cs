using SharePointAgent.Models;

namespace SharePointAgent.Services;

/// <summary>
/// Fetches the signed-in user's merged Graph + SharePoint profile.
/// </summary>
public interface ISharePointProfileService
{
    Task<UserProfile> GetMyProfileAsync(CancellationToken ct = default);
}
