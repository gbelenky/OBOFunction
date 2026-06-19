namespace SharePointAgent.Models;

/// <summary>
/// Merged Microsoft Graph + SharePoint User Profile Service (UPS) view of the
/// signed-in user. Returned by the embedded <c>get_sharepoint_profile</c> tool.
/// </summary>
public sealed record UserProfile
{
    public string? Id { get; init; }
    public string? DisplayName { get; init; }
    public string? GivenName { get; init; }
    public string? Surname { get; init; }
    public string? UserPrincipalName { get; init; }
    public string? Mail { get; init; }
    public string? JobTitle { get; init; }
    public string? Department { get; init; }
    public string? OfficeLocation { get; init; }
    public string? PreferredLanguage { get; init; }
    public string? MobilePhone { get; init; }
    public IReadOnlyList<string> BusinessPhones { get; init; } = [];
    public SharePointProfile? SharePointProfile { get; init; }
}

/// <summary>
/// SharePoint User Profile Service properties (PeopleManager/GetMyProperties).
/// </summary>
public sealed record SharePointProfile
{
    public string? AccountName { get; init; }
    public string? AboutMe { get; init; }
    public string? PersonalUrl { get; init; }
    public IReadOnlyList<string> Skills { get; init; } = [];
    public IReadOnlyList<string> Interests { get; init; } = [];
    public IReadOnlyList<string> PastProjects { get; init; } = [];
    public IReadOnlyList<string> Responsibilities { get; init; } = [];
    public IReadOnlyList<string> Schools { get; init; } = [];
}
