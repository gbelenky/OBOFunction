namespace SharePointMcp.Models;

/// <summary>
/// Slim profile returned by the <c>get_sharepoint_profile</c> MCP tool. Only the fields the
/// agent needs are delivered, so the payload contains no unused/null properties.
/// <list type="bullet">
/// <item><description><see cref="Name"/> / <see cref="Email"/> come from Microsoft Graph (<c>/me</c>).</description></item>
/// <item><description>The remaining fields come from the SharePoint User Profile Service (UPS)
/// via <c>PeopleManager/GetMyProperties</c>.</description></item>
/// </list>
/// </summary>
public sealed record UserProfile
{
    /// <summary>Graph <c>displayName</c>.</summary>
    public string? Name { get; init; }

    /// <summary>Graph <c>givenName</c>.</summary>
    public string? FirstName { get; init; }

    /// <summary>Graph <c>surname</c>.</summary>
    public string? LastName { get; init; }

    /// <summary>Graph <c>mail</c> (falls back to <c>userPrincipalName</c>).</summary>
    public string? Email { get; init; }

    /// <summary>UPS <c>SPS-JobTitle</c> (falls back to Graph <c>jobTitle</c>).</summary>
    public string? JobTitle { get; init; }

    /// <summary>UPS <c>SPS-Responsibility</c>.</summary>
    public IReadOnlyList<string> Responsibilities { get; init; } = [];

    /// <summary>UPS <c>SPS-PastProjects</c>.</summary>
    public IReadOnlyList<string> PastProjects { get; init; } = [];

    /// <summary>UPS <c>SPS-Interests</c>.</summary>
    public IReadOnlyList<string> Interests { get; init; } = [];

    /// <summary>Custom UPS attribute <c>IntranetCountry</c>.</summary>
    public string? Country { get; init; }
}

