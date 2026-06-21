namespace OBOFunction.Models;

/// <summary>
/// The slim signed-in-user profile the proxy resolves via OBO (Option A) and injects as
/// conversation context so the hosted agent can greet the user by name and use the country for
/// country-filtered features (e.g. <c>search_faq</c>). Built from SharePoint UPS (<c>GetMyProperties</c>)
/// merged with Graph <c>/me</c>.
/// <para>
/// Only non-personal-bulk fields the agent actually needs are carried. All fields are optional;
/// any that cannot be resolved are simply omitted from the injected context.
/// </para>
/// </summary>
public sealed record UserProfileContext
{
    /// <summary>Graph <c>displayName</c> / UPS <c>PreferredName</c>.</summary>
    public string? Name { get; init; }

    /// <summary>Graph <c>givenName</c> — preferred for a first-name greeting.</summary>
    public string? FirstName { get; init; }

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

    /// <summary>Custom UPS attribute <c>IntranetCountry</c> (falls back to Graph country).</summary>
    public string? Country { get; init; }

    /// <summary>True when at least a name or country could be resolved for the signed-in user.</summary>
    public bool HasAny =>
        !string.IsNullOrWhiteSpace(Name)
        || !string.IsNullOrWhiteSpace(FirstName)
        || !string.IsNullOrWhiteSpace(Country)
        || !string.IsNullOrWhiteSpace(JobTitle)
        || Responsibilities.Count > 0
        || PastProjects.Count > 0
        || Interests.Count > 0;
}
