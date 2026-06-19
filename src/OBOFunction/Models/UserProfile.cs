namespace OBOFunction.Models;

public sealed record UserProfile(
    string Id,
    string DisplayName,
    string? GivenName,
    string? Surname,
    string? UserPrincipalName,
    string? Mail,
    string? JobTitle,
    string? Department,
    string? OfficeLocation,
    string? PreferredLanguage,
    string? MobilePhone,
    IReadOnlyList<string> BusinessPhones,
    string? PhotoBase64,
    string? PhotoContentType,
    SharePointProfile? SharePointProfile);

public sealed record SharePointProfile(
    string AccountName,
    string? AboutMe,
    string? PictureUrl,
    string? PersonalUrl,
    IReadOnlyList<string> Skills,
    IReadOnlyList<string> Interests,
    IReadOnlyList<string> PastProjects,
    IReadOnlyList<string> Responsibilities,
    IReadOnlyList<string> Schools,
    IReadOnlyDictionary<string, string> ExtendedProperties);
