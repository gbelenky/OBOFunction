using System.ComponentModel;
using System.Text.Json;
using Microsoft.Extensions.AI;
using SharePointAgent.Services;

namespace SharePointAgent.Tools;

/// <summary>
/// Embedded agent tool. Exposed to the model as <c>get_sharepoint_profile</c> and,
/// per the agent instructions, invoked once at the very start of every session so the
/// agent always knows who the signed-in user is before answering.
/// </summary>
public sealed class ProfileTools
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    private readonly ISharePointProfileService _profileService;

    public ProfileTools(ISharePointProfileService profileService)
        => _profileService = profileService;

    [Description("Fetch the current signed-in user's Microsoft Graph and SharePoint profile " +
                 "(name, job title, department, office, skills, interests, responsibilities, etc.). " +
                 "Call this once at the start of every conversation before answering, so you know who the user is. " +
                 "Returns a JSON object.")]
    public async Task<string> GetSharePointProfileAsync(CancellationToken cancellationToken = default)
    {
        var profile = await _profileService.GetMyProfileAsync(cancellationToken).ConfigureAwait(false);
        return JsonSerializer.Serialize(profile, JsonOptions);
    }

    /// <summary>Builds the AIFunction the agent exposes, with a stable tool name.</summary>
    public AIFunction CreateTool() =>
        AIFunctionFactory.Create(
            GetSharePointProfileAsync,
            new AIFunctionFactoryOptions { Name = "get_sharepoint_profile" });
}
