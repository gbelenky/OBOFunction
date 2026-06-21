using OBOFunction.Models;

namespace OBOFunction.Services;

/// <summary>
/// Resolves the signed-in user's slim profile (name + the existing SharePoint/Graph fields,
/// including country) via On-Behalf-Of, so the proxy can hand the agent enough context to greet
/// the user by name and drive country-filtered features such as <c>search_faq</c>. This is the
/// "Option A" path: it is required because Foundry OAuth identity passthrough cannot deliver the
/// per-user profile through a custom SPFx → proxy → agent chain (see <c>ARCHITECTURE.md</c> §4).
/// <para>
/// The service stays <b>tool-agnostic</b>: it never enumerates, declares, or whitelists agent
/// tools. It returns only profile data; the proxy injects it as conversation context.
/// </para>
/// </summary>
public interface IProfileContextService
{
    /// <summary>
    /// Returns the signed-in user's slim profile (SharePoint UPS via <c>GetMyProperties</c> merged
    /// with Graph <c>/me</c>), or a profile with <see cref="UserProfileContext.HasAny"/> false when
    /// nothing could be resolved. Best-effort: never throws for missing data/consent — the caller
    /// degrades gracefully (e.g. global FAQ results, generic greeting).
    /// </summary>
    /// <param name="userAssertion">The inbound SPFx user token (OBO assertion).</param>
    Task<UserProfileContext> GetProfileAsync(string userAssertion, CancellationToken ct = default);
}
