namespace OBOFunction.Services;

/// <summary>
/// Resolves the signed-in user's country/region via On-Behalf-Of, so the proxy can hand the
/// agent a single context value (the country) for country-filtered features such as
/// <c>search_faq</c>. This is the "Option A" path: it is required because Foundry OAuth identity
/// passthrough cannot deliver the per-user profile through a custom SPFx → proxy → agent chain
/// (see <c>docs/foundry-oauth-passthrough-findings.md</c>).
/// <para>
/// The service stays <b>tool-agnostic</b>: it never enumerates, declares, or whitelists agent
/// tools. It returns only a plain country string; the proxy injects it as conversation context.
/// </para>
/// </summary>
public interface IProfileCountryService
{
    /// <summary>
    /// Returns the signed-in user's country (UPS <c>IntranetCountry</c>, falling back to Graph
    /// <c>country</c>/<c>usageLocation</c>), or <c>null</c> when it cannot be resolved. Best-effort:
    /// never throws for missing data/consent — the caller degrades gracefully (global results).
    /// </summary>
    /// <param name="userAssertion">The inbound SPFx user token (OBO assertion).</param>
    Task<string?> GetCountryAsync(string userAssertion, CancellationToken ct = default);
}
