using System.ComponentModel;
using SharePointAgent.Services;

namespace SharePointAgent.Tools;

/// <summary>
/// In-process agent tool surface for the FAQ / Q&amp;A knowledge base. Wraps
/// <see cref="FaqSearchService"/> and exposes a single delegate the agent registers via
/// <c>AIFunctionFactory.Create</c>. The <see cref="DescriptionAttribute"/> annotations become the
/// function/parameter descriptions the model reasons over, so it knows to pass the user's country.
/// </summary>
public sealed class FaqTools
{
    private readonly FaqSearchService _service;

    public FaqTools(FaqSearchService service) => _service = service;

    [Description(
        "Search the corporate FAQ / Q&A knowledge base and return entries relevant to the user. " +
        "Results are filtered to the user's country (the SharePoint 'Location' field) plus " +
        "globally-applicable entries. Call this to answer policy, IT, HR, finance, facilities or " +
        "general how-to questions. Pass the country taken from the user's profile, and answer ONLY " +
        "from the returned entries — cite the FAQ Title and do not invent answers.")]
    public Task<string> SearchFaqAsync(
        [Description(
            "The user's country or region from their SharePoint profile (e.g. 'Europe', " +
            "'North America', 'Latin America'). Filters FAQs by the index 'Location' field; " +
            "globally-applicable entries are always included. Pass an empty string to get only " +
            "globally-applicable FAQs.")]
        string country,
        [Description(
            "The user's question or keywords to match against FAQ titles, questions and answers. " +
            "Pass '*' to list all FAQs applicable to the country.")]
        string question,
        CancellationToken cancellationToken = default)
        => _service.SearchAsync(country, question, top: 5, cancellationToken);
}
