using System.Text.Json;
using System.Text.Json.Serialization;
using Azure;
using Azure.Identity;
using Azure.Search.Documents;
using Azure.Search.Documents.Models;

namespace SharePointAgent.Services;

/// <summary>
/// Queries the Azure AI Search <c>faq-index</c> for FAQ / Q&amp;A entries relevant to the
/// signed-in user. This is a <b>local, in-process</b> tool on the agent — it needs no OBO and
/// no user identity: it only requires the user's <em>country</em> (a plain string sourced from
/// the SharePoint profile) to build a server-side filter.
///
/// <para>
/// The index has no dedicated country column; the filterable <c>Location</c> field is used as the
/// country/region (per design). Globally-applicable entries (<c>Location = "Global"</c>) are
/// always included so a user whose exact country has no region-specific entries still gets answers.
/// </para>
///
/// <para>
/// Auth is the agent's <b>own</b> identity, never the user's: a read-only query API key when one
/// is configured, otherwise <see cref="DefaultAzureCredential"/> (Managed Identity in Azure, the
/// developer credential locally — requires the <c>Search Index Data Reader</c> role on the service).
/// </para>
/// </summary>
public sealed class FaqSearchService
{
    private static readonly string[] SelectFields =
        ["Id", "Title", "Question", "Answer", "Category", "Language", "Location", "Department", "ItemUrl"];

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly SearchClient _client;
    private readonly string _countryField;
    private readonly bool _includeGlobal;

    public FaqSearchService(
        string endpoint,
        string indexName,
        string countryField,
        string? apiKey,
        bool includeGlobal = true)
    {
        _countryField = string.IsNullOrWhiteSpace(countryField) ? "Location" : countryField;
        _includeGlobal = includeGlobal;

        _client = string.IsNullOrWhiteSpace(apiKey)
            ? new SearchClient(new Uri(endpoint), indexName, new DefaultAzureCredential())
            : new SearchClient(new Uri(endpoint), indexName, new AzureKeyCredential(apiKey));
    }

    /// <summary>
    /// Runs a full-text search over the FAQ index, filtered to the user's country (plus Global),
    /// and returns a compact JSON document the agent can ground its answer on.
    /// </summary>
    public async Task<string> SearchAsync(
        string? country, string? question, int top = 5, CancellationToken ct = default)
    {
        var options = new SearchOptions
        {
            Size = top <= 0 ? 5 : Math.Min(top, 20),
            IncludeTotalCount = true
        };
        foreach (var f in SelectFields)
            options.Select.Add(f);

        var filter = BuildFilter(country);
        if (filter is not null)
            options.Filter = filter;

        var searchText = string.IsNullOrWhiteSpace(question) ? "*" : question;

        Response<SearchResults<SearchDocument>> resp =
            await _client.SearchAsync<SearchDocument>(searchText, options, ct).ConfigureAwait(false);

        var items = new List<FaqHit>();
        await foreach (SearchResult<SearchDocument> r in resp.Value.GetResultsAsync().WithCancellation(ct))
        {
            var d = r.Document;
            items.Add(new FaqHit
            {
                Title = Get(d, "Title"),
                Question = Get(d, "Question"),
                Answer = Get(d, "Answer"),
                Category = Get(d, "Category"),
                Language = Get(d, "Language"),
                Location = Get(d, "Location"),
                Department = Get(d, "Department"),
                Url = Get(d, "ItemUrl"),
                Score = r.Score
            });
        }

        // Keyword search can miss entries written in the user's local language (e.g. an English
        // query "vacation" will not match a German "Urlaubsantrag" entry). When a real query
        // returns nothing, broaden to the full country (+Global) set so the multilingual model can
        // map the question to the right entry across languages instead of falsely reporting "none".
        bool broadened = false;
        if (items.Count == 0 && !string.IsNullOrWhiteSpace(question) && searchText != "*")
        {
            // Use a generous page size here: an unscored "*" match returns rows in arbitrary order,
            // so a small Size could drop the very (region-specific) entries we are trying to surface.
            var broadOptions = new SearchOptions { Size = 50, IncludeTotalCount = true };
            foreach (var f in SelectFields)
                broadOptions.Select.Add(f);
            if (filter is not null)
                broadOptions.Filter = filter;

            Response<SearchResults<SearchDocument>> fallback =
                await _client.SearchAsync<SearchDocument>("*", broadOptions, ct).ConfigureAwait(false);

            await foreach (SearchResult<SearchDocument> r in fallback.Value.GetResultsAsync().WithCancellation(ct))
            {
                var d = r.Document;
                items.Add(new FaqHit
                {
                    Title = Get(d, "Title"),
                    Question = Get(d, "Question"),
                    Answer = Get(d, "Answer"),
                    Category = Get(d, "Category"),
                    Language = Get(d, "Language"),
                    Location = Get(d, "Location"),
                    Department = Get(d, "Department"),
                    Url = Get(d, "ItemUrl"),
                    Score = r.Score
                });
            }
            broadened = items.Count > 0;
        }

        var payload = new FaqSearchResult
        {
            Country = country,
            CountryField = _countryField,
            Filter = filter,
            IncludeGlobal = _includeGlobal,
            Broadened = broadened,
            Count = items.Count,
            Results = items
        };

        return JsonSerializer.Serialize(payload, JsonOptions);
    }

    /// <summary>
    /// Builds the OData filter. The user's country maps to the configured field (default
    /// <c>Location</c>); single quotes are escaped per OData rules. Global entries are OR-ed in
    /// when enabled so universally-applicable FAQs always surface.
    /// </summary>
    internal string? BuildFilter(string? country)
    {
        if (string.IsNullOrWhiteSpace(country))
            return _includeGlobal ? $"{_countryField} eq 'Global'" : null;

        var escaped = country.Trim().Replace("'", "''");
        return _includeGlobal
            ? $"{_countryField} eq '{escaped}' or {_countryField} eq 'Global'"
            : $"{_countryField} eq '{escaped}'";
    }

    private static string? Get(SearchDocument d, string key) =>
        d.TryGetValue(key, out var v) ? v?.ToString() : null;

    private sealed record FaqHit
    {
        public string? Title { get; init; }
        public string? Question { get; init; }
        public string? Answer { get; init; }
        public string? Category { get; init; }
        public string? Language { get; init; }
        public string? Location { get; init; }
        public string? Department { get; init; }
        public string? Url { get; init; }
        public double? Score { get; init; }
    }

    private sealed record FaqSearchResult
    {
        public string? Country { get; init; }
        public string? CountryField { get; init; }
        public string? Filter { get; init; }
        public bool IncludeGlobal { get; init; }
        public bool Broadened { get; init; }
        public int Count { get; init; }
        public IReadOnlyList<FaqHit> Results { get; init; } = [];
    }
}
