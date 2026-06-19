using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using OBOFunction.Auth;
using OBOFunction.Services;

var host = new HostBuilder()
    .ConfigureFunctionsWebApplication()
    .ConfigureAppConfiguration((ctx, cfg) =>
    {
        var built = cfg.Build();
        var kvUri = built["KeyVault:Uri"];
        if (!string.IsNullOrWhiteSpace(kvUri))
        {
            cfg.AddAzureKeyVault(new Uri(kvUri), new DefaultAzureCredential());
        }
    })
    .ConfigureServices((ctx, services) =>
    {
        services
            .AddApplicationInsightsTelemetryWorkerService()
            .ConfigureFunctionsApplicationInsights();

        services.Configure<LoggerFilterOptions>(o =>
        {
            var rule = o.Rules.FirstOrDefault(r =>
                r.ProviderName == "Microsoft.Extensions.Logging.ApplicationInsights.ApplicationInsightsLoggerProvider");
            if (rule is not null) o.Rules.Remove(rule);
        });

        services
            .AddMicrosoftIdentityWebApiAuthentication(ctx.Configuration, "AzureAd")
            .EnableTokenAcquisitionToCallDownstreamApi()
            .AddInMemoryTokenCaches();

        services.AddScoped<IUserTokenAccessor, UserTokenAccessor>();
        services.AddScoped<IGraphProfileService, GraphProfileService>();
        services.AddSingleton<IAgentChatClient, AgentChatClient>();
    })
    .Build();

host.Run();
