using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.WebUtilities;

const string CloudGovernanceTokenHeader = "X-Cloud-Governance-Token";
const string InternalKeyHeader = "X-INTERNAL-KEY";
const string FunctionKeyHeader = "x-functions-key";

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddHttpClient("FunctionForwarder", client =>
{
    client.Timeout = TimeSpan.FromMinutes(10);
});

var app = builder.Build();

app.MapGet("/", () => "CAJ Webhook API Running");

app.MapPost("/caj/webhook", async (HttpRequest request, IHttpClientFactory httpClientFactory, ILogger<Program> logger) =>
{
    if (!WebhookConfiguration.TryLoad(out var config, out var configError))
    {
        return Results.Json(
            new { status = "Error", message = configError },
            statusCode: StatusCodes.Status500InternalServerError);
    }

    if (!request.Headers.TryGetValue(CloudGovernanceTokenHeader, out var incomingToken))
    {
        return Unauthorized("Missing X-Cloud-Governance-Token header.");
    }

    if (config is null || !WebhookConfiguration.SecureEquals(incomingToken.ToString(), config.CloudGovernanceToken))
    {
        return Unauthorized("Invalid X-Cloud-Governance-Token.");
    }

    using var reader = new StreamReader(request.Body);
    var body = await reader.ReadToEndAsync();

    if (string.IsNullOrWhiteSpace(body))
    {
        return Results.Json(
            new { status = "BadRequest", message = "Request body is required." },
            statusCode: StatusCodes.Status400BadRequest);
    }

    try
    {
        using var _ = JsonDocument.Parse(body);
    }
    catch (JsonException)
    {
        logger.LogError("Request body must be valid JSON. {Body}", body);
        return Results.Json(
            new { status = "BadRequest", message = "Request body must be valid JSON." },
            statusCode: StatusCodes.Status400BadRequest);
    }

    logger.LogInformation("Received webhook payload: {Payload}", body);

    var httpClient = httpClientFactory.CreateClient("FunctionForwarder");
    using var httpRequest = new HttpRequestMessage(HttpMethod.Post, config.FunctionEndpoint);
    httpRequest.Content = new StringContent(body, Encoding.UTF8, "application/json");
    httpRequest.Headers.TryAddWithoutValidation(InternalKeyHeader, config.FunctionHeaderValue);

    if (!string.IsNullOrWhiteSpace(config.FunctionKey))
    {
        httpRequest.Headers.TryAddWithoutValidation(FunctionKeyHeader, config.FunctionKey);
    }

    using var response = await httpClient.SendAsync(httpRequest);
    var result = await response.Content.ReadAsStringAsync();

    logger.LogInformation("Function responded {StatusCode}: {Result}", (int)response.StatusCode, result);

    return Results.Content(
        result,
        "application/json",
        statusCode: (int)response.StatusCode);
});

app.Run();

static IResult Unauthorized(string message) =>
    Results.Json(new { status = "Unauthorized", message }, statusCode: StatusCodes.Status401Unauthorized);

sealed file class WebhookConfiguration
{
    public required string CloudGovernanceToken { get; init; }
    public required Uri FunctionEndpoint { get; init; }
    public required string FunctionHeaderValue { get; init; }
    public string? FunctionKey { get; init; }

    public static bool TryLoad(out WebhookConfiguration? config, out string? error)
    {
        config = null;
        error = null;

        var cloudGovernanceToken =
            Environment.GetEnvironmentVariable("CLOUD_GOVERNANCE_TOKEN")
            ?? Environment.GetEnvironmentVariable("CAJ_API_KEY");

        var functionUrl = Environment.GetEnvironmentVariable("FUNCTION_URL");
        var functionHeaderValue = Environment.GetEnvironmentVariable("FUNCTION_HEADER_VALUE");
        var functionKey = Environment.GetEnvironmentVariable("FUNCTION_KEY");

        if (string.IsNullOrWhiteSpace(cloudGovernanceToken))
        {
            error = "CLOUD_GOVERNANCE_TOKEN is not configured.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(functionUrl))
        {
            error = "FUNCTION_URL is not configured.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(functionHeaderValue))
        {
            error = "FUNCTION_HEADER_VALUE is not configured.";
            return false;
        }

        if (WebhookConfiguration.SecureEquals(cloudGovernanceToken, functionHeaderValue))
        {
            error = "CLOUD_GOVERNANCE_TOKEN and FUNCTION_HEADER_VALUE must be different values.";
            return false;
        }

        if (!TryResolveFunctionEndpoint(functionUrl, functionKey, out var endpoint, out var resolvedFunctionKey, out var endpointError))
        {
            error = endpointError;
            return false;
        }

        config = new WebhookConfiguration
        {
            CloudGovernanceToken = cloudGovernanceToken,
            FunctionEndpoint = endpoint,
            FunctionHeaderValue = functionHeaderValue,
            FunctionKey = resolvedFunctionKey
        };

        return true;
    }

    public static bool SecureEquals(string? left, string? right)
    {
        if (left is null || right is null)
        {
            return false;
        }

        var leftBytes = Encoding.UTF8.GetBytes(left);
        var rightBytes = Encoding.UTF8.GetBytes(right);

        if (leftBytes.Length != rightBytes.Length)
        {
            return false;
        }

        return CryptographicOperations.FixedTimeEquals(leftBytes, rightBytes);
    }

    private static bool TryResolveFunctionEndpoint(
        string functionUrl,
        string? configuredFunctionKey,
        out Uri endpoint,
        out string? functionKey,
        out string? error)
    {
        endpoint = null!;
        functionKey = configuredFunctionKey;
        error = null;

        if (!Uri.TryCreate(functionUrl, UriKind.Absolute, out var uri))
        {
            error = "FUNCTION_URL is not a valid absolute URL.";
            return false;
        }

        var query = QueryHelpers.ParseQuery(uri.Query);
        if (string.IsNullOrWhiteSpace(functionKey) && query.TryGetValue("code", out var codeValues))
        {
            functionKey = codeValues.FirstOrDefault();
        }

        query.Remove("code");

        var endpointBuilder = new UriBuilder(uri)
        {
            Query = BuildQueryString(query)
        };

        endpoint = endpointBuilder.Uri;

        if (string.IsNullOrWhiteSpace(functionKey))
        {
            error = "FUNCTION_KEY is not configured and FUNCTION_URL does not contain a code query parameter.";
            return false;
        }

        return true;
    }

    private static string BuildQueryString(Dictionary<string, Microsoft.Extensions.Primitives.StringValues> query)
    {
        if (query.Count == 0)
        {
            return string.Empty;
        }

        var pairs = query.SelectMany(
            pair => pair.Value,
            (pair, value) => $"{Uri.EscapeDataString(pair.Key)}={Uri.EscapeDataString(value ?? string.Empty)}");

        return string.Join("&", pairs);
    }
}
