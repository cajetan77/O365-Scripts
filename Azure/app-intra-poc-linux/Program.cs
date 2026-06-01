using Azure.Identity;
using Azure.ResourceManager;
using Azure.ResourceManager.Automation;
using Azure.ResourceManager.Automation.Models;
using System.Text.Json;
using Microsoft.AspNetCore.Http;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

app.UseHttpsRedirection();

// --- CRITICAL CONFIGURATION CONFIG ---
 string ? ExpectedBearerToken = Environment.GetEnvironmentVariable("EXPECTED_BEARER_TOKEN"); // Change this to a strong password/secret
 string ? SubscriptionId     = Environment.GetEnvironmentVariable("SUBSCRIPTION_ID");
 string ? ResourceGroupName  = Environment.GetEnvironmentVariable("RESOURCE_GROUP_NAME"); 
     string ? AutomationAccount  = Environment.GetEnvironmentVariable("AUTOMATION_ACCOUNT");
 string ? RunbookName        = Environment.GetEnvironmentVariable("RUNBOOK_NAME");
// -------------------------------------

app.MapPost("/process-automation", async (HttpContext context, JsonElement dynamicData) =>
{
    // 1. TOKEN VALIDATION
    if (!context.Request.Headers.TryGetValue("Authorization", out var authHeader))
    {
        return Results.Json(new { error = "Authorization header is missing" }, statusCode: 401);
    }

    string headerValue = authHeader.ToString();
    if (!headerValue.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
    {
        return Results.Json(new { error = "Authorization header must use Bearer scheme" }, statusCode: 401);
    }

    // Extract the token string and compare it
    string providedToken = headerValue.Substring("Bearer ".Length).Trim();
    if (providedToken != ExpectedBearerToken)
    {
        return Results.Json(new { error = "Invalid token. Access Denied." }, statusCode: 403);
    }

    // 2. PROCESS DYNAMIC PAYLOAD (Runs only if token is valid)
    string rawJson = dynamicData.GetRawText();
    Console.WriteLine($"[SPO WEB APP] Securely received JSON Payload: {rawJson}");

    try
    {
        var credential = new DefaultAzureCredential();
        var armClient = new ArmClient(credential);

        var automationAccountResourceId = AutomationAccountResource.CreateResourceIdentifier(SubscriptionId, ResourceGroupName, AutomationAccount);
        var automationAccountResource = armClient.GetAutomationAccountResource(automationAccountResourceId);
        var jobCollection = automationAccountResource.GetAutomationJobs();
        
        var jobParameters = new AutomationJobCreateOrUpdateContent
        {
            RunbookName = RunbookName,
            Parameters = 
            {
                { "RequestBody", rawJson } 
            }
        };

        string jobName = Guid.NewGuid().ToString(); 
        await jobCollection.CreateOrUpdateAsync(Azure.WaitUntil.Completed, jobName, jobParameters);

        return Results.Ok(new { 
            Status = "Authenticated and dispatched to Azure Automation", 
            JobGuid = jobName
        });
    }
    catch (Exception ex)
    {
        return Results.Problem(detail: ex.Message, statusCode: 500, title: "Automation Dispatch Failed");
    }
});

app.Run();
