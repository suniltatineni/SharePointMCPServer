using System.Text.Json;
using Microsoft.Graph;
using Azure.Identity;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

// MCP protocol endpoints
app.MapPost("/mcp/v1/initialize", async (HttpContext context) =>
{
    var response = new
    {
        protocolVersion = "1.0",
        serverInfo = new
        {
            name = "sharepoint-mcp-server",
            version = "1.0.0"
        },
        capabilities = new
        {
            tools = new { }
        }
    };
    
    await context.Response.WriteAsJsonAsync(response);
});

app.MapPost("/mcp/v1/tools/list", async (HttpContext context) =>
{
    var tools = new object[]
    {
        new
        {
            name = "search_list_items",
            description = "Search for items in a SharePoint Online list. Returns matching list items based on the search query.",
            inputSchema = new
            {
                type = "object",
                properties = new
                {
                    siteUrl = new
                    {
                        type = "string",
                        description = "SharePoint site URL (e.g., https://contoso.sharepoint.com/sites/mysite)"
                    },
                    listTitle = new
                    {
                        type = "string",
                        description = "Title or name of the SharePoint list"
                    },
                    searchQuery = new
                    {
                        type = "string",
                        description = "Search query to filter list items"
                    },
                    maxResults = new
                    {
                        type = "number",
                        description = "Maximum number of results to return (default: 50)",
                        @default = 50
                    }
                },
                required = new[] { "siteUrl", "listTitle", "searchQuery" }
            }
        },
        new
        {
            name = "get_list_items",
            description = "Get all items from a SharePoint Online list with optional filtering.",
            inputSchema = new
            {
                type = "object",
                properties = new
                {
                    siteUrl = new
                    {
                        type = "string",
                        description = "SharePoint site URL"
                    },
                    listTitle = new
                    {
                        type = "string",
                        description = "Title or name of the SharePoint list"
                    },
                    filter = new
                    {
                        type = "string",
                        description = "OData filter query (optional)"
                    },
                    maxResults = new
                    {
                        type = "number",
                        description = "Maximum number of results (default: 100)",
                        @default = 100
                    }
                },
                required = new[] { "siteUrl", "listTitle" }
            }
        }
    };
    
    await context.Response.WriteAsJsonAsync(new { tools });
});

app.MapPost("/mcp/v1/tools/call", async (HttpContext context) =>
{
    try
    {
        var request = await JsonSerializer.DeserializeAsync<JsonElement>(context.Request.Body);
        var toolName = request.GetProperty("name").GetString();
        var arguments = request.GetProperty("arguments");
        
        var config = context.RequestServices.GetRequiredService<IConfiguration>();
        var graphClient = await GetGraphClient(config);
        
        object result = toolName switch
        {
            "search_list_items" => await SearchListItems(graphClient, arguments),
            "get_list_items" => await GetListItems(graphClient, arguments),
            _ => new { error = "Unknown tool" }
        };
        
        var response = new
        {
            content = new[]
            {
                new
                {
                    type = "text",
                    text = JsonSerializer.Serialize(result, new JsonSerializerOptions 
                    { 
                        WriteIndented = true 
                    })
                }
            }
        };
        
        await context.Response.WriteAsJsonAsync(response);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsJsonAsync(new
        {
            content = new[]
            {
                new
                {
                    type = "text",
                    text = $"Error: {ex.Message}"
                }
            }
        });
    }
});

app.Run();

static Task<GraphServiceClient> GetGraphClient(IConfiguration config)
{
    var tenantId = config["SharePoint:TenantId"];
    var clientId = config["SharePoint:ClientId"];
    var clientSecret = config["SharePoint:ClientSecret"];

    var credential = new ClientSecretCredential(tenantId!, clientId!, clientSecret!);
    return Task.FromResult(new GraphServiceClient(credential));
}

static string GetRequiredString(JsonElement args, string name)
{
    if (args.TryGetProperty(name, out var prop) && prop.ValueKind == JsonValueKind.String)
    {
        return prop.GetString()!;
    }

    throw new ArgumentException($"{name} is required");
}

static async Task<object> SearchListItems(GraphServiceClient client, JsonElement args)
{
    var siteUrl = GetRequiredString(args, "siteUrl");
    var listTitle = GetRequiredString(args, "listTitle");
    var searchQuery = GetRequiredString(args, "searchQuery");
    var maxResults = args.TryGetProperty("maxResults", out var max) && max.ValueKind == JsonValueKind.Number ? max.GetInt32() : 50;
    
    var siteId = await GetSiteId(client, siteUrl);
    var list = await client.Sites[siteId].Lists[listTitle].GetAsync();

    if (list?.Id == null)
    {
        return new
        {
            success = false,
            error = "List not found"
        };
    }

    var items = await client.Sites[siteId].Lists[list.Id]
        .Items
        .GetAsync(config =>
        {
            config.QueryParameters.Expand = new[] { "fields" };
            config.QueryParameters.Top = maxResults;
        });
    
    if (items?.Value == null)
    {
        return new
        {
            success = true,
            count = 0,
            items = new object[0]
        };
    }

    var results = new List<object>();
    
    foreach (var item in items.Value)
    {
        if (item.Fields?.AdditionalData != null)
        {
            var matchesSearch = item.Fields.AdditionalData.Values
                .Any(v => v?.ToString()?.Contains(searchQuery, StringComparison.OrdinalIgnoreCase) ?? false);
            
            if (matchesSearch)
            {
                results.Add(new
                {
                    id = item.Id,
                    fields = item.Fields.AdditionalData
                });
            }
        }
    }
    
    return new
    {
        success = true,
        count = results.Count,
        items = results
    };
}

static async Task<object> GetListItems(GraphServiceClient client, JsonElement args)
{
    var siteUrl = GetRequiredString(args, "siteUrl");
    var listTitle = GetRequiredString(args, "listTitle");
    var filter = args.TryGetProperty("filter", out var f) && f.ValueKind == JsonValueKind.String ? f.GetString() : null;
    var maxResults = args.TryGetProperty("maxResults", out var max) && max.ValueKind == JsonValueKind.Number ? max.GetInt32() : 100;
    
    var siteId = await GetSiteId(client, siteUrl);
    var list = await client.Sites[siteId].Lists[listTitle].GetAsync();

    if (list?.Id == null)
    {
        return new
        {
            success = false,
            error = "List not found"
        };
    }

    var items = await client.Sites[siteId].Lists[list.Id]
        .Items
        .GetAsync(config =>
        {
            config.QueryParameters.Expand = new[] { "fields" };
            config.QueryParameters.Top = maxResults;
            if (!string.IsNullOrEmpty(filter))
            {
                config.QueryParameters.Filter = filter;
            }
        });
    
    if (items?.Value == null)
    {
        return new
        {
            success = true,
            count = 0,
            items = new object[0]
        };
    }

    var results = items.Value.Select(item => new
    {
        id = item.Id,
        fields = item.Fields?.AdditionalData
    }).ToList();
    
    return new
    {
        success = true,
        count = results.Count,
        items = results
    };
}

static async Task<string> GetSiteId(GraphServiceClient client, string siteUrl)
{
    var uri = new Uri(siteUrl);
    var hostname = uri.Host;
    var sitePath = uri.AbsolutePath;
    
    var site = await client.Sites[$"{hostname}:{sitePath}"].GetAsync();
    if (site?.Id == null)
    {
        throw new InvalidOperationException("Could not resolve site id");
    }

    return site.Id;
}