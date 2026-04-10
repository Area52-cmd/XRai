using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace XRai.Demo.MacroGuard;

public class AiAssistant
{
    private static readonly HttpClient HttpClient = new();

    public record AiResponse(string Response, int TokensUsed);

    public async Task<AiResponse> AskAsync(
        string apiKey, string prompt, string model, double temperature, int maxTokens, int timeoutSeconds)
    {
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));

        var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");

        var body = new
        {
            model,
            max_tokens = maxTokens,
            temperature,
            messages = new[]
            {
                new
                {
                    role = "user",
                    content = $"You are a VBA expert assistant for Excel. {prompt}"
                }
            }
        };

        request.Content = JsonContent.Create(body);
        var response = await HttpClient.SendAsync(request, cts.Token);
        var json = await response.Content.ReadAsStringAsync(cts.Token);

        if (!response.IsSuccessStatusCode)
        {
            throw new Exception($"API error ({response.StatusCode}): {json}");
        }

        var result = JsonSerializer.Deserialize<ApiResponse>(json);
        var text = result?.Content?.FirstOrDefault()?.Text ?? "No response";
        var tokens = (result?.Usage?.InputTokens ?? 0) + (result?.Usage?.OutputTokens ?? 0);

        return new AiResponse(text, tokens);
    }

    // JSON deserialization models
    private class ApiResponse
    {
        [JsonPropertyName("content")]
        public List<ContentBlock>? Content { get; set; }

        [JsonPropertyName("usage")]
        public UsageInfo? Usage { get; set; }
    }

    private class ContentBlock
    {
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("text")]
        public string? Text { get; set; }
    }

    private class UsageInfo
    {
        [JsonPropertyName("input_tokens")]
        public int InputTokens { get; set; }

        [JsonPropertyName("output_tokens")]
        public int OutputTokens { get; set; }
    }
}
