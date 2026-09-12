using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Text;
namespace WordFunctions;

public class PopulateWord(ILogger<PopulateWord> logger)
{
    private readonly ILogger<PopulateWord> _logger = logger;

    [Function("PopulateWord")]
    public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req)
    {
        _logger.LogInformation("Processing word document generation request.");
        
        string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
        var data = JsonSerializer.Deserialize<DocumentRequest>(requestBody, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        if (data == null || string.IsNullOrEmpty(data.TemplateBase64))
        {
            return new BadRequestObjectResult("Please provide a valid TemplateBase64 and Tokens.");
        }


        byte[] templateBytes = Convert.FromBase64String(data.TemplateBase64);
        using var memoryStream = new MemoryStream();
        memoryStream.Write(templateBytes, 0, templateBytes.Length);

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, true))
        {
            var body = wordDoc!.MainDocumentPart!.Document!.Body!;
            foreach (var token in data.Tokens)
            {
                var tokenText = $"{token.Key}";
                var valueText = token.Value?.ToString() ?? string.Empty;
                Console.WriteLine($"Replacing token: {token.Key} with value: {token.Value}");
                ReplaceContentControlText(body, tokenText, valueText);
            }
            wordDoc.MainDocumentPart.Document.Save();
        }

        byte[] fileBytes = memoryStream.ToArray();
        string base64Output = Convert.ToBase64String(fileBytes);

        var responsePayload = new
        {
            FileName = $"GeneratedDocument_{Guid.NewGuid().ToString("N").Substring(0, 8)}.docx",
            FileContentBase64 = base64Output
        };

        return new OkObjectResult(responsePayload);
    }

    private static void ReplaceContentControlText(Body body, string tag, string value)
    {
        var contentControls = body.Descendants<SdtElement>()
            .Where(sdt => sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == tag)
            .ToList();
        Console.WriteLine($"Found {contentControls.Count} content controls for tag '{tag}'.");
        foreach (var sdt in contentControls)
        {
            var textNodes = sdt.Descendants<Text>().ToList();
            Console.WriteLine($"Found text nodes for tag '{tag}': {textNodes.Count}");
            if (textNodes.Count != 0)
            {
                textNodes.First().Text = value;

                foreach (var textNode in textNodes.Skip(1))
                {
                    textNode.Text = string.Empty;
                }
            }
            else
            {
                sdt.AppendChild(new SdtContentRun(new Run(new Text(value))));
            }
        }
    }
}
