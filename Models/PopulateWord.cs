using System.Collections.Generic;

namespace WordFunctions
{
    public class DocumentRequest
    {
        // The .docx template converted to a Base64 string
        public required string TemplateBase64 { get; set; } 
        public required Dictionary<string, object> Tokens { get; set; }
    }

    public class TokenReplacement
    {
        public required string Token { get; set; }
        public required string Value { get; set; }
    }

    public class DocumentResponse
    {
        public required string FileName { get; set; }
        public required string FileContentBase64 { get; set; } 
    }
}