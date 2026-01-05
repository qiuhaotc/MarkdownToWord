using Markdig;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToWordWeb.Services;

public class MarkdownToWordConverter
{
    public byte[] ConvertMarkdownToWord(string markdownContent)
    {
        // Parse markdown to HTML
        var html = Markdown.ToHtml(markdownContent);
        
        // Create Word document in memory
        using var memoryStream = new MemoryStream();
        using (var wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
        {
            // Add main document part
            var mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            
            // Parse HTML and convert to Word paragraphs
            var lines = markdownContent.Split('\n');
            
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    body.AppendChild(new Paragraph());
                    continue;
                }
                
                var paragraph = new Paragraph();
                var run = new Run();
                var text = new Text(line);
                
                // Handle headers
                if (line.StartsWith("# "))
                {
                    text = new Text(line.Substring(2));
                    run.AppendChild(new RunProperties(
                        new Bold(),
                        new FontSize { Val = "32" }
                    ));
                }
                else if (line.StartsWith("## "))
                {
                    text = new Text(line.Substring(3));
                    run.AppendChild(new RunProperties(
                        new Bold(),
                        new FontSize { Val = "28" }
                    ));
                }
                else if (line.StartsWith("### "))
                {
                    text = new Text(line.Substring(4));
                    run.AppendChild(new RunProperties(
                        new Bold(),
                        new FontSize { Val = "24" }
                    ));
                }
                else if (line.StartsWith("#### "))
                {
                    text = new Text(line.Substring(5));
                    run.AppendChild(new RunProperties(
                        new Bold(),
                        new FontSize { Val = "22" }
                    ));
                }
                else if (line.StartsWith("##### "))
                {
                    text = new Text(line.Substring(6));
                    run.AppendChild(new RunProperties(
                        new Bold(),
                        new FontSize { Val = "20" }
                    ));
                }
                else if (line.StartsWith("###### "))
                {
                    text = new Text(line.Substring(7));
                    run.AppendChild(new RunProperties(
                        new Bold(),
                        new FontSize { Val = "18" }
                    ));
                }
                // Handle bold text **text**
                else if (line.Contains("**"))
                {
                    text = new Text(line.Replace("**", ""));
                    run.AppendChild(new RunProperties(new Bold()));
                }
                // Handle italic text *text*
                else if (line.Contains("*") && !line.StartsWith("* "))
                {
                    text = new Text(line.Replace("*", ""));
                    run.AppendChild(new RunProperties(new Italic()));
                }
                // Handle bullet points
                else if (line.TrimStart().StartsWith("* ") || line.TrimStart().StartsWith("- "))
                {
                    var bulletText = line.TrimStart().Substring(2);
                    text = new Text("• " + bulletText);
                    paragraph.AppendChild(new ParagraphProperties(
                        new Indentation { Left = "720" }
                    ));
                }
                // Handle numbered lists
                else if (char.IsDigit(line.TrimStart()[0]) && line.TrimStart().Contains(". "))
                {
                    text = new Text(line.TrimStart());
                    paragraph.AppendChild(new ParagraphProperties(
                        new Indentation { Left = "720" }
                    ));
                }
                
                run.AppendChild(text);
                paragraph.AppendChild(run);
                body.AppendChild(paragraph);
            }
            
            mainPart.Document.AppendChild(body);
            mainPart.Document.Save();
        }
        
        return memoryStream.ToArray();
    }
}
