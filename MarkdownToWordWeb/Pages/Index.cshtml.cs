using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using MarkdownToWordWeb.Services;

namespace MarkdownToWordWeb.Pages;

public class IndexModel : PageModel
{
    private readonly ILogger<IndexModel> _logger;
    private readonly MarkdownToWordConverter _converter;

    public string? ErrorMessage { get; set; }
    public string? SuccessMessage { get; set; }

    public IndexModel(ILogger<IndexModel> logger, MarkdownToWordConverter converter)
    {
        _logger = logger;
        _converter = converter;
    }

    public void OnGet()
    {
    }

    public async Task<IActionResult> OnPostAsync(IFormFile markdownFile)
    {
        try
        {
            if (markdownFile == null || markdownFile.Length == 0)
            {
                ErrorMessage = "Please select a file to upload.";
                return Page();
            }

            if (!IsMarkdownFile(markdownFile.FileName))
            {
                ErrorMessage = "Please upload a valid Markdown file (.md, .markdown, or .txt).";
                return Page();
            }

            // Read the markdown content
            string markdownContent;
            using (var reader = new StreamReader(markdownFile.OpenReadStream()))
            {
                markdownContent = await reader.ReadToEndAsync();
            }

            // Convert to Word
            var wordBytes = _converter.ConvertMarkdownToWord(markdownContent);

            // Return the Word document
            var fileName = Path.GetFileNameWithoutExtension(markdownFile.FileName) + ".docx";
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error converting markdown to Word");
            ErrorMessage = $"An error occurred during conversion: {ex.Message}";
            return Page();
        }
    }

    public IActionResult OnPostText(string markdownText)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(markdownText))
            {
                ErrorMessage = "Please enter some markdown content.";
                return Page();
            }

            // Convert to Word
            var wordBytes = _converter.ConvertMarkdownToWord(markdownText);

            // Return the Word document
            var fileName = $"converted_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error converting markdown text to Word");
            ErrorMessage = $"An error occurred during conversion: {ex.Message}";
            return Page();
        }
    }

    private bool IsMarkdownFile(string fileName)
    {
        var extension = Path.GetExtension(fileName).ToLowerInvariant();
        return extension == ".md" || extension == ".markdown" || extension == ".txt";
    }
}
