using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Tables;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MarkdigTable = Markdig.Extensions.Tables.Table;
using MarkdigTableRow = Markdig.Extensions.Tables.TableRow;
using MarkdigTableCell = Markdig.Extensions.Tables.TableCell;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WordTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WordTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace MarkdownToWordWeb.Services;

public class MarkdownToWordConverter
{
    private WordprocessingDocument? _wordDocument;
    private MainDocumentPart? _mainPart;

    public byte[] ConvertMarkdownToWord(string markdownContent)
    {
        // Parse markdown using Markdig with advanced extensions
        var pipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .Build();
        
        var document = Markdown.Parse(markdownContent, pipeline);
        
        // Create Word document in memory
        using var memoryStream = new MemoryStream();
        using (_wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
        {
            // Add main document part
            _mainPart = _wordDocument.AddMainDocumentPart();
            _mainPart.Document = new Document();
            var body = new Body();
            
            // Convert markdown AST to Word document
            foreach (var block in document)
            {
                ConvertBlock(block, body);
            }
            
            _mainPart.Document.AppendChild(body);
            _mainPart.Document.Save();
        }
        
        return memoryStream.ToArray();
    }

    private void ConvertBlock(Block block, Body body)
    {
        switch (block)
        {
            case HeadingBlock heading:
                ConvertHeading(heading, body);
                break;
            case ParagraphBlock paragraph:
                ConvertParagraph(paragraph, body);
                break;
            case ListBlock list:
                ConvertList(list, body);
                break;
            case MarkdigTable table:
                ConvertTable(table, body);
                break;
            case CodeBlock codeBlock:
                ConvertCodeBlock(codeBlock, body);
                break;
            case QuoteBlock quote:
                ConvertQuoteBlock(quote, body);
                break;
            case ThematicBreakBlock:
                ConvertThematicBreak(body);
                break;
            default:
                // For any unhandled block types, try to render as paragraph
                var para = new Paragraph();
                para.AppendChild(new Run(new Text(block.ToString() ?? "")));
                body.AppendChild(para);
                break;
        }
    }

    private void ConvertHeading(HeadingBlock heading, Body body)
    {
        var paragraph = new Paragraph();
        var run = new Run();
        
        // Set heading style based on level
        var fontSize = heading.Level switch
        {
            1 => "32",
            2 => "28",
            3 => "24",
            4 => "22",
            5 => "20",
            _ => "18"
        };
        
        var runProps = new RunProperties(
            new Bold(),
            new FontSize { Val = fontSize }
        );
        run.AppendChild(runProps);
        
        // Extract text from inline elements
        if (heading.Inline != null)
        {
            AppendInlines(heading.Inline, run);
        }
        
        paragraph.AppendChild(run);
        body.AppendChild(paragraph);
    }

    private void ConvertParagraph(ParagraphBlock paragraphBlock, Body body)
    {
        var paragraph = new Paragraph();
        
        if (paragraphBlock.Inline != null)
        {
            ConvertInlines(paragraphBlock.Inline, paragraph);
        }
        
        body.AppendChild(paragraph);
    }

    private void ConvertInlines(ContainerInline container, Paragraph paragraph)
    {
        foreach (var inline in container)
        {
            var run = new Run();
            ProcessInline(inline, run);
            if (run.HasChildren)
            {
                paragraph.AppendChild(run);
            }
        }
    }

    private void ProcessInline(Inline inline, Run run)
    {
        switch (inline)
        {
            case LiteralInline literal:
                run.AppendChild(new Text(literal.Content.ToString()) { Space = SpaceProcessingModeValues.Preserve });
                break;
            case EmphasisInline emphasis:
                var emphasisRun = new Run();
                var runProps = new RunProperties();
                
                if (emphasis.DelimiterCount == 2) // Bold
                {
                    runProps.AppendChild(new Bold());
                }
                else if (emphasis.DelimiterCount == 1) // Italic
                {
                    runProps.AppendChild(new Italic());
                }
                
                emphasisRun.AppendChild(runProps);
                AppendInlines(emphasis, emphasisRun);
                
                // Copy children from emphasisRun to the main run
                foreach (var child in emphasisRun.ChildElements.ToList())
                {
                    run.AppendChild(child.CloneNode(true));
                }
                break;
            case CodeInline code:
                var codeRun = new Run();
                var codeProps = new RunProperties(
                    new RunFonts { Ascii = "Courier New" },
                    new Color { Val = "C7254E" },
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "F9F2F4" }
                );
                codeRun.AppendChild(codeProps);
                codeRun.AppendChild(new Text(code.Content) { Space = SpaceProcessingModeValues.Preserve });
                
                foreach (var child in codeRun.ChildElements.ToList())
                {
                    run.AppendChild(child.CloneNode(true));
                }
                break;
            case LinkInline link:
                var linkRun = new Run();
                var linkProps = new RunProperties(
                    new Underline { Val = UnderlineValues.Single },
                    new Color { Val = "0563C1" }
                );
                linkRun.AppendChild(linkProps);
                
                if (link.IsImage)
                {
                    // Handle images
                    if (!string.IsNullOrEmpty(link.Url))
                    {
                        try
                        {
                            AddImage(link.Url, link.Title ?? "", run);
                        }
                        catch
                        {
                            // If image can't be added, show alt text
                            var altText = link.FirstChild is LiteralInline lit ? lit.Content.ToString() : "Image";
                            linkRun.AppendChild(new Text($"[Image: {altText}]"));
                            foreach (var child in linkRun.ChildElements.ToList())
                            {
                                run.AppendChild(child.CloneNode(true));
                            }
                        }
                    }
                }
                else
                {
                    // Handle regular links - show link text with URL in parentheses
                    AppendInlines(link, linkRun);
                    if (!string.IsNullOrEmpty(link.Url))
                    {
                        linkRun.AppendChild(new Text($" ({link.Url})") { Space = SpaceProcessingModeValues.Preserve });
                    }
                    
                    foreach (var child in linkRun.ChildElements.ToList())
                    {
                        run.AppendChild(child.CloneNode(true));
                    }
                }
                break;
            case LineBreakInline:
                run.AppendChild(new Break());
                break;
            case ContainerInline container:
                AppendInlines(container, run);
                break;
        }
    }

    private void AppendInlines(ContainerInline container, Run run)
    {
        foreach (var inline in container)
        {
            ProcessInline(inline, run);
        }
    }

    private void ConvertList(ListBlock list, Body body)
    {
        int itemNumber = 1;
        
        foreach (var item in list.Cast<ListItemBlock>())
        {
            bool isFirstBlock = true;
            
            foreach (var block in item)
            {
                if (block is ParagraphBlock paragraph)
                {
                    var para = new Paragraph();
                    var paraProps = new ParagraphProperties(
                        new Indentation { Left = "720", Hanging = isFirstBlock ? "360" : "0" }
                    );
                    para.AppendChild(paraProps);
                    
                    var run = new Run();
                    
                    // Add bullet or number only for the first block in the list item
                    if (isFirstBlock)
                    {
                        if (list.IsOrdered)
                        {
                            run.AppendChild(new Text($"{itemNumber}. ") { Space = SpaceProcessingModeValues.Preserve });
                        }
                        else
                        {
                            run.AppendChild(new Text("• ") { Space = SpaceProcessingModeValues.Preserve });
                        }
                    }
                    
                    para.AppendChild(run);
                    
                    // Add content
                    if (paragraph.Inline != null)
                    {
                        ConvertInlines(paragraph.Inline, para);
                    }
                    
                    body.AppendChild(para);
                    isFirstBlock = false;
                }
                else
                {
                    ConvertBlock(block, body);
                    isFirstBlock = false;
                }
            }
            
            // Increment item number after processing all blocks in the list item
            if (list.IsOrdered)
            {
                itemNumber++;
            }
        }
    }

    private void ConvertTable(MarkdigTable table, Body body)
    {
        var wordTable = new WordTable();
        
        // Table properties
        var tableProps = new TableProperties(
            new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
            ),
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
        );
        wordTable.AppendChild(tableProps);
        
        // Process table rows
        foreach (var row in table.Cast<MarkdigTableRow>())
        {
            var wordRow = new WordTableRow();
            
            foreach (var cell in row.Cast<MarkdigTableCell>())
            {
                var wordCell = new WordTableCell();
                
                // Cell properties
                var cellProps = new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Auto }
                );
                
                // Header row styling
                if (row.IsHeader)
                {
                    cellProps.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = "D9D9D9" });
                }
                
                wordCell.AppendChild(cellProps);
                
                // Process cell content
                foreach (var block in cell)
                {
                    if (block is ParagraphBlock cellPara)
                    {
                        var para = new Paragraph();
                        if (row.IsHeader)
                        {
                            var run = new Run(new RunProperties(new Bold()));
                            if (cellPara.Inline != null)
                            {
                                AppendInlines(cellPara.Inline, run);
                            }
                            para.AppendChild(run);
                        }
                        else
                        {
                            if (cellPara.Inline != null)
                            {
                                ConvertInlines(cellPara.Inline, para);
                            }
                        }
                        wordCell.AppendChild(para);
                    }
                }
                
                // Ensure cell has at least one paragraph
                if (!wordCell.Elements<Paragraph>().Any())
                {
                    wordCell.AppendChild(new Paragraph());
                }
                
                wordRow.AppendChild(wordCell);
            }
            
            wordTable.AppendChild(wordRow);
        }
        
        body.AppendChild(wordTable);
        
        // Add spacing after table
        body.AppendChild(new Paragraph());
    }

    private void ConvertCodeBlock(CodeBlock codeBlock, Body body)
    {
        var paragraph = new Paragraph();
        var paraProps = new ParagraphProperties(
            new Shading { Val = ShadingPatternValues.Clear, Fill = "F5F5F5" },
            new SpacingBetweenLines { Before = "100", After = "100" }
        );
        paragraph.AppendChild(paraProps);
        
        var run = new Run();
        var runProps = new RunProperties(
            new RunFonts { Ascii = "Courier New" },
            new FontSize { Val = "20" }
        );
        run.AppendChild(runProps);
        
        var code = codeBlock.Lines.ToString();
        
        run.AppendChild(new Text(code) { Space = SpaceProcessingModeValues.Preserve });
        paragraph.AppendChild(run);
        body.AppendChild(paragraph);
    }

    private void ConvertQuoteBlock(QuoteBlock quote, Body body)
    {
        foreach (var block in quote)
        {
            if (block is ParagraphBlock paragraph)
            {
                var para = new Paragraph();
                var paraProps = new ParagraphProperties(
                    new Indentation { Left = "720" },
                    new ParagraphBorders(
                        new LeftBorder { Val = BorderValues.Single, Color = "CCCCCC", Size = 12, Space = 4 }
                    )
                );
                para.AppendChild(paraProps);
                
                if (paragraph.Inline != null)
                {
                    ConvertInlines(paragraph.Inline, para);
                }
                
                body.AppendChild(para);
            }
            else
            {
                ConvertBlock(block, body);
            }
        }
    }

    private void ConvertThematicBreak(Body body)
    {
        var paragraph = new Paragraph();
        var paraProps = new ParagraphProperties(
            new ParagraphBorders(
                new BottomBorder { Val = BorderValues.Single, Color = "CCCCCC", Size = 6, Space = 1 }
            )
        );
        paragraph.AppendChild(paraProps);
        body.AppendChild(paragraph);
    }

    private void AddImage(string imageUrl, string altText, Run run)
    {
        // For web URLs, we can't directly embed them in the Word document
        // Instead, show a placeholder with the URL
        var imgRun = new Run();
        var imgProps = new RunProperties(
            new Italic(),
            new Color { Val = "808080" }
        );
        imgRun.AppendChild(imgProps);
        imgRun.AppendChild(new Text($"[Image: {altText}]") { Space = SpaceProcessingModeValues.Preserve });
        imgRun.AppendChild(new Break());
        imgRun.AppendChild(new Text($"URL: {imageUrl}") { Space = SpaceProcessingModeValues.Preserve });
        
        foreach (var child in imgRun.ChildElements.ToList())
        {
            run.AppendChild(child.CloneNode(true));
        }
    }
}
