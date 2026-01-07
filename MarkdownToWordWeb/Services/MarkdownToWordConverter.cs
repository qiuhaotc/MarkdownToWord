using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Tables;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using MarkdigTable = Markdig.Extensions.Tables.Table;
using MarkdigTableRow = Markdig.Extensions.Tables.TableRow;
using MarkdigTableCell = Markdig.Extensions.Tables.TableCell;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WordTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WordTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace MarkdownToWordWeb.Services;

public class MarkdownToWordConverter
{
    private static readonly HttpClient _httpClient = new HttpClient
    {
        Timeout = TimeSpan.FromSeconds(10)
    };

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

            // Add document settings for modern Word format
            AddDocumentSettings();

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

    private void AddDocumentSettings()
    {
        if (_wordDocument == null || _mainPart == null) return;

        // Add document settings part for modern Word format
        var settingsPart = _mainPart.AddNewPart<DocumentSettingsPart>();
        var settings = new Settings(
            new Compatibility(
                new CompatibilitySetting
                {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "15" // Word 2013 and later (15 = Word 2013, 16 = Word 2016+)
                }
            )
        );
        settingsPart.Settings = settings;
        settingsPart.Settings.Save();
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
            case Markdig.Syntax.LinkReferenceDefinitionGroup:
                // Skip link reference definitions - they are not rendered
                break;
            default:
                // For any unhandled block types, skip silently
                // Do not render unknown block types to avoid outputting class names
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
            ProcessInlineIntoParagraph(inline, paragraph);
        }
    }

    private void ProcessInlineIntoParagraph(Inline inline, Paragraph paragraph)
    {
        switch (inline)
        {
            case LinkInline link:
                if (link.IsImage)
                {
                    // Handle images - embed actual images from URLs
                    if (!string.IsNullOrEmpty(link.Url))
                    {
                        var altText = link.FirstChild is LiteralInline lit ? lit.Content.ToString() : "";
                        var run = new Run();
                        AddImage(link.Url, altText, run);
                        if (run.HasChildren)
                        {
                            paragraph.AppendChild(run);
                        }
                    }
                }
                else
                {
                    // Handle regular links - create proper Word hyperlinks
                    if (!string.IsNullOrEmpty(link.Url))
                    {
                        CreateHyperlinkInParagraph(link, paragraph);
                    }
                    else
                    {
                        // If no URL, just render the text
                        var run = new Run();
                        AppendInlines(link, run);
                        if (run.HasChildren)
                        {
                            paragraph.AppendChild(run);
                        }
                    }
                }
                break;
            case EmphasisInline emphasis:
                // Handle emphasis that might contain links
                ProcessEmphasis(emphasis, null, paragraph);
                break;
            default:
                var defaultRun = new Run();
                ProcessInline(inline, defaultRun);
                if (defaultRun.HasChildren)
                {
                    paragraph.AppendChild(defaultRun);
                }
                break;
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
                // Process emphasis with possible nested links
                ProcessEmphasis(emphasis, run, null);
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
            case LineBreakInline:
                run.AppendChild(new Break());
                break;
            case LinkInline link:
                // Handle links when not directly in paragraph (e.g., in emphasis)
                if (!link.IsImage && !string.IsNullOrEmpty(link.Url))
                {
                    // Need to create hyperlink at paragraph level, so we can't handle it here
                    // Just add the text for now
                    run.AppendChild(new Text(GetLinkText(link)) { Space = SpaceProcessingModeValues.Preserve });
                }
                else if (!link.IsImage)
                {
                    AppendInlines(link, run);
                }
                break;
            case ContainerInline container:
                AppendInlines(container, run);
                break;
        }
    }

    private void ProcessEmphasis(EmphasisInline emphasis, Run? existingRun, Paragraph? paragraph)
    {
        // Check if emphasis contains links
        bool containsLink = ContainsLink(emphasis);

        if (containsLink && paragraph != null)
        {
            // Process emphasis with links at paragraph level
            foreach (var child in emphasis)
            {
                if (child is LinkInline link && !link.IsImage && !string.IsNullOrEmpty(link.Url))
                {
                    CreateHyperlinkWithEmphasis(link, paragraph, emphasis.DelimiterCount);
                }
                else if (child is LiteralInline literal)
                {
                    var run = new Run();
                    var runProps = new RunProperties();

                    if (emphasis.DelimiterCount == 2) // Bold
                    {
                        runProps.AppendChild(new Bold());
                    }
                    else if (emphasis.DelimiterCount == 1) // Italic
                    {
                        runProps.AppendChild(new Italic());
                    }

                    run.AppendChild(runProps);
                    run.AppendChild(new Text(literal.Content.ToString()) { Space = SpaceProcessingModeValues.Preserve });
                    paragraph.AppendChild(run);
                }
                else if (child is ContainerInline container)
                {
                    ProcessEmphasizedContainer(container, paragraph, emphasis.DelimiterCount);
                }
            }
        }
        else if (paragraph != null)
        {
            // Process emphasis without links at paragraph level
            var run = new Run();
            var runProps = new RunProperties();

            if (emphasis.DelimiterCount == 2) // Bold
            {
                runProps.AppendChild(new Bold());
            }
            else if (emphasis.DelimiterCount == 1) // Italic
            {
                runProps.AppendChild(new Italic());
            }
            else if (emphasis.DelimiterCount == 3) // Bold + Italic
            {
                runProps.AppendChild(new Bold());
                runProps.AppendChild(new Italic());
            }

            run.AppendChild(runProps);
            AppendInlines(emphasis, run);

            if (run.HasChildren)
            {
                paragraph.AppendChild(run);
            }
        }
        else if (existingRun != null)
        {
            // Process emphasis in existing run (nested in other formatting)
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
            else if (emphasis.DelimiterCount == 3) // Bold + Italic
            {
                runProps.AppendChild(new Bold());
                runProps.AppendChild(new Italic());
            }

            emphasisRun.AppendChild(runProps);
            AppendInlines(emphasis, emphasisRun);

            // Copy children from emphasisRun to the main run
            foreach (var child in emphasisRun.ChildElements.ToList())
            {
                existingRun.AppendChild(child.CloneNode(true));
            }
        }
    }

    private bool ContainsLink(ContainerInline container)
    {
        foreach (var inline in container)
        {
            if (inline is LinkInline link && !link.IsImage)
            {
                return true;
            }
            if (inline is ContainerInline nested && ContainsLink(nested))
            {
                return true;
            }
        }
        return false;
    }

    private void ProcessEmphasizedContainer(ContainerInline container, Paragraph paragraph, int delimiterCount)
    {
        foreach (var inline in container)
        {
            if (inline is LinkInline link && !link.IsImage && !string.IsNullOrEmpty(link.Url))
            {
                CreateHyperlinkWithEmphasis(link, paragraph, delimiterCount);
            }
            else if (inline is LiteralInline literal)
            {
                var run = new Run();
                var runProps = new RunProperties();

                if (delimiterCount == 2) // Bold
                {
                    runProps.AppendChild(new Bold());
                }
                else if (delimiterCount == 1) // Italic
                {
                    runProps.AppendChild(new Italic());
                }

                run.AppendChild(runProps);
                run.AppendChild(new Text(literal.Content.ToString()) { Space = SpaceProcessingModeValues.Preserve });
                paragraph.AppendChild(run);
            }
            else if (inline is ContainerInline nested)
            {
                ProcessEmphasizedContainer(nested, paragraph, delimiterCount);
            }
        }
    }

    private void CreateHyperlinkWithEmphasis(LinkInline link, Paragraph paragraph, int delimiterCount)
    {
        if (_mainPart == null) return;

        try
        {
            Uri uri;
            try
            {
                var normalizedUrl = NormalizeUrl(link.Url ?? "");
                uri = new Uri(normalizedUrl, UriKind.RelativeOrAbsolute);
            }
            catch (Exception)
            {
                uri = new Uri("http://invalid-url/", UriKind.Absolute);
            }

            var hyperlinkRel = _mainPart.AddHyperlinkRelationship(uri, true);

            var runProps = new RunProperties(
                new Underline { Val = UnderlineValues.Single },
                new Color { Val = "0563C1" }
            );

            if (delimiterCount == 2) // Bold
            {
                runProps.AppendChild(new Bold());
            }
            else if (delimiterCount == 1) // Italic
            {
                runProps.AppendChild(new Italic());
            }

            var hyperlink = new Hyperlink(new Run(
                runProps,
                new Text(GetLinkText(link)) { Space = SpaceProcessingModeValues.Preserve }
            ))
            {
                Id = hyperlinkRel.Id
            };

            paragraph.AppendChild(hyperlink);
        }
        catch
        {
            var run = new Run();
            var runProps = new RunProperties();

            if (delimiterCount == 2)
            {
                runProps.AppendChild(new Bold());
            }
            else if (delimiterCount == 1)
            {
                runProps.AppendChild(new Italic());
            }

            run.AppendChild(runProps);
            run.AppendChild(new Text(GetLinkText(link)) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
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
        // Get all lines as a list
        var lines = codeBlock.Lines.Lines.ToList();

        if (lines.Count == 0)
        {
            // If code block is empty, add one empty paragraph
            var emptyParagraph = new Paragraph();
            var emptyParaProps = new ParagraphProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = "F5F5F5" },
                new SpacingBetweenLines { Before = "100", After = "100" }
            );
            emptyParagraph.AppendChild(emptyParaProps);
            body.AppendChild(emptyParagraph);
            return;
        }

        // Process each line of code as a separate paragraph to preserve formatting
        for (int i = 0; i < lines.Count; i++)
        {
            var paragraph = new Paragraph();
            var paraProps = new ParagraphProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = "F5F5F5" },
                new SpacingBetweenLines { Before = i == 0 ? "100" : "0", After = i == lines.Count - 1 ? "100" : "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
            );
            paragraph.AppendChild(paraProps);

            var run = new Run();
            var runProps = new RunProperties(
                new RunFonts { Ascii = "Courier New" },
                new FontSize { Val = "20" }
            );
            run.AppendChild(runProps);

            var lineText = lines[i].ToString();
            run.AppendChild(new Text(lineText) { Space = SpaceProcessingModeValues.Preserve });

            paragraph.AppendChild(run);
            body.AppendChild(paragraph);
        }
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

    private string NormalizeUrl(string url)
    {
        if (string.IsNullOrEmpty(url))
            return url;

        // Handle paths starting with backslash (e.g., \File\README.md)
        // Convert backslashes to forward slashes for URI compatibility
        if (url.StartsWith("\\") || url.Contains("\\"))
        {
            // Replace backslashes with forward slashes
            url = url.Replace("\\", "/");
        }

        return url;
    }

    private void CreateHyperlinkInParagraph(LinkInline link, Paragraph paragraph)
    {
        // Create a proper Word hyperlink with relationship
        if (_mainPart == null) return;

        try
        {
            Uri uri;
            try
            {
                var normalizedUrl = NormalizeUrl(link.Url ?? "");
                uri = new Uri(normalizedUrl, UriKind.RelativeOrAbsolute);
            }
            catch (Exception)
            {
                // TODO: Add logs for invalid URL formats
                uri = new Uri("http://invalid-url/", UriKind.Absolute);
            }

            // Add hyperlink relationship to the document
            var hyperlinkRel = _mainPart.AddHyperlinkRelationship(uri, true);

            // Create hyperlink element with explicit styling
            var hyperlink = new Hyperlink(new Run(
                new RunProperties(
                    new Underline { Val = UnderlineValues.Single },
                    new Color { Val = "0563C1" } // Standard blue hyperlink color
                ),
                new Text(GetLinkText(link)) { Space = SpaceProcessingModeValues.Preserve }
            ))
            {
                Id = hyperlinkRel.Id
            };

            paragraph.AppendChild(hyperlink);
        }
        catch
        {
            // If hyperlink creation fails, add as regular text
            var run = new Run(new Text(GetLinkText(link)) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
        }
    }

    private string GetLinkText(LinkInline link)
    {
        // Extract text from link
        foreach (var inline in link)
        {
            if (inline is LiteralInline literal)
            {
                return literal.Content.ToString();
            }
        }

        // If no text, use URL
        return link.Url ?? "";
    }

    private void AddImage(string imageUrl, string altText, Run run)
    {
        try
        {
            // Try to download and embed the image using static HttpClient
            var imageBytes = Task.Run(async () => await _httpClient.GetByteArrayAsync(imageUrl)).GetAwaiter().GetResult();

            if (_mainPart == null)
            {
                AddImageFallback(altText, run);
                return;
            }

            if (imageBytes == null || imageBytes.Length == 0)
            {
                AddImageFallback(altText, run);
                return;
            }

            // Determine image type from URL and add image part
            var imagePart = AddImagePartByType(_mainPart, imageUrl, imageBytes);

            using (var stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }

            // Add the image to the document
            var relationshipId = _mainPart.GetIdOfPart(imagePart);

            // Get image dimensions (simplified - using fixed size)
            long widthEmus = 3000000; // ~3.17 inches
            long heightEmus = 2000000; // ~2.11 inches

            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = widthEmus, Cy = heightEmus },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = 1U, Name = altText ?? "Image" },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = altText ?? "Image" },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relationshipId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = widthEmus, Cy = heightEmus }),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                });

            run.AppendChild(element);
        }
        catch (HttpRequestException ex)
        {
            // Network or HTTP error - show alt text with error info
            Console.WriteLine($"Image download failed for {imageUrl}: {ex.Message}");
            AddImageFallback(altText, run);
        }
        catch (TaskCanceledException ex)
        {
            // Timeout - show alt text
            Console.WriteLine($"Image download timeout for {imageUrl}: {ex.Message}");
            AddImageFallback(altText, run);
        }
        catch (Exception ex)
        {
            // Any other error (invalid format, etc.) - show alt text
            Console.WriteLine($"Image embedding failed for {imageUrl}: {ex.Message}");
            AddImageFallback(altText, run);
        }
    }

    private void AddImageFallback(string altText, Run run)
    {
        var imgRun = new Run();
        var imgProps = new RunProperties(
            new Italic(),
            new Color { Val = "808080" }
        );
        imgRun.AppendChild(imgProps);
        imgRun.AppendChild(new Text($"[Image: {altText}]") { Space = SpaceProcessingModeValues.Preserve });

        foreach (var child in imgRun.ChildElements.ToList())
        {
            run.AppendChild(child.CloneNode(true));
        }
    }

    private ImagePart AddImagePartByType(MainDocumentPart mainPart, string url, byte[] imageBytes)
    {
        // Try to determine from URL extension
        var extension = Path.GetExtension(url).ToLowerInvariant();

        return extension switch
        {
            ".png" => mainPart.AddImagePart(ImagePartType.Png),
            ".jpg" or ".jpeg" => mainPart.AddImagePart(ImagePartType.Jpeg),
            ".gif" => mainPart.AddImagePart(ImagePartType.Gif),
            ".bmp" => mainPart.AddImagePart(ImagePartType.Bmp),
            ".tiff" or ".tif" => mainPart.AddImagePart(ImagePartType.Tiff),
            _ => mainPart.AddImagePart(ImagePartType.Jpeg) // Default to JPEG
        };
    }
}
