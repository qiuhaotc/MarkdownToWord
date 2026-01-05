# MarkdownToWord

Easily convert Markdown files to Word documents using C# and ASP.NET web application.

![Web Interface](https://github.com/user-attachments/assets/1b4d50a7-ed57-4131-ba33-662279c31ab4)

## Features

- 📁 **File Upload**: Upload `.md`, `.markdown`, or `.txt` files for conversion
- 📝 **Text Input**: Paste or type markdown content directly in the browser
- 🎨 **Formatting Support**: Converts headers, bold, italic, lists, and paragraphs
- 💻 **Modern Web UI**: Clean, responsive interface built with Bootstrap
- ⚡ **Fast Conversion**: Uses Markdig and DocumentFormat.OpenXml for efficient processing

## Supported Markdown Features

- ✓ Headings (# H1, ## H2, ### H3, etc.)
- ✓ Bold text (\*\*text\*\*)
- ✓ Italic text (\*text\*)
- ✓ Bullet lists (- item or * item)
- ✓ Numbered lists (1. item)
- ✓ Paragraphs with proper spacing
- ✓ **Tables** - Full table support with headers
- ✓ **Links** - [text](url) format
- ✓ **Images** - ![alt](url) format (shows as placeholder with URL)
- ✓ **Code blocks** - Fenced code blocks with ``` 
- ✓ **Inline code** - `code` format with styling
- ✓ **Blockquotes** - > quote format
- ✓ **Horizontal rules** - --- separator

## Technologies Used

- **ASP.NET Core 8.0**: Modern web framework
- **Razor Pages**: Server-side page rendering
- **Markdig**: High-performance Markdown parser
- **DocumentFormat.OpenXml**: Create Word documents (.docx)
- **Bootstrap 5**: Responsive UI framework

## Getting Started

### Prerequisites

- [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) or later

### Running the Application

1. Clone the repository:
   ```bash
   git clone https://github.com/qiuhaotc/MarkdownToWord.git
   cd MarkdownToWord/MarkdownToWordWeb
   ```

2. Run the application:
   ```bash
   dotnet run
   ```

3. Open your browser and navigate to:
   ```
   http://localhost:5000
   ```

### Building for Production

To build the application for production:

```bash
cd MarkdownToWordWeb
dotnet publish -c Release -o ./publish
```

## Usage

### Method 1: Upload a Markdown File

1. Click the "Choose File" button
2. Select a Markdown file (.md, .markdown, or .txt)
3. Click "Convert to Word"
4. The Word document will be downloaded automatically

### Method 2: Enter Markdown Text

1. Scroll to the "Or Enter Markdown Text" section
2. Type or paste your Markdown content in the text area
3. Click "Convert Text to Word"
4. The Word document will be downloaded automatically

### Example

Try converting this markdown:

```markdown
# My Document

## Introduction

This is a **bold** statement and this is *italic*.

### Features

* First feature
* Second feature
* Third feature

### Steps

1. First step
2. Second step
3. Third step
```

## Project Structure

```
MarkdownToWord/
├── MarkdownToWordWeb/           # ASP.NET web application
│   ├── Pages/                   # Razor Pages
│   │   ├── Index.cshtml        # Main conversion page
│   │   └── Index.cshtml.cs     # Page logic
│   ├── Services/                # Business logic
│   │   └── MarkdownToWordConverter.cs  # Conversion service
│   ├── wwwroot/                 # Static files (CSS, JS, libs)
│   ├── Program.cs               # Application entry point
│   └── MarkdownToWordWeb.csproj # Project file
└── README.md                    # This file
```

## NuGet Packages

- **Markdig** (0.44.0): Markdown parsing
- **DocumentFormat.OpenXml** (3.3.0): Word document generation

## License

This project is open source and available under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

If you encounter any issues or have questions, please open an issue on GitHub.

