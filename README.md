# MarkdownToWord

Easily convert Markdown files to Word documents using C# and Blazor WebAssembly - all processing happens locally in your browser!

![Web Interface](https://raw.githubusercontent.com/qiuhaotc/MarkdownToWord/refs/heads/main/docs/Index.png)

## Features

- 🔒 **Privacy First**: All conversions are processed locally in your browser. Your data is never uploaded to any server!
- 🌐 **Bilingual Support**: Built-in language switching between English and Chinese (中文)
- 📁 **File Upload**: Upload `.md`, `.markdown`, or `.txt` files for conversion
- 📝 **Text Input**: Paste or type markdown content directly in the browser
- 🎨 **Rich Formatting**: Full support for headers, bold, italic, lists, tables, links, images, code blocks, and more
- 💻 **Modern Web UI**: Clean, responsive interface built with Bootstrap
- ⚡ **Fast Conversion**: Uses Markdig and DocumentFormat.OpenXml for efficient processing
- 📦 **No Backend Required**: Pure client-side application - can be hosted as static files

## Supported Markdown Features

- ✓ Headings (# H1, ## H2, ### H3, etc.)
- ✓ Bold text (\*\*text\*\*)
- ✓ Italic text (\*text\*)
- ✓ Bullet lists (- item or * item)
- ✓ Numbered lists (1. item)
- ✓ Paragraphs with proper spacing
- ✓ **Tables** - Full table support with headers
- ✓ **Links** - [text](url) format with styled hyperlinks
- ✓ **Images** - ![alt](url) format - automatically downloads and embeds images from web URLs
- ✓ **Code blocks** - Fenced code blocks with ``` 
- ✓ **Inline code** - `code` format with styling
- ✓ **Blockquotes** - > quote format
- ✓ **Horizontal rules** - --- separator

## Technologies Used

- **Blazor WebAssembly**: Client-side web framework running on WebAssembly
- **ASP.NET Core 8.0**: .NET runtime
- **Markdig**: High-performance Markdown parser
- **DocumentFormat.OpenXml**: Create Word documents (.docx)
- **Bootstrap 5**: Responsive UI framework
- **C# in Browser**: All conversion logic runs client-side via WebAssembly

## Getting Started

### Prerequisites

- [.NET 10.0 SDK](https://dotnet.microsoft.com/download/dotnet/10.0) or later

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
   https://localhost:7144
   ```

Note: Since this is a Blazor WebAssembly application, the first load may take a moment as it downloads the .NET runtime to your browser.

### Building for Production

To build the application for production deployment as static files:

```bash
cd MarkdownToWordWeb
dotnet publish -c Release -o ./publish
```

The published files in `./publish/wwwroot` can be hosted on any static file hosting service like:
- GitHub Pages
- Netlify
- Vercel
- Azure Static Web Apps
- AWS S3 + CloudFront
- Any web server (nginx, Apache, IIS, etc.)

## Usage

### Language Selection

Switch between English and Chinese (中文) using the language buttons at the top of the page.

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
├── MarkdownToWordWeb/           # Blazor WebAssembly application
│   ├── Pages/                   # Razor components
│   │   └── Index.razor          # Main conversion page
│   ├── Services/                # Business logic
│   │   └── MarkdownToWordConverter.cs  # Conversion service
│   ├── Shared/                  # Shared components
│   │   └── MainLayout.razor     # App layout
│   ├── wwwroot/                 # Static files (CSS, JS, libs)
│   │   ├── index.html           # Entry HTML page
│   │   ├── js/                  # JavaScript files
│   │   └── css/                 # Stylesheets
│   ├── App.razor                # Root component
│   ├── Program.cs               # Application entry point
│   ├── _Imports.razor           # Global using statements
│   └── MarkdownToWordWeb.csproj # Project file
└── README.md                    # This file
```

## NuGet Packages

- **Markdig** (0.44.0): Markdown parsing
- **DocumentFormat.OpenXml** (3.3.0): Word document generation
- **Microsoft.AspNetCore.Components.WebAssembly** (8.0.0): Blazor WebAssembly framework

## How It Works

This application uses Blazor WebAssembly to run C# code directly in your browser:

1. **Client-Side Processing**: The entire conversion happens in your browser using WebAssembly. No data is sent to any server.
2. **Privacy & Security**: Your markdown content and generated Word documents never leave your computer.
3. **Offline Capable**: Once loaded, the application can work offline (except for downloading images from URLs).
4. **Image Handling**: Images referenced in markdown (via URLs) are downloaded and embedded into the Word document.

## Deployment

Since this is a static Blazor WebAssembly application, you can deploy it anywhere:

### GitHub Pages
1. Build: `dotnet publish -c Release -o ./publish`
2. Copy contents of `./publish/wwwroot` to your GitHub Pages repository
3. Add a `.nojekyll` file to prevent Jekyll processing

### Netlify/Vercel
1. Build: `dotnet publish -c Release -o ./publish`
2. Deploy the `wwwroot` folder from `./publish/wwwroot`

### Traditional Web Server
1. Build: `dotnet publish -c Release -o ./publish`
2. Copy contents of `./publish/wwwroot` to your web server
3. Configure server to handle client-side routing (all routes should serve `index.html`)

## License

This project is open source and available under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

If you encounter any issues or have questions, please open an issue on GitHub.

