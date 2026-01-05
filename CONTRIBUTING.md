# Contributing to MarkdownToWord

Thank you for your interest in contributing to MarkdownToWord! This document provides guidelines for contributing to the project.

## How to Contribute

### Reporting Issues

If you find a bug or have a suggestion for improvement:

1. Check if the issue already exists in the [GitHub Issues](https://github.com/qiuhaotc/MarkdownToWord/issues)
2. If not, create a new issue with a clear title and description
3. Include steps to reproduce (for bugs)
4. Include expected vs actual behavior

### Submitting Changes

1. Fork the repository
2. Create a new branch for your feature or bugfix:
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. Make your changes
4. Test your changes thoroughly
5. Commit your changes with clear, descriptive commit messages
6. Push to your fork
7. Submit a Pull Request

### Code Style

- Follow C# coding conventions
- Use meaningful variable and method names
- Add comments for complex logic
- Keep methods focused and concise

### Testing

- Test your changes locally before submitting
- Ensure the application builds without errors
- Test both file upload and text input methods
- Verify generated Word documents open correctly

### Pull Request Guidelines

- Provide a clear description of the changes
- Reference any related issues
- Include screenshots for UI changes
- Ensure all tests pass
- Keep PRs focused on a single feature or fix

## Development Setup

1. Clone the repository
2. Install [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
3. Navigate to `MarkdownToWordWeb` directory
4. Run `dotnet restore` to install dependencies
5. Run `dotnet build` to build the project
6. Run `dotnet run` to start the application

## Areas for Contribution

We welcome contributions in the following areas:

- **New Features**: Additional markdown syntax support, export options, etc.
- **UI Improvements**: Better styling, accessibility, responsiveness
- **Bug Fixes**: Any issues you encounter
- **Documentation**: Improvements to README, code comments, guides
- **Testing**: Unit tests, integration tests
- **Performance**: Optimization of conversion process

## Questions?

If you have questions about contributing, feel free to:
- Open an issue for discussion
- Contact the maintainers

Thank you for contributing to MarkdownToWord!
