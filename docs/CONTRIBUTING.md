# Contributing to M365 Dashboard

First off, thank you for considering contributing to M365 Dashboard! It's people like you that make this project better for everyone.

## Code of Conduct

By participating in this project, you are expected to uphold our Code of Conduct: be respectful, inclusive, and constructive.

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check existing issues. When you create a bug report, include as many details as possible:

- **Use a clear and descriptive title**
- **Describe the exact steps to reproduce the problem**
- **Provide specific examples** (code snippets, screenshots)
- **Describe the behavior you observed and what you expected**
- **Include your environment details** (OS, browser, .NET version, Node version)

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion:

- **Use a clear and descriptive title**
- **Provide a detailed description of the suggested enhancement**
- **Explain why this enhancement would be useful**
- **List any alternative solutions you've considered**

### Pull Requests

1. **Fork the repository** and create your branch from `main`
2. **Install dependencies**: 
   ```bash
   # Backend
   cd src/M365Dashboard.Api
   dotnet restore
   
   # Frontend
   cd ClientApp
   npm install
   ```
3. **Make your changes** following our coding standards
4. **Test your changes** thoroughly
5. **Commit your changes** with a clear commit message
6. **Push to your fork** and submit a pull request

## Development Setup

### Prerequisites

- .NET 8 SDK
- Node.js 20 LTS
- Azure CLI (for deployment)
- VS Code or Visual Studio 2022

### Local Development

1. Clone the repository:
   ```bash
   git clone https://github.com/YOUR_USERNAME/m365-dashboard.git
   cd m365-dashboard
   ```

2. Set up the backend:
   ```bash
   cd src/M365Dashboard.Api
   cp appsettings.Development.json.example appsettings.Development.json
   # Edit appsettings.Development.json with your Entra ID credentials
   dotnet restore
   dotnet run
   ```

3. Set up the frontend (in a new terminal):
   ```bash
   cd src/M365Dashboard.Api/ClientApp
   npm install
   npm start
   ```

4. Access the app at `https://localhost:5001`

### Running Tests

```bash
# Backend tests
cd src/M365Dashboard.Api
dotnet test

# Frontend tests
cd ClientApp
npm test
```

## Coding Standards

### C# / .NET

- Follow [Microsoft's C# Coding Conventions](https://docs.microsoft.com/en-us/dotnet/csharp/fundamentals/coding-style/coding-conventions)
- Use meaningful variable and method names
- Add XML documentation comments for public APIs
- Keep methods small and focused
- Use async/await for I/O operations

### TypeScript / React

- Use TypeScript for all new code
- Follow the existing ESLint configuration
- Use functional components with hooks
- Keep components small and reusable
- Use meaningful prop names

### Commits

- Use clear, concise commit messages
- Start with a verb in present tense: "Add", "Fix", "Update", "Remove"
- Reference issues when applicable: "Fix #123: Resolve login redirect issue"

### Pull Request Guidelines

- Keep PRs focused on a single feature or fix
- Update documentation if needed
- Add tests for new functionality
- Ensure all existing tests pass
- Request review from maintainers

## Project Structure

```
m365-dashboard/
├── src/M365Dashboard.Api/     # .NET 8 backend
│   ├── Controllers/           # API endpoints
│   ├── Services/              # Business logic
│   ├── Models/                # Data models
│   ├── Data/                  # EF Core context
│   └── ClientApp/             # React frontend
│       ├── src/
│       │   ├── components/    # React components
│       │   ├── hooks/         # Custom hooks
│       │   ├── services/      # API services
│       │   └── types/         # TypeScript types
├── infra/                     # Bicep templates
└── scripts/                   # Deployment scripts
```

## Questions?

Feel free to open an issue with your question or reach out to the maintainers.

Thank you for contributing! 🎉
