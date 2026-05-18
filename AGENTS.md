# cs-office-mcp-server

A Model Context Protocol (MCP) server for operating Microsoft Office files (Excel, Word, PowerPoint, and Outlook) on Windows.

## Project Overview

- **Purpose:** Enables LLMs to interact with Office applications through the Model Context Protocol.
- **Target Platform:** Windows only (requires Office 2016+ 64-bit).
- **Tech Stack:**
  - **Language:** C# (.NET 8.0)
  - **Communication Protocol:** Model Context Protocol (MCP) using Standard I/O transport.
  - **Office Integration:** Microsoft Office Interop (via `Microsoft.Office.Interop.*` DLLs in `Lib/`).
  - **Server Framework:** `Microsoft.Extensions.Hosting`.

## Architecture

- **Server Entry Point:** `cs-office-mcp-server/Program.cs` initializes the MCP server and registers tools from the assembly.
- **Tools:** Implemented in `cs-office-mcp-server/Tools/`. 
  - Classes are static and marked with `[McpServerToolType]`.
  - Methods are marked with `[McpServerTool]` and `[Description]`.
- **Session Management:** Specialized session classes (e.g., `ExcelSession.cs`, `WordSession.cs`) handle COM object lifecycle and Office-specific logic.
- **Interop DLLs:** Pre-compiled interop libraries are located in the `Lib/` directory and referenced by the project.

## Building and Running

### Prerequisites
- Windows OS
- .NET 8.0 SDK
- Microsoft Office 2016 or later (64-bit version recommended)

### Build
```powershell
dotnet build
```

### Run
```powershell
dotnet run --project cs-office-mcp-server
```
*Note: The server uses `stdout` for MCP communication, so all logs are directed to `stderr`.*

### Test
```powershell
dotnet test
```
The test suite uses MSTest and relies on files in the `TestData/` directory.

## Development Conventions

- **Tool Implementation:**
  - Always use `[McpServerTool]` with a clear `Name` and `Description`.
  - Use `[Description]` on all parameters to provide context to the LLM.
  - Wrap Office interop calls in `Session` classes to ensure proper COM object disposal.
- **COM Object Safety:** Use `session.RegisterComObject()` or similar mechanisms provided by the session classes to track and release COM objects.
- **Error Handling:** Throw `McpException` for errors that should be reported back to the LLM.
- **Logging:** Use `ILogger` provided by the host. Ensure all output that is not part of the MCP protocol goes to `stderr`.
- **Testing:**
  - Add new tests in the `TestTools` project.
  - Use `TestBase` as a foundation for tests requiring access to `TestData/`.
  - Prefer `[DataRow]` for testing tools with various inputs.

## Key Directories

- `cs-office-mcp-server/Tools/`: Tool implementations.
- `cs-office-mcp-server/Properties/PublishProfiles/`: Deployment configurations.
- `Lib/`: Microsoft Office Interop assemblies.
- `TestData/`: Sample Office files for testing.
- `TestTools/`: MSTest unit tests.
