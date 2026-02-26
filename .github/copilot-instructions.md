# Copilot Instructions — SPO Search Probe

## Build Commands

```powershell
# Debug build (requires .NET 10 SDK)
dotnet build SPOSearchProbe

# Full release build with prerequisite checks (x64 + arm64)
.\Build.ps1
```

There are no tests or linters in this project.

## Architecture

A **.NET 10 Windows Forms** application that monitors SharePoint Online search index freshness. Authenticates via OAuth2 PKCE, executes scheduled KQL queries against the SharePoint REST search API, and logs results with full request/response archiving.

### Two Launch Modes

- **End-User mode** (default) — `MainForm.cs` — Single-user GUI. Requires a pre-configured `search-config.json` next to the exe.
- **Admin mode** (`-admin` flag) — `AdminForm.cs` — Multi-user management, page validation tracking, EndUser package creation, full config editor. Uses an extended `AdminConfig` model (defined inside `AdminForm.cs`) with user entries.

Both modes share the same core services and are selected in `Program.cs` based on CLI args.

### Core Services

| File | Responsibility |
|---|---|
| `SearchConfig.cs` | JSON-serializable config model (`search-config.json`). Has a `GetIntervalMs()` helper that converts `intervalValue`/`intervalUnit` to milliseconds via switch expression. |
| `OAuthHelper.cs` | OAuth2 PKCE flow via local `HttpListener` on ports 18700–18799. Uses `CancellationTokenSource` with 120s timeout and linked tokens for auth callbacks. |
| `TokenCache.cs` | Persists `TokenData` (access + refresh tokens) encrypted with Windows DPAPI (`DataProtectionScope.CurrentUser`). Load failures return `null` silently. |
| `SearchClient.cs` | Executes GET requests against `/_api/search/query` with OData verbose JSON. Logs each request/response pair as a ZIP archive. Sanitizes filenames for zip entries. |
| `LogCollector.cs` | Packages `.log` and `.tsv` files plus request ZIPs into a single archive. Generates a standalone HTML report with an embedded canvas chart by parsing TSV data. |
| `LiveChartForm.cs` | Real-time GDI+ chart popup using a `DoubleBufferedPanel` custom control to prevent flicker. Supports stacked per-user sub-charts (8 predefined colors). |

### Data Flow

1. User authenticates → `OAuthHelper.InteractiveLoginAsync` opens browser, listens for redirect → exchanges code for tokens
2. Tokens encrypted via DPAPI → stored as `.token-user-{email}.dat`
3. On each timer tick, `OAuthHelper.GetCachedOrRefreshedTokenAsync` returns a valid access token (auto-refreshes if within 5 min of expiry)
4. `SearchClient.ExecuteSearchAsync` calls SPO REST API → parses `d.query.PrimaryQueryResult.RelevantResults`
5. Results logged to TSV + `.log` file; request/response pairs archived as ZIP in `logs/requests/`
6. Optional page validation: resolves a URL to a `WorkId` via `path:` query, then monitors if that WorkId appears in scheduled query results

## Key Conventions

- **Single namespace**: All types live in `namespace SPOSearchProbe;` (file-scoped).
- **No NuGet dependencies**: Zero package references — all functionality uses BCL types only. Do not add NuGet packages.
- **JSON serialization**: `System.Text.Json` with `[JsonPropertyName]` attributes (camelCase). No Newtonsoft.
- **UI layout**: All controls created programmatically in constructors (no `.Designer.cs`, no `.resx`). Use `SuspendLayout()`/`ResumeLayout()`, manual `Location`/`Size` properties, and `Anchor` for responsive resizing.
- **Dual timer system**: Both forms use two `System.Windows.Forms.Timer` instances — one for search execution, one for countdown display.
- **Error handling**: File I/O and token operations use catch-all blocks that return `null` or silently continue. Do not add `throw` to these paths without careful consideration — silent failure is intentional for resilience.
- **Thread safety**: UI updates from async code use the `if (InvokeRequired) { Invoke(...); return; }` pattern.
- **Logging**: Dual output — `RichTextBox` in the UI (color-coded) and plain-text `.log` file. Structured data goes to a parallel `.tsv` file with fixed columns.
- **Single-instance enforcement**: Named `Mutex` per mode (`SPOSearchProbe_Admin_SingleInstance` / `SPOSearchProbe_EndUser_SingleInstance`).
- **Versioning**: `Build.ps1` computes `v1.YY.MDD.Build` from the current date and a `.build-counter` file.
- **Publishing**: Always self-contained single-file (`PublishSingleFile=true`, `--self-contained true`). Output goes to `bin/publish/{runtime}/`, distribution ZIPs to `dist/`.
- **Token files and config preserved across builds**: `Build.ps1` backs up `search-config.json` and `.dat` token files to a `../test/` folder before cleaning `bin/obj`, then restores them after publish.
