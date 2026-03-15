# Winget Lab

Winget Lab is a .NET 8 WPF utility for curating `winget` package installations, previewing Microsoft Store entries, and generating repeatable installation scripts plus JSON manifests.

## Prerequisites
- Windows 10/11 with the `winget` CLI available on `PATH`
- .NET 8 SDK + Visual Studio 2022 17.8 (or newer) for building `WingetLab.slnx`
- Internet access (optional) to stream the Bing wallpaper shown in `WingetLab.xaml`

## Key Features
- **Package discovery** – runs `winget search`/`winget list` via worker scripts (`WingetLab.xaml.vb`) and shows the formatted output in `SearchResultsListBox`.
- **Source filtering & Store links** – the custom filter popup (`SourceFilterPanel`) scopes results to specific repositories and surfaces a Microsoft Store button/context menu when `msstore` packages are selected.
- **Curated selection workspace** – double-click or press `Enter` to add packages, dedupe detection warns on name collisions, and the `SelectedSearchTextBox` provides instant filtering across name/id/source/match metadata.
- **Script generation** – `BuildButton_Click` writes an interactive `.cmd` installer, a structured log, and a JSON manifest describing the chosen packages plus build options.
- **Start menu + shortcut automation** – optional toggles delete desktop `.lnk` files and capture/reenforce the Start Menu layout by copying `StartMenuLayout.bin` into the Windows shell data folder.
- **Configuration import/export** – `File > Open` reloads a JSON manifest back into `_selectedPackages`, enabling edits and rebuilds; `File > New` resets the state.
- **Environment awareness** – the UI picks up system dark mode, pulls Bing’s daily wallpaper, and exposes a live winget activity light (`WingetActivityLight`) so long-running commands are obvious.

## Building & Running
1. Open `WingetLab.slnx` (or `Winget Lab/WingetLab.vbproj`) in Visual Studio and build the `net8.0-windows` project.
2. Press F5 to launch the WPF window; it will auto-focus the search box and start fetching the Bing background.

## Using the App
1. **Search or list packages** – type a query and press `Enter`, or choose `Tools > Get Current List` to inventory installed software.
2. **Refine results** – use the source filter dropdown or right-click a result to open “More detail” when Microsoft Store metadata is available.
3. **Build your list** – double-click entries to add them; remove with the dedicated button, Delete key, or context selection. Use the selected search box to find items inside the curated list.
4. **Set build options** – provide the project/script name and opt into shortcut cleanup or Start Menu replication.
5. **Generate artifacts** – click `Build Install Script` to create:
   - `<name>.cmd` – interactive installer with logging, success/error counts, and optional Start Menu deployment.
   - `<name>_install.log` – cumulative log referenced in the script.
   - `<name>.json` – manifest storing package metadata and selected options for later reloads.
   - `StartMenuLayout.bin` – produced when capture is requested and later copied into the shell experience host.
6. The working folder opens in Explorer after a successful build.

## Tooling Shortcuts
- `File > New` resets the UI (`ClearAll`).
- `File > Open` loads a previously generated JSON manifest.
- `Tools > Get Current List` populates results with `winget list` output.
- `Tools > Capture Start Menu` saves the current Start Menu layout without building a script.

## Troubleshooting
- Ensure `winget` can run without elevation inside `cmd.exe`; the app shells out to `WorkerFile.cmd`/`WorkerOutput.txt` in the working directory.
- If no results show, verify network connectivity and that `winget` sources are synced; errors surface in the progress indicator at the bottom of the window.
 
