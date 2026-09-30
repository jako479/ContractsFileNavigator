# Contracts File Navigator

An Excel add-in that adds a **Worksheets** pane listing the sheets in `Contracts.xlsx`. The pane opens every time that file is opened. Double-click a name to jump to that sheet. Excel remembers how wide you made the pane.

This is the single-file edition of [Sheet Navigator](https://github.com/jako479/SheetNavigator). Install one or the other, not both: Contracts File Navigator stays idle and shows a message while Sheet Navigator is installed.

## Requirements

- Windows with desktop Microsoft Excel 2013 or later.
- .NET Framework 4.7.2 and the Visual Studio 2010 Tools for Office Runtime. The installer adds both if they are missing.

## Install

1. Download the latest release ZIP and extract it. Don't run the installer from inside the ZIP window.
2. Close Excel and run `setup.exe`. Click **Install** when Office asks.
3. Open `Contracts.xlsx`. The pane appears on the left.

## Use

- The pane opens whenever `Contracts.xlsx` is opened. Close it with its **X**; it comes back the next time the file is opened.
- Double-click a sheet in the pane to activate it.
- The pane width is remembered in your user profile; nothing is written to the workbook.
- To switch the add-in off without uninstalling it, set `Enabled` to `False` in its `user.config` and restart Excel. The file is under `%LOCALAPPDATA%\Microsoft_Corporation`, in the folder whose name starts with `ContractsFileNavigator.vsto`, and is created the first time the add-in runs.

## Uninstall

Windows Settings, Apps, **Contracts File Navigator**, Uninstall.

## Build

- Visual Studio 2022 or later with the **Office/SharePoint development** workload.
- Open `ContractsFileNavigator.slnx` and build the Release configuration. Ctrl+F5 starts Excel with the add-in loaded.
- ClickOnce signing needs a certificate, which is not in the repo. Before the first build, open Project Properties, **Signing**, and click **Create Test Certificate** (or choose your own).

## Release

1. In Visual Studio, right-click the project and choose **Publish**. Output goes to `ContractsFileNavigator\publish\`; `release\README.txt` and `LICENSE` are copied there automatically.
2. Zip the contents of `publish\` and attach it to a GitHub release.

## License

MIT. See [LICENSE](LICENSE).
