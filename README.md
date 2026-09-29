# BGWidget

BGWidget is a lightweight Windows desktop blood-glucose widget written in **Visual Basic 6**. It reads glucose history directly from a **Juggluco** HTTP endpoint and displays current glucose information, trend/delta data, long-term average information, estimated HbA1c, history graphs, and last-reading age in a transparent desktop widget.

The source in this repository is the working VB6 project, with the shared class paths adjusted only so the repository is self-contained when cloned.

## Features

- Reads Juggluco data from the `/x/stream` endpoint.
- Displays current glucose in mg/dL.
- Shows the change from the previous valid reading.
- Uses Juggluco rate data to display a trend arrow.
- Ignores zero glucose samples for the active reading and long-term accumulator.
- Calculates average glucose from up to 90 days of available history.
- Displays an estimated HbA1c derived from average glucose.
- Maintains an in-memory history cache for the graph.
- Configurable graph range, glucose scale, line/grid colors, and line width.
- Configurable colors and visibility for the individual widget fields.
- Can display the last reading as a clock time or relative time such as `5 Mins Ago`.
- Supports placement on a selected monitor or mirrored widgets across all monitors.
- Writes `BGWidget_A1c.log` when the long-term calculation runs, containing diagnostic information about the data used.

## Requirements

- Windows
- Visual Basic 6.0 if compiling from source
- Juggluco with its HTTP server enabled and reachable from the PC
- Windows COM components used by the project, including MSXML 6.0 and standard Windows scripting/ADO components

The VB6 project does not reference third-party OCX controls.

## Repository layout

```text
BGWidget/
|-- Project/
|   |-- BGWidget.vbp
|   |-- frmSplash.frm
|   |-- frmSplash.frx
|   |-- modGetGlucose.bas
|   |-- modHistoryGraph.bas
|   `-- modWidgetPlacement.bas
|-- SharedClasses/
|   |-- INIFile.cls
|   `-- ReadWrite.cls
|-- BGWidget.example.ini
|-- .gitignore
`-- README.md
```

Open `Project/BGWidget.vbp` in Visual Basic 6 to work with the source.

## Configuration

BGWidget reads `BGWidget.ini` from the same directory as `BGWidget.exe`.

Copy `BGWidget.example.ini` to `BGWidget.ini` and edit at least this value:

```ini
[Main]
DataPath=http://YOUR_PHONE_IP:17580
```

`DataPath` is the base address of the Juggluco HTTP server. BGWidget appends the appropriate `/x/stream` request itself.

The real `BGWidget.ini` is intentionally excluded by `.gitignore` so a private LAN or VPN address is not accidentally committed to GitHub.

### Main settings

- `GetGlucoseInterval` — polling interval in seconds.
- `IntervalToCalculateA1cInMinutes` — interval for recalculating the long-term values.
- `ShowGlucose` — show/hide current glucose.
- `ShowAverageGlucose` — show/hide long-term average glucose.
- `ShowA1c` — show/hide estimated HbA1c.
- `ShowDaysOfRecords` — show/hide the data-day count.
- `ShowLastUpdate` — show/hide last-reading time.
- `UseRelativeTime` — show last reading as relative time instead of clock time.

Colors are six-digit RGB hex values such as `FFFFFF`.

### History graph

The `[HistoryGraph]` section controls whether the graph is shown, the number of hours retained for the graph, its glucose range, colors, and line width.

### Multi-monitor placement

The `[Placement]` section controls monitor selection and widget position. `Monitor=all` creates mirrored widgets on all detected displays. Supported position names are:

- `topleft`
- `topright`
- `bottomleft`
- `bottomright`
- `center`

## Glucose and long-term calculations

BGWidget requests Juggluco history and parses the returned stream records. A glucose value must be numeric and greater than zero before it is accepted into the working accumulator.

The widget currently calculates estimated HbA1c from mean glucose with the formula present in the source:

```text
HbA1c = (AverageGlucose + 46.7) / 28.7
```

The source also calculates GMI internally. These calculated values are estimates derived from CGM data and are not laboratory HbA1c measurements.

## Diagnostic log

When the long-term calculation runs, BGWidget overwrites `BGWidget_A1c.log` beside the executable. The log includes the requested history range, accepted/rejected records, oldest/newest accepted timestamps, calculated average, and calculated HbA1c. Log files are excluded from Git by default.

## Building

1. Clone or download the repository.
2. Open `Project/BGWidget.vbp` in Visual Basic 6.
3. Compile `BGWidget.exe`.
4. Copy `BGWidget.example.ini` beside the executable as `BGWidget.ini`.
5. Set `DataPath` to the address of the phone running Juggluco.
6. Start BGWidget.

The project uses late-bound COM objects for network/file operations, so there are no additional third-party source libraries included in the project.

## Privacy

BGWidget communicates with the Juggluco address configured in your local INI file. No server address is hard-coded in the source in this repository.

Be careful when sharing diagnostic logs: `BGWidget_A1c.log` records the configured request URL along with statistics about the glucose-history response.

## Medical notice

BGWidget is a personal information/display utility. Its calculated averages and HbA1c estimate should not be treated as a laboratory result or used by themselves for medical decisions.

## License

No open-source license has been selected yet. Until a license is added, normal copyright rules apply to the source code.
