# XRai Support Matrix

This document lists the Excel versions, Windows versions, and Office bitness
combinations that XRai is tested against. Combinations not listed are
"best-effort" — they may work but are not covered by automated tests.

## Supported (tested on every release)

| Windows | Excel | Bitness | Status |
|---------|-------|---------|--------|
| Windows 11 23H2+ | Microsoft 365 (Current Channel) | x64 | Supported |
| Windows 10 22H2 | Microsoft 365 (Current Channel) | x64 | Supported |
| Windows 11 23H2+ | Excel 2021 | x64 | Supported |
| Windows 10/11 | Excel 2019 | x64 | Supported |

## Best-effort (no automated tests, community-reported)

| Windows | Excel | Bitness | Status |
|---------|-------|---------|--------|
| Windows 10/11 | Excel 2016 | x64 | Best-effort |
| Windows 10/11 | Microsoft 365 (Monthly Enterprise Channel) | x64 | Best-effort |
| Windows 10/11 | Microsoft 365 (Semi-Annual Channel) | x64 | Best-effort |

## Not supported

| Version | Reason |
|---------|--------|
| Excel 2013 and earlier | COM interop surface missing required APIs |
| 32-bit Office | XRai.Tool.exe is published for win-x64 only |
| Excel for Mac | No COM interop on macOS |
| Excel Online / web | Uses Office.js, not COM |
| Windows 7/8 | .NET 8 runtime not supported on these OS versions |

## Per-command capability matrix

Some commands require specific Excel features that are only available on
certain versions. When a command is called on an unsupported Excel version,
XRai returns `{"ok":false,"error":"not_supported","code":"XRAI_NOT_SUPPORTED_IN_EXCEL_VERSION","required":"Excel 2019+"}`.

| Command | Requires |
|---------|----------|
| `powerquery.*` | Excel 2016+ |
| `dax.*`, data model | Excel 2013+ (Power Pivot) |
| `slicer.create` on Tables | Excel 2013+ |
| `sparkline.*` | Excel 2010+ (all supported) |
| `comment.thread` (threaded comments) | Microsoft 365 only |
| `workbook.query.*` | Excel 2016+ |

## Runtime detection

XRai detects the Excel version at `connect` time and exposes it in the response:

```json
{"ok":true,"attached":true,"version":"16.0",...}
```

The `version` string follows Excel's internal versioning:
- `14.0` = Excel 2010
- `15.0` = Excel 2013
- `16.0` = Excel 2016, 2019, 2021, Microsoft 365 (all share the same major)

To distinguish Microsoft 365 from perpetual Office 2019/2021, check the
registry key `HKCU\Software\Microsoft\Office\ClickToRun\Configuration\ProductReleaseIds`
or use `excel.product_edition` command.

## Reporting unsupported combinations

If XRai fails on a combination you expect to work, file an issue at
https://github.com/xrai/xrai/issues with:
- Windows version (`winver`)
- Excel version + bitness (File > Account > About)
- Full error response from XRai including `code` and `stack_frame`
