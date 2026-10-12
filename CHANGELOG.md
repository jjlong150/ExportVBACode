# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0] - 2026-10-12

Version 2.0 focuses the export on what a reader of the published code needs, and
hardens the export against several failure modes in 1.0. The output layout and
contents have changed, so existing export folders will show renames and
deletions on the first commit after upgrading.

### Breaking changes

- **Component folders no longer contain spaces**, so documentation tools such as
  DeepWiki generate valid links:
  - `Class Modules` → `ClassModules`
  - `Microsoft Excel Objects` → `MicrosoftExcelObjects`
- **The unzipped workbook package is filtered.** `xl`, `docProps`,
  `[Content_Types].xml` and every `_rels` folder are removed after unzipping,
  leaving only content of interest such as the `customUI` ribbon XML. The list
  is set by `PACKAGE_EXCLUDES` in `modExport`.
- **Empty sheet and ThisWorkbook modules are no longer exported.** Document
  modules containing only blank lines and `Option` statements are skipped. Set
  `SKIP_EMPTY_DOCUMENT_MODULES` to `False` to export them.
- **Empty folders are removed** after export (for example `ActiveX` when the
  workbook has no designers).
- **The unzipped package folder is named after the file extension**
  (`src\xlsm`, `src\xlsb`, ...). Output for `.xlsm` workbooks is unchanged.
- **ActiveX designers are exported as `.dsr`**, the extension the VBA editor
  actually writes, instead of `.ocx`.

### Added

- **Pretty-printed ribbon XML.** A formatted copy of each `customUI\*.xml` file
  is written beside the original (for example `customUI14.pretty.xml`); the
  original file is never modified. Controlled by a checkbox on the Export
  worksheet linked to the `PRETTY_PRINT_XML` named cell.
- **Pre-flight checks** with clear messages: the workbook and export folder must
  exist, "Trust access to the VBA project object model" must be enabled, and the
  VBA project must not be password-locked.
- **Summary dialog** at the end of each export showing the number of components
  exported and skipped, package items removed, XML files pretty-printed, empty
  folders removed, and any warnings.
- **Support for more workbook types:** `.xlsb`, `.xlam`, `.xltm` and `.xls`, in
  addition to `.xlsm`. Formats that are not zip-based (`.xls`) skip the unzip
  step.
- **Configuration constants** at the top of `modExport` for package exclusions,
  pretty-print file patterns and suffix, empty-module skipping and component
  folder names.
- **`.gitattributes`** for the repository. VBA source files are stored exactly
  as exported (CRLF, ANSI) so downloaded modules import cleanly; `.frx` and
  workbook files are marked binary; `.bas`/`.cls`/`.frm` are highlighted as VBA
  on GitHub.

### Changed

- Only the generated `src` folder is cleared before each export. Other content
  in the export folder, including a `.git` folder or README, is no longer
  deleted.
- The workbook is unzipped with Windows' built-in `tar.exe`, which runs to
  completion before the export continues. The Shell.Application method is kept
  as a fallback and now waits for extraction to finish.
- The target workbook is opened read-only, without updating links, without
  adding it to the recent files list, and with events disabled so its
  `Workbook_Open` and `Workbook_BeforeClose` code does not run.
- The file picker defaults to the previously selected workbook, and the folder
  picker to the previously selected folder.
- Shell objects are late-bound, so no reference to "Microsoft Shell Controls and
  Automation" is required.

### Fixed

- The unzip step could produce partial output or fail, because the workbook copy
  was renamed back from `.zip` while Windows was still extracting it.
- The unzip destination could be rejected by Windows (error 91) because the
  folder path was passed to `Shell.Namespace` as a String instead of a Variant.
- Exporting a workbook that was already open closed it without saving,
  discarding unsaved changes. Exporting the exporter workbook itself closed the
  tool mid-run. An already-open workbook is now reused and left open, with a
  warning if it has unsaved changes.
- Cancelling the file picker cleared the previously selected workbook.
- An unrecognized component type was exported using the previous component's
  folder and file extension.
- An error partway through the export left a hidden workbook open, screen
  updating turned off and a stale status bar message.
- Folder creation errors were silently ignored, surfacing later as unrelated
  export failures.
- Clearing the export folder failed when it contained read-only files.
- "Done!" remained on the Excel status bar after the export finished.
- The internal component-type enum could conflict with the VBA Extensibility
  library when that reference was added to the project.

### Upgrading from 1.0

1. Import the new `modExport.bas`, replacing the old module.
2. On the Export worksheet, add a checkbox linked to a cell named
   `PRETTY_PRINT_XML`. If the name is missing, pretty-printing is off.
3. Add `.gitattributes` to the repository root. If VBA files are already
   committed, run `git add --renormalize .` and commit once.
4. Run an export and commit. Git will show the renamed component folders and the
   removed package files.

## [1.0] - 2025-06-21

- Initial release: exports a workbook's VBA components, a copy of the workbook
  and its unzipped package contents.

[2.0]: https://github.com/jjlong150/ExportVBACode/compare/1.0...2.0
[1.0]: https://github.com/jjlong150/ExportVBACode/releases/tag/1.0
