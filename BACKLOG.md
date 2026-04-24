# Backlog

## To-do

### Pivot Enhancements

- Extend `PIVOT.COUNT` with a `calculation` argument:
  - `0` = absolute counts (default)
  - `1` = row percentages
  - `2` = column percentages
  - `3` = percentage of the grand total

### Release Versioning

- Add a deliberate workflow for changing `major` and `minor` versions on top of the current automatic deployment bump of `patch` in `1.0.x`.
- Decide whether higher-level version changes should also update user-facing release notes, installer naming, and website download metadata automatically.


## Future Features

### Cross-Platform Support

- Add a Mac/Web-compatible version of `XLStatUDF` as an Office Add-in based on Office.js custom functions.
- Keep the current `Excel-DNA` implementation as the Windows-specific distribution.
- Reuse the existing statistical specification and documentation structure when designing the Office.js variant.
- Define a migration plan for the first batch of functions to port, prioritizing the most-used descriptive functions and hypothesis tests.
- Investigate how to share documentation, naming conventions, and test cases between the Windows and Office.js implementations.
- Confirm current platform limits for Excel custom functions on iPad/iOS before committing to any mobile roadmap.

