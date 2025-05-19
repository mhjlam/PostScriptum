# PostScriptum

Collection of (useful) Powershell scripts for video file management, subtitle handling, and file/folder renaming automation. Each script streamlines common media and file organization tasks.

## Scripts

- **autotrim.ps1**:     Scans video files for segments that are mostly black, white, or static, and automatically trims the video up to the detected segment.
- **fetch.ps1**:        Scans a folder for Git repositories and and pulls the latest changes from the remote, including submodules.
- **hevc.ps1**:         Converts video files to HEVC (H.265) format to save space while maintaining quality.
- **isodatify.ps1**:    Renames files and folders by converting any date formats in their names to the ISO standard (YYYY-MM-DD).
- **subs.ps1**:         Downloads (English) subtitles for video files using [`subliminal`](https://github.com/Diaoul/subliminal).

### Use Cases

- Batch-trim and compress video files for archiving or sharing.
- Standardize file and folder names by normalizing embedded dates to ISO format.
- Organize and synchronize subtitle files for your video library.
- Update a batch of repositories to their latest (development) versions.
