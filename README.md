# IMDB Rating Scraper

A PowerShell script that automatically fetches IMDB ratings and vote counts for movies and updates NFO files with the retrieved data. Perfect for media center applications like Kodi, Plex, or Jellyfin.

## 🎯 Purpose

This script scans directories containing NFO files (XML format), extracts IMDB IDs from them, fetches the corresponding ratings from IMDB, and updates the NFO files with accurate rating and vote count information. It's designed to enhance your media library with up-to-date IMDB ratings.

## ✨ Features

- **Recursive Directory Processing**: Scans all subdirectories automatically
- **Precise Data Extraction**: Uses IMDB's JSON-LD structured data for exact ratings (not truncated display values)
- **Fallback Mechanism**: If JSON-LD parsing fails, falls back to HTML parsing
- **Smart Skip Logic**: Won't overwrite existing IMDB ratings in NFO files
- **Rate Limiting**: Configurable delays between directory processing to be respectful to IMDB
- **UTF-8 Encoding**: Properly handles international characters in NFO files
- **Comprehensive Logging**: Detailed progress reporting with color-coded messages

## 🚨 Important Notice

**This script will NOT update or overwrite existing IMDB ratings.** If an NFO file already contains valid IMDB rating data (both rating value and vote count), the script will skip that file entirely. This preserves any manual ratings or existing data you may have.

## 📋 Requirements

- Windows PowerShell 5.1 or PowerShell Core 6+
- Internet connection for IMDB access
- NFO files in XML format containing IMDB unique IDs

## 🚀 Usage

### Basic Usage
```powershell
.\updatenfo.ps1
```
This will process the current directory with a 15-second delay between directories.

### Custom Root Directory
```powershell
.\updatenfo.ps1 -RootPath "C:\Movies"
```

### Custom Delay
```powershell
.\updatenfo.ps1 -DelaySeconds 30
```

### Combined Parameters
```powershell
.\updatenfo.ps1 -RootPath "D:\Media\Movies" -DelaySeconds 10
```

## 📁 Expected File Structure

The script expects NFO files containing IMDB unique IDs in this format:

```xml
<?xml version="1.0" encoding="utf-8"?>
<movie>
    <title>Movie Title</title>
    <uniqueid type="imdb">tt1234567</uniqueid>
    <!-- Other movie data -->
</movie>
```

## 🔄 How It Works

1. **Directory Scanning**: Recursively finds all directories under the specified root path
2. **NFO Detection**: Locates `.nfo` files in each directory
3. **IMDB ID Extraction**: Parses XML to find `<uniqueid type="imdb">` elements
4. **Existing Rating Check**: Verifies if IMDB rating already exists (skips if found)
5. **Data Fetching**: Retrieves rating information from IMDB using two methods:
   - **Primary**: Extracts from JSON-LD structured data (most accurate)
   - **Fallback**: Parses HTML elements (if JSON-LD fails)
6. **NFO Update**: Adds rating information to the XML structure
7. **Rate Limiting**: Waits between directories that had files processed

## 📊 Output Format

The script adds rating data in this format:

```xml
<ratings>
    <rating name="imdb" default="true" max="10">
        <value>7.8</value>
        <votes>125487</votes>
    </rating>
</ratings>
```

## 🎨 Console Output

The script provides color-coded feedback:
- **Green**: Success messages and completion
- **Cyan**: IMDB URLs being fetched
- **Yellow**: Warnings and skipped items
- **Red**: Errors
- **Blue**: Progress indicators
- **Magenta**: Directory processing headers
- **White**: File processing status

## ⚙️ Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `RootPath` | String | `"."` | Root directory to start processing |
| `DelaySeconds` | Integer | `15` | Seconds to wait between directories |

## 🛡️ Error Handling

The script includes comprehensive error handling for:
- Invalid or missing IMDB IDs
- Network connectivity issues
- Malformed XML files
- File permission problems
- IMDB page structure changes

## 🚦 Rate Limiting

To be respectful to IMDB's servers, the script:
- Only applies delays between directories where files were actually processed
- Skips delays if no NFO files were found or processed
- Uses configurable delay intervals (default: 15 seconds)
- Includes proper HTTP headers to appear as a legitimate browser request

## 📝 Example Session

```
IMDB Rating Scraper Started
Root Path: C:\Movies
Delay between directories: 15 seconds
==================================================
Found 247 directories to process (including subdirectories)

[1/247]
Processing directory: C:\Movies\Action\John Wick (2014)
  Processing NFO: movie.nfo
    Found IMDB ID: tt2911666
  Fetching: https://www.imdb.com/title/tt2911666/
  Found (JSON-LD): Rating 7.4, Votes 578429
  Updated NFO file successfully with UTF-8 encoding
    Successfully updated movie.nfo

Waiting 15 seconds before next directory...

[2/247]
Processing directory: C:\Movies\Action\John Wick Chapter 2 (2017)
  Processing NFO: movie.nfo
    Found IMDB ID: tt4425200
    IMDB rating already exists (Rating: 7.5, Votes: 425123) - skipping
  Skipping delay (no files processed)
```

## 🤝 Contributing

Feel free to submit issues, feature requests, or pull requests to improve the script's functionality.

## 📜 License

This project is provided as-is for personal use. Please be respectful of IMDB's terms of service and rate limits when using this script.

## ⚠️ Disclaimer

This script is for personal use only. Users are responsible for complying with IMDB's terms of service and robots.txt policies. The script includes respectful delays and standard browser headers to minimize server impact.
