# IMDB Rating Scraper for XML Files
# Processes directories containing XML files with IMDB tt numbers
# Adds rating and vote information to each XML file

param(
    [Parameter(Mandatory=$false)]
    [string]$RootPath = ".",
    [Parameter(Mandatory=$false)]
    [int]$DelaySeconds = 15
)

function Convert-VoteCount {
    param([string]$VoteString)
    
    if ($VoteString -match '^([\d.]+)K$') {
        return [int]([double]$Matches[1] * 1000)
    }
    elseif ($VoteString -match '^([\d.]+)M$') {
        return [int]([double]$Matches[1] * 1000000)
    }
    elseif ($VoteString -match '^\d+$') {
        return [int]$VoteString
    }
    else {
        # Remove commas and convert
        $cleaned = $VoteString -replace ',', ''
        if ($cleaned -match '^\d+$') {
            return [int]$cleaned
        }
        return $null
    }
}

function Get-IMDBRating {
    param([string]$TTNumber)
    
    try {
        $url = "https://www.imdb.com/title/$TTNumber/"
        Write-Host "  Fetching: $url" -ForegroundColor Cyan
        
        $headers = @{
            'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            'Accept' = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
            'Accept-Language' = 'en-US,en;q=0.5'
        }
        
        $response = Invoke-WebRequest -Uri $url -Headers $headers -TimeoutSec 30
        $html = $response.Content
        
        # First try to extract from JSON-LD structured data (more reliable and exact)
        $jsonLdMatch = [regex]::Match($html, '<script type="application/ld\+json">(.*?)</script>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
        if ($jsonLdMatch.Success) {
            try {
                $jsonContent = $jsonLdMatch.Groups[1].Value
                $jsonData = $jsonContent | ConvertFrom-Json
                
                if ($jsonData.aggregateRating) {
                    $rating = $jsonData.aggregateRating.ratingValue
                    $votes = $jsonData.aggregateRating.ratingCount
                    
                    if ($rating -and $votes) {
                        Write-Host "  Found (JSON-LD): Rating $rating, Votes $votes" -ForegroundColor Green
                        return @{
                            Rating = $rating.ToString()
                            Votes = [int]$votes
                        }
                    }
                }
            }
            catch {
                Write-Host "  Error parsing JSON-LD data: $($_.Exception.Message)" -ForegroundColor Yellow
                # Fall back to HTML parsing
            }
        }
        
        # Fallback to HTML parsing if JSON-LD method fails
        Write-Host "  JSON-LD method failed, trying HTML parsing..." -ForegroundColor Yellow
        
        # Extract rating using regex pattern for the specific class
        $ratingMatch = [regex]::Match($html, '<span class="sc-4dc495c1-1[^"]*"[^>]*>([0-9.]+)</span>')
        $rating = if ($ratingMatch.Success) { $ratingMatch.Groups[1].Value } else { $null }
        
        # Extract vote count using regex pattern for the vote class
        $voteMatch = [regex]::Match($html, '<div class="sc-4dc495c1-3[^"]*"[^>]*>([^<]+)</div>')
        $voteString = if ($voteMatch.Success) { $voteMatch.Groups[1].Value.Trim() } else { $null }
        
        if ($rating -and $voteString) {
            $votes = Convert-VoteCount $voteString
            if ($votes) {
                Write-Host "  Found (HTML): Rating $rating, Votes $votes" -ForegroundColor Green
                return @{
                    Rating = $rating
                    Votes = $votes
                }
            }
        }
        
        Write-Host "  Could not extract rating/votes from page" -ForegroundColor Yellow
        return $null
    }
    catch {
        Write-Host "  Error fetching IMDB data: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}


function Update-XMLRating {
    param(
        [string]$XmlPath,
        [hashtable]$RatingData
    )
    
    try {
        # Load XML document
        [xml]$xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.Load($XmlPath)
        
        # Find or create ratings node
        $ratingsNode = $xmlDoc.SelectSingleNode("//ratings")
        if (-not $ratingsNode) {
            $ratingsNode = $xmlDoc.CreateElement("ratings")
            $xmlDoc.DocumentElement.AppendChild($ratingsNode) | Out-Null
        }
        
        # Remove existing IMDB rating if it exists
        $existingImdbRating = $ratingsNode.SelectSingleNode("rating[@name='imdb']")
        if ($existingImdbRating) {
            $ratingsNode.RemoveChild($existingImdbRating) | Out-Null
        }
        
        # Create new rating element
        $ratingElement = $xmlDoc.CreateElement("rating")
        $ratingElement.SetAttribute("name", "imdb")
        $ratingElement.SetAttribute("default", "true")
        $ratingElement.SetAttribute("max", "10")
        
        # Add value element
        $valueElement = $xmlDoc.CreateElement("value")
        $valueElement.InnerText = $RatingData.Rating
        $ratingElement.AppendChild($valueElement) | Out-Null
        
        # Add votes element
        $votesElement = $xmlDoc.CreateElement("votes")
        $votesElement.InnerText = $RatingData.Votes
        $ratingElement.AppendChild($votesElement) | Out-Null
        
        # Add to ratings node
        $ratingsNode.AppendChild($ratingElement) | Out-Null
        
        # Save XML document with UTF-8 encoding
        $xmlDoc.PreserveWhitespace = $true
        $xmlWriterSettings = New-Object System.Xml.XmlWriterSettings
        $xmlWriterSettings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $xmlWriterSettings.Indent = $true
        $xmlWriterSettings.IndentChars = "  "
        $xmlWriterSettings.NewLineChars = "`r`n"
        $xmlWriterSettings.OmitXmlDeclaration = $false
        
        $xmlWriter = [System.Xml.XmlWriter]::Create($XmlPath, $xmlWriterSettings)
        $xmlDoc.Save($xmlWriter)
        $xmlWriter.Close()
        
        Write-Host "  Updated NFO file successfully with UTF-8 encoding" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  Error updating NFO: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}


function Process-Directory {
    param([string]$DirectoryPath)
    
    Write-Host "`nProcessing directory: $DirectoryPath" -ForegroundColor Magenta
    
    # Find NFO files in the directory
    $nfoFiles = Get-ChildItem -Path $DirectoryPath -Filter "*.nfo" -File
    
    if ($nfoFiles.Count -eq 0) {
        Write-Host "  No NFO files found" -ForegroundColor Yellow
        return $false  # Return false to indicate no processing was done
    }
    
    $processedAny = $false
    
    foreach ($nfoFile in $nfoFiles) {
        Write-Host "  Processing NFO: $($nfoFile.Name)" -ForegroundColor White
        
        try {
            # Load and parse NFO (XML format)
            [xml]$xml = Get-Content $nfoFile.FullName -Encoding UTF8
            
            # Find the IMDB unique ID
            $uniqueidNode = $xml.SelectSingleNode("//uniqueid[@type='imdb']")
            
            if (-not $uniqueidNode) {
                Write-Host "    No IMDB unique ID found" -ForegroundColor Yellow
                continue
            }
            
            $ttNumber = $uniqueidNode.InnerText.Trim()
            
            if (-not $ttNumber -or -not $ttNumber.StartsWith("tt")) {
                Write-Host "    Invalid IMDB ID: $ttNumber" -ForegroundColor Yellow
                continue
            }
            
            Write-Host "    Found IMDB ID: $ttNumber" -ForegroundColor Cyan
            
            # Check if IMDB rating already exists
            $existingRating = $xml.SelectSingleNode("//ratings/rating[@name='imdb']")
            if ($existingRating) {
                $existingValue = $existingRating.SelectSingleNode("value")
                $existingVotes = $existingRating.SelectSingleNode("votes")
                if ($existingValue -and $existingVotes -and $existingValue.InnerText -and $existingVotes.InnerText) {
                    Write-Host "    IMDB rating already exists (Rating: $($existingValue.InnerText), Votes: $($existingVotes.InnerText)) - skipping" -ForegroundColor Yellow
                    continue
                }
            }
            
            # Get rating from IMDB
            $ratingData = Get-IMDBRating $ttNumber
            
            if ($ratingData) {
                # Update the NFO file
                $success = Update-XMLRating $nfoFile.FullName $ratingData
                if ($success) {
                    Write-Host "    Successfully updated $($nfoFile.Name)" -ForegroundColor Green
                    $processedAny = $true
                }
            }
            else {
                Write-Host "    Failed to get rating data for $ttNumber" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "    Error processing $($nfoFile.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    return $processedAny  # Return true if any files were actually processed
}

# Main execution
Write-Host "IMDB Rating Scraper Started" -ForegroundColor Green
Write-Host "Root Path: $RootPath" -ForegroundColor Gray
Write-Host "Delay between directories: $DelaySeconds seconds" -ForegroundColor Gray
Write-Host "=" * 50

# Get all directories recursively
$directories = Get-ChildItem -Path $RootPath -Directory -Recurse

if ($directories.Count -eq 0) {
    Write-Host "No directories found in $RootPath" -ForegroundColor Red
    exit 1
}

Write-Host "Found $($directories.Count) directories to process (including subdirectories)`n" -ForegroundColor Green

# Process each directory with delay only when work was done
for ($i = 0; $i -lt $directories.Count; $i++) {
    $dir = $directories[$i]
    
    Write-Host "[$($i + 1)/$($directories.Count)]" -NoNewline -ForegroundColor Blue
    $didWork = Process-Directory $dir.FullName
    
    # Add delay only if we actually processed files AND it's not the last directory
    if ($didWork -and $i -lt ($directories.Count - 1)) {
        Write-Host "`nWaiting $DelaySeconds seconds before next directory..." -ForegroundColor Gray
        Start-Sleep -Seconds $DelaySeconds
    }
    elseif (-not $didWork) {
        Write-Host "  Skipping delay (no files processed)" -ForegroundColor Gray
    }
}

Write-Host "`n" + "=" * 50
Write-Host "IMDB Rating Scraper Completed" -ForegroundColor Green