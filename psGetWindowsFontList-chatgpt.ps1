param (
    [switch]$show,
    [string]$fontList,
    [string]$fontNames
)

$objShell = New-Object -ComObject Shell.Application
$folder = "C:\Windows\Fonts"

$fileList = @()
$attrList = @{}
$details = @(
    "Title",
    "Font style",
    "Show/hide",
    "Designed for",
    "Category",
    "Designer/foundry",
    "Font Embeddability",
    "Font type",
    "Family",
    "Date created",
    "Date modified",
    "Collection",
    "Font file names",
    "Font version"
)

# Sort the details array
$details = $details | Sort-Object

# Figure out what the possible metadata is
$objFolder = $objShell.namespace($folder)
for ($attr = 0; $attr -le 500; $attr++)
{
    $attrName = $objFolder.getDetailsOf($objFolder.items, $attr)
    if ($attrName -and (-not $attrList.Contains($attrName)))
    {
        $attrList.add($attrName, $attr)
    }
}

# Loop through all the fonts and process
$objFolder = $objShell.namespace($folder)
foreach ($file in $objFolder.items())
{
    foreach ($attr in $details)
    {
        $attrValue = $objFolder.getDetailsOf($file, $attrList[$attr])

        if ($attr -eq "Date modified" -and $attrValue)
        {
            $attrValue = $attrValue -replace '[\p{C}]'
        }

        if ($attr -eq "Font file names" -and $attrValue -match "\\")
        {
            $attrValue = $attrValue -replace "\\", "/"
        }

        if ($attrValue)
        {
            Add-Member -InputObject $file -MemberType NoteProperty -Name $attr -Value $attrValue
        }
    }
    $fileList += $file
    Write-Verbose "Processing file number $($fileList.Count)"
}

# Sort fileList by Title attribute
$fileList = $fileList | Sort-Object -Property Title

if ($show)
{
    $fileList | Select-Object $details | Out-GridView
}
elseif ($fontList)
{
    if (-not $fontList)
    {
        Write-Host "Please provide a filename for the font list."
        exit
    }
    $fileList | Select-Object $details | ConvertTo-Json -Depth 3 | Out-File -FilePath $fontList
}
elseif ($fontNames)
{
    if (-not $fontNames)
    {
        Write-Host "Please provide a filename for the font names."
        exit
    }
    $fontNamesList = $fileList | Select-Object -ExpandProperty Name
    $fontNamesList | Out-File -FilePath $fontNames
}
else
{
    Write-Host "Usage: script.ps1 [-show] [-fontList <fontfilename>] [-fontNames <fontNamesfile>]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -show                      : Display font data using Out-GridView."
    Write-Host "  -fontList <fontfilename>   : Dump font data to a JSON file."
    Write-Host "  -fontNames <fontNamesfile> : Dump font names to a text file."
}
