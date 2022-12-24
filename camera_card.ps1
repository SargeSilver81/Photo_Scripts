$shell = New-Object -ComObject Shell.Application
function Get-File-Date {
    [CmdletBinding()]
    Param (
        $object
    )

    $dir = $shell.NameSpace( $object.Directory.FullName )
    $file = $dir.ParseName( $object.Name )

    # First see if we have Date Taken, which is at index 12
    $date = Get-Date-Property-Value $dir $file 12

    if ($null -eq $date) {
        # If we don't have Date Taken, then find the oldest date from all date properties
        0..287 | ForEach-Object {
            $name = $dir.GetDetailsof($dir.items, $_)

            if ( $name -match '(date)|(created)') {
            
                # Only get value if date field because the GetDetailsOf call is expensive
                $tmp = Get-Date-Property-Value $dir $file $_
                if ( ($null -ne $tmp) -and (($null -eq $date) -or ($tmp -lt $date))) {
                    $date = $tmp
                }
            }
        }
    }
    return $date
}

function Get-Date-Property-Value {
    [CmdletBinding()]

    Param (
        $dir,
        $file,
        $index
    )

    $value = ($dir.GetDetailsof($file, $index) -replace "`u{200e}") -replace "`u{200f}"
    $value = $value.TrimStart()
    if ($value -and $value -ne '') {
        return [DateTime]::ParseExact(($value -replace "[^0-9/\:\s]"), 'dd/MM/yyyy HH:mm', $null)
    }
    return $null
}

# Find SD Card Drive letter
$SDPath = Get-CimInstance -ClassName Win32_Volume | ? {$_.DriveType -eq 2} | % DriveLetter

# Set Source and Destinations
$sourcePath = $SDPath
$destBasePath = 'C:\Users\'+$env:UserName+'\OneDrive\'
$destPathRAW = $destBasePath+'Pictures\RAW_Import'
$destPathJPG = $destBasePath+'Pictures\Nikon_Z6\JPG'
$destPathVIDEO = $destBasePath+'Videos\Z6_Videos'

# Lightroom import Directory
$LRPath = 'P:\LR_To_Import\'

# Reset Calculated Variables
$destPath = ''
$SortFolder = ''

# Test Source exists
if(!(Test-Path -Path $sourcePath)){
    Write-Host "No QXD Card Found!" -ForegroundColor Red
} else {
    # Lets Do It!!
    Write-Host "## RAW Camera Import ##" -ForegroundColor Yellow
    $Total = ( Get-ChildItem $sourcePath -Recurse -filter *.NEF | Measure-Object ).Count;
    $lc = 0

    if($Total -eq 0) {
        Write-Host "No RAW Files Found!" -ForegroundColor Red
    } else {

        Get-ChildItem $sourcePath -Recurse -filter *.NEF | ForEach-Object {
            ++$lc
            Write-Host "Moving ($lc/$Total): " $_.FullName
            $SortFolder = ''

            $DateTaken = Get-File-Date $_
            #Write-Host $_.FullName " was taken on: $DateTaken"
            $SortFolder = ($DateTaken).tostring("yyyy_MMM")
            #Write-Host "Folder derived from date taken: $SortFolder"           
            $destPath = $destPathRAW+'\'+$SortFolder
            #Write-Host "To Directory: " $destPath
            if((Test-Path -Path $LRPath )){
                Copy-Item -Path $_.FullName -Destination $LRPath -Force
            }
            if(!(Test-Path -Path $destPath )){
                New-Item -ItemType directory -Path $destPath
            }
            Move-Item $_.FullName -Destination $destPath -ErrorAction SilentlyContinue
        }
    }

    Write-Host "## JPG Camera Import ##" -ForegroundColor Yellow
    $Total = ( Get-ChildItem $sourcePath -Recurse -filter *.JPG | Measure-Object ).Count;
    $lc = 0

    if($Total -eq 0) {
        Write-Host "No JPG Files Found!" -ForegroundColor Red
    } else {

        Get-ChildItem $sourcePath -Recurse -filter *.JPG | ForEach-Object {
            ++$lc
            Write-Host "Moving ($lc/$Total): " $_.FullName
            $SortFolder = ''
                        
            $DateTaken = Get-File-Date $_
            #Write-Host $_.FullName " was taken on: $DateTaken"
            $SortFolder = ($DateTaken).tostring("yyyy_MMM")
            #Write-Host "Folder derived from date taken: $SortFolder"           
            $destPath = $destPathJPG+'\'+$SortFolder
            #Write-Host "To Directory: " $destPath
            if(!(Test-Path -Path $destPath )){
                New-Item -ItemType directory -Path $destPath
            }
            Move-Item $_.FullName -Destination $destPath -ErrorAction SilentlyContinue
        }
    }

    Write-Host "## Video Import ##" -ForegroundColor Yellow
    $Total = ( Get-ChildItem $sourcePath -Recurse -filter *.MP4 | Measure-Object ).Count;
    $lc = 0

    if($Total -eq 0) {
        Write-Host "No Video Files Found!" -ForegroundColor Red
    } else {

        Get-ChildItem $sourcePath -Recurse -filter *.MP4 | ForEach-Object {
            ++$lc
            Write-Host "Moving ($lc/$Total): " $_.FullName
            $SortFolder = ''
                        
            $DateTaken = Get-File-Date $_
            #Write-Host $_.FullName " was taken on: $DateTaken"
            $SortFolder = ($DateTaken).tostring("yyyy_MMM")
            #Write-Host "Folder derived from date taken: $SortFolder"           
            $destPath = $destPathVIDEO+'\'+$SortFolder
            #Write-Host "To Directory: " $destPath
            if(!(Test-Path -Path $destPath )){
                New-Item -ItemType directory -Path $destPath
            }
            Move-Item $_.FullName -Destination $destPath -ErrorAction SilentlyContinue
        }
    }
}