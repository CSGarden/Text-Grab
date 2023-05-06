$BuildPath = "$PSScriptRoot\bld\x86"
$Version = Get-Date -Format "yyyy-MM-dd" # 2020-11-1
$VersionDot = $Version -replace '-','.'
$Project = "Text-Grab"
$Archive = "$BuildPath\$Project-$Version.zip"

# Clean up
if(Test-Path -Path $BuildPath)
{
    Remove-Item $BuildPath -Recurse
}

# Dotnet restore and build
dotnet publish "$PSScriptRoot\$Project\$Project.csproj" `
	   --runtime win-x86 `
	   --self-contained false `
	   -c Release `
	   -v minimal `
	   -o $BuildPath `
	   -p:PublishReadyToRun=true `
	   -p:PublishSingleFile=true `
	   -p:CopyOutputSymbolsToPublishDirectory=false `
	   -p:Version=$VersionDot `
	   --nologo

# Archiv Build
Compress-Archive -Path "$BuildPath\$Project.exe" -DestinationPath $Archive
