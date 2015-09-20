# How to run the script
# GenerateAppPackage.ps1 -ProjectFile -OutputPath -VisualStudioVersion -BuildConfiguration
 
# Define input parameters
Param(
    [Parameter(Mandatory = $true)] [String]$ProjectFile,         # Point to the .csproj file of the project you want to deploy
	[Parameter(Mandatory = $true)] [String]$OutputPath,
	[Parameter(Mandatory = $true)] [String]$VisualStudioVersion,
	[Parameter(Mandatory = $true)] [String]$BuildConfiguration
)
# Begin - Actual script -----------------------------------------------------------------------------------------------------------------------------
 
# Set the output level to verbose and make the script stop on error
$VerbosePreference = "Continue"
$ErrorActionPreference = "Stop"
$scriptPath = Split-Path -parent $PSCommandPath
Write-Output $scriptPath
# Mark the start time of the script execution
$startTime = Get-Date
 # Build and publish the project via web deploy package using msbuild.exe 

Write-Verbose ("[Start] App Package creation for project {0}" -f $ProjectFile)

# Run MSBuild to publish the project
& "$env:windir\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" /t:Package $ProjectFile /p:Configuration=$BuildConfiguration /p:OutputPath=$OutputPath /p:VisualStudioVersion=$VisualStudioVersion

Write-Verbose ("[Finish] App Package creation for project {0}" -f $ProjectFile)



# Mark the finish time of the script execution
$finishTime = Get-Date

# Output the time consumed in seconds
Write-Output ("Total time used (seconds): {0}" -f ($finishTime - $startTime).TotalSeconds)


# End - Actual script -----------------------------------------------------------------------------------------------------------