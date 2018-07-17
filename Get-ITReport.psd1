#################################
#
# Get-ITReport
#
# James Arnett (c) 2015
#
#################################

@{

# Script module or binary module file associated with this manifest
ModuleToProcess = 'Get-ITReport.psm1'

# Version number of this module.
ModuleVersion = '1.0.0.0'

# ID used to uniquely identify this module
GUID = 'd2c8f844-60a9-405f-92cc-82d5721646cc'

# Author of this module
Author = 'James Arnett'

# Company or vendor of this module
CompanyName = ''

# Copyright statement for this module
Copyright = '(c) 2015 James Arnett. All rights reserved.'

# Description of the functionality provided by this module
Description = 'PowerShell module to get various IT related reports.'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '2.0'

# Name of the Windows PowerShell host required by this module
PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
PowerShellHostVersion = ''

# Minimum version of the .NET Framework required by this module
DotNetFrameworkVersion = '2.0.50727'

# Minimum version of the common language runtime (CLR) required by this module
CLRVersion = ''

# Processor architecture (None, X86, Amd64, IA64) required by this module
ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module
ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
FormatsToProcess = @()

# Modules to import as nested modules of the module specified in ModuleToProcess
NestedModules = @()

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module
CmdletsToExport = 'Get-*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# List of all modules packaged with this module
ModuleList = @('Get-ITReport','Get-HardwareReport','Get-ShutdownLogReport','Get-UpTimeReport','Get-ServiceReport','Get-QFEReport')

# List of all files packaged with this module
FileList = @('Get-ITReport.psm1', 'Get-ITReport.psd1')

# Private data to pass to the module specified in ModuleToProcess
PrivateData = ''

}