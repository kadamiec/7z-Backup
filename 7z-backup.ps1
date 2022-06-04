# 7z-based backup script v0.1
# 7z documentation: https://sevenzip.osdn.jp/chm/cmdline/

<#
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
7z-backup.ps1 -config 'config.cfg' -Verbode -Debug
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [Alias('config')]
    [string] $configFile,

    [Parameter()]
    [ValidateSet('None','Fast')]
    [String] $compression = "None",

    [Parameter()]
    [ValidateSet('Auto', 'FullNew', 'FullUpdate', 'Diff')]
    [String] $mode = "Auto",

    [Parameter()]
    [String] $password = "",

    [Parameter()]
    [boolean] $forceNoPassword = $false
)

Set-StrictMode -Version 2.0     # Throw exceptions when trying to access uninitialized variables or non-existent properties
$ErrorActionPreference  = "Stop" # Stop execution at exceptions/errors
$DebugPreference        = "Continue" # Stop execution at exceptions/errors

enum modeEnum { Auto; FullNew; FullUpdate; Diff; }
enum compEnum { None; Fast; }

# New-Object -TypeName System.Version -ArgumentList "1.2.3.4"

# Common functions
    function DoesFileExists {
        [OutputType([boolean])]
        Param (
            [Parameter()]
            $path,
            [Parameter()]
            $throwException = $False
        )

        if ([String]::IsNullOrEmpty($path)) {
            throw "Path is null"
        }

        if(-not (Test-Path -path $path -PathType Leaf)) {
            if ($throwException) {
                throw "File: ${path} does not exists!"
            }
            return $False;
        }
        return $True;
    }
    Function Get-StringHash 
    { 
        param
        (
            [String] $String,
            $HashName = "MD5"
        )
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($String)
        $algorithm = [System.Security.Cryptography.HashAlgorithm]::Create('MD5')
        $StringBuilder = New-Object System.Text.StringBuilder 
    
        $algorithm.ComputeHash($bytes) | 
        ForEach-Object { 
            $null = $StringBuilder.Append($_.ToString("x2")) 
        } 
    
        $StringBuilder.ToString() 
    }

    function QuoteStr {
    param (
        [Parameter(Mandatory)]
        [string] $str,
        [string] $chr = [char]34
    )
        # if ([string]::IsNullOrEmpty($str) -eq $true) { 
        #     return "";
        # } 
        return [string]::Format("{0}{1}{0}", $chr, $str);
    }

    function Save-ListAsFile($list, $toFile) {
        Out-File -FilePath $toFile -InputObject $list -Encoding UTF8
    }
#/ Common functions

class CmdArgs {
    $Command = "a"
    $Compression = $null
    $UpdateOptions = $null
    $ToBackup = $null
    $ToExclude = $null
}

class BackupConfig {
    $ConfigFilePath  = $null
    $HasBaseFile = $null
    $BackupMode  =  [modeEnum]::Auto
    $Compresssion = [compEnum]::None
    $BackupListFilePath = $null
    $ExcludeListFilePath = $null
    $BaseBackupFilePath = $null     # filename/path indicating source for $BaseBAckupArchivePath
    $BaseBackupArchivePath = $null  # target full path of the 7z file
    $BackupTargetDir = $null
    $ConfigName = $null
    $ArchiveName = $null
    $CmdArgs = [CmdArgs]::new()
    $BackupKey = $null # fallback to default

    hidden [hashtable] $cfg
    hidden [string] $ArchiveNameSuffix
    hidden [string] $Timestamp = $(get-date -f yyyyMMdd_HHmmss)
    # hidden [string] $baseBackupFile

    # Constructor - initialize from config file
    BackupConfig([String] $configFilePath, $mode, $compression) {
        $this.configFilePath = $configFilePath;
        $this.cfg = $this.LoadConfig($configFilePath);
        $this.ParseConfig();
        $this.SetBackupMode($mode); # Set Default Mode
        $this.SetCompression($compression);
    }

    [string] GetConfigName() {
        return [System.IO.Path]::GetFileNameWithoutExtension($this.ConfigFilePath); 
    }
    [string] GenerateArchiveName() {
        return $this.GenerateArchiveName(($this.ArchiveNameSuffix), $False)
    }
    [string] GenerateArchiveName($suffix) {
        return $this.GenerateArchiveName($suffix, $False)
    }
    [string] GenerateArchiveName($suffix, [boolean] $includeFullPath) {
        $filename = [string]::Format("{0}-{1}.{2}.7z", $this.ConfigName, $this.Timestamp, $suffix);
        if ($includeFullPath) {
            $filename = Join-Path -path $this.BackupTargetDir -ChildPath $filename
        }
        return $filename
    }
    [string] GetArchiveFilename() {
        return [System.IO.Path]::GetFileName($this.ArchiveName); 
    }
    [string] GetWorkDirFullPath($filename) {
        if ([System.IO.Path]::IsPathRooted($filename)) {
            return $filename
        } else {
            return Join-Path -Path $this.BackupTargetDir -ChildPath $filename
        }
    }
    [string] GetArchiveFilePath() {
        return $this.GetWorkDirFullPath($this.ArchiveName)
    }
    [string] GetPasswordParam() {
        # Password can be a plaintext or point to a text file
        $tmp = $null
        $pp = $null
        if ($null -eq $this.BackupKey) {
            # BackupKey is null, so apply default key
            if (-not [string]::IsNullOrEmpty($env:7zBackupKey)) {
                $tmp = $env:7zBackupKey;
            } else {
                $tmp = [string]::Format("{0}:{1}", [System.Net.Dns]::GetHostName(), $this.GetArchiveFilename());
            }
        }
        elseif ($this.backupKey -eq "") {
            # BackupKey set explicitly as empty, so disable encryption
            $tmp = ''
        }
        if ($tmp -eq "") { return "" }

        #TODO: Check the Current directory or Target directory?
        if (Test-Path $tmp -PathType Leaf) {
            $pp = Get-Content $tmp -First 1
        } else {
            $pp = $tmp
        }
        return "-p'$($pp)'"
    }

    [void] SetCompression([compEnum] $compression) { 
        $this.Compresssion = $compression
        if ($compression -eq [compEnum]::None) {
            $this.CmdArgs.Compression = "-m0=Copy"
        }
        elseif ($compression -eq [compEnum]::Fast) {
            $this.CmdArgs.Compression = "-m1=LZMA2"
        }
        else {
            throw 'UNHANDLED COMPRESSION METHOD'
        }
    }

    [void] SetBackupMode([modeEnum] $mode) {
        Write-Debug "SetBackupMode(${mode}) start"
        # Mode
        # Auto: FullNew or Diff - based on condition if file exists or not
        # FullNew: Force to create new archive (especially when current one does exists)
        # Diff: Force diff, fail if baseArchive does not exists
        # FullUpdate - Update existing archive (and not create diff). Fail if BaseArchive does not exists

        # $this.ParseBaseBackupFile(); #just in case
        if ($mode -eq [modeEnum]::FullUpdate)
        {
            if (-not ($this.HasBaseFile))
            {
                throw "The BaseBackupArchivePath does not exists!"
            }

            $this.BackupMode = $mode
            $this.ArchiveNameSuffix = "full"
            $this.ArchiveName = [System.IO.Path]::GetFileName($this.BaseBackupArchivePath)
            $this.CmdArgs.command = "u"
            # $this.CmdArgs.UpdateOptions = "-uq1!$($this.GenerateArchiveName('deleted')) -up0r2r2y2z0w0!'$($this.ArchiveName)'" # CHECK THIS
            #$this.CmdArgs.UpdateOptions = "-up0q0r2y2z0w0" # CHECK THIS
            $this.CmdArgs.UpdateOptions = "-up1q0r2x1y2z1w2" # CHECK THIS
        }
        elseif ($mode -eq [modeEnum]::FullNew)
        {
            $this.BackupMode = $mode
            $this.ArchiveNameSuffix = "full"
            $this.ArchiveName = $this.GenerateArchiveName("full")
            $this.CmdArgs.command = "a"
            $this.CmdArgs.UpdateOptions = ""
        }
        elseif ($mode -eq [modeEnum]::Diff)
        {
            if (($this.BaseBackupArchivePath -ne "") -and (-not (DoesFileExists $this.BaseBackupArchivePath)))
            {
                throw "The BaseBackupArchivePath does not exists!"
            }

            $this.BackupMode = $mode
            $this.ArchiveNameSuffix = "diff"
            $this.ArchiveName = $this.BaseBackupArchivePath
            $this.CmdArgs.command = "u"
            <#
            p0 - ignore not matched files
            q3 - create Delete Item for oryginal file in new archive
            x2 - compress newly created files to the new archive
            z0 - Ignore existing files which are the same as in oryginal archive
            #>
            $this.CmdArgs.UpdateOptions = [string]::Format("-u- -up0q3x2z0!'{0}'", $this.GenerateArchiveName("diff", $True))
        }
        elseif ($mode -eq [modeEnum]::Auto)
        {
            # AUTO
            if ((-not [string]::IsNullOrEmpty($this.BaseBackupArchivePath)) -and (DoesFileExists $this.BaseBackupArchivePath)) {
                $this.SetBackupMode([modeEnum]::Diff)
            } else {
                $this.SetBackupMode([modeEnum]::FullNew)
            }
        }
        else 
        {
            Write-Error "Unhandled BackupMode"
        }
        Write-Debug "SetBackupMode(${mode}) end"
    }

    hidden [void] ParseConfig() {
        Write-Debug "ParseConfig() start"
        $this.BackupTargetDir       = $this.cfg['global']['BackupTargetDir']
        $this.ConfigName            = $this.GetConfigName();
        $this.BaseBackupFilePath    = [string]::Format("{0}.{1}", $this.configName, "bakbase");
        $this.ParseBaseBackupFile();
        $this.BackupListFilePath    = New-TemporaryFile
        $this.ExcludeListFilePath   = New-TemporaryFile
        Save-ListAsFile $this.cfg['backupdirs'].values $this.BackupListFilePath
        Save-ListAsFile $this.cfg['exclude'].values $this.ExcludeListFilePath
        $this.CmdArgs.ToBackup      = [string]::Format("-i@'{0}'", $this.BackupListFilePath)
        $this.CmdArgs.ToExclude     = [string]::Format("-xr@'{0}'", $this.ExcludeListFilePath)
        Write-Debug "ParseConfig() end"
    }

    hidden [void] ParseBaseBackupFile() {
        if (Test-Path -Path $this.BaseBackupFilePath -PathType Leaf) {
            $checkPath = Get-Content $this.BaseBackupFilePath -First 1
            Write-Verbose "BakBase file found and indicating FULL backup archive path: $checkPath"
            $this.BaseBackupArchivePath = $checkPath;
            if (Test-Path -Path $checkPath -PathType Leaf) {
                $this.HasBaseFile = $true
            } else {
                $this.HasBaseFile = $false
                Write-Verbose "The BakBase file points to non-existing archive file"
            }
        }
    }

    hidden [Hashtable] LoadConfig($FilePath) {
        if ((DoesFileExists $FilePath) -eq $False) {
            Write-Error "Config file: ${FilePath} could not be loaded. Check if file exists and try again."
            return $null
        }
    
        $ini = [Hashtable]@{}
        $section = $null
        $index = 0
        switch -regex -file $FilePath
        {
            "^([\#;].*)$" # Comment - ignore line
            {
                continue;
            }
    
            "^\[(.+)\]" # Section
            {
                $section = $matches[1]
                $ini[$section] = @{}
                continue;
            }
    
            "(.+?)\s*=(.*)" # Key
            {
                $name,$value = $matches[1..2]
                $ini[$section][$name] = $value
                continue;
            }
    
             "(.*)" # value only
            {
                $index++
                $value = $matches[1]
                $ini[$section][$index] = $value
                continue;
            }
        }
        return $ini
    }
}

# Build 7z command and run it
function Run-Backup {
    param(
    [Parameter(Mandatory)]
    [string]$ConfigFile,

    [Parameter(HelpMessage="Backup Command/Method")]
    [modeEnum]$Mode = [modeEnum]::Auto,

    [Parameter(HelpMessage="Compression Type")]
    [compEnum]$Compression = [compEnum]::None

    # [Parameter(Mandatory=$False, HelpMessage='Encryption password/key or path to file containing the key. Fallback to ENV variable')]
    # [string]$password, 

    # [Parameter(Mandatory=$False, HelpMessage='Force no password')]
    # [boolean]$forceNoPassword = $False
)

    $cfg = [BackupConfig]::new($ConfigFile, $Mode, $Compression)
    $cfg
    $cfg.CmdArgs

    Write-Verbose "Config: $($cfg.ConfigName)"
    Write-Verbose "ArchiveFile: $($cfg.GetArchiveFilename())"
    # Write-Verbose "Password: $($cfg.GetPasswordParam())"

    $backupTargetDir = $cfg.BackupTargetDir
    $targetBackupArchivePath = $cfg.GetArchiveFilePath()

    $7zArgs = New-Object Collections.Generic.List[string]

    $a = $cfg.CmdArgs
    $7zArgs.add($a.Command);                            # Add or Update
    $7zArgs.add((QuoteStr $targetBackupArchivePath));   # Archive File 
    $7zArgs.add($a.UpdateOptions);                      # Update/Differential options 
    $7zArgs.add($a.ToBackup);                           # What to backup (multiple directories supported)
    $7zArgs.add($a.ToExclude);                          # Exclude files from backup
    $7zArgs.add($a.Compression);                        # Compression Method - https://sevenzip.osdn.jp/chm/cmdline/switches/method.htm
    $7zArgs.add("-t7z");                                # Set 7z format
    $7zArgs.add("-bsp2");                               # Output Stream to StdOut -  https://sevenzip.osdn.jp/chm/cmdline/switches/bs.htm
    $7zArgs.add("-spf2");                               # Save Full Path except RootDrive
    $7zArgs.add("-w'${backupTargetDir}'")               # WorkingDir https://sevenzip.osdn.jp/chm/cmdline/switches/working_dir.htm)
    $7zArgs.add($cfg.GetPasswordParam())                # Encryption Password
    
    $7z = "7z.exe " + ($7zArgs -join " ") + " > 7z.log"
    
    Write-Host "Command: $7z"
    Invoke-Expression $7z
    #CreateConfig $backupFilePath
    #ShowNotification "Full Backup of `"$pathToBackup`" finished!"

    Write-Host "Command: $7z"
    Write-Host "Backup file: $targetBackupArchivePath"

    if (($cfg.BackupMode -eq [modeEnum]::FullNew) -and (DoesFileExists $targetBackupArchivePath))  {
        Save-ListAsFile $targetBackupArchivePath $cfg.baseBackupFilePath
    }
    
}

Run-Backup -configFile $configFile -Mode $mode -Compression $compression