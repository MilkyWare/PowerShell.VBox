#Requires -Version 5.0

#region Classes
class VBoxMachine
{
    VBoxMachine()
    { }

    VBoxMachine ([__ComObject]$object)
    { 
        $this.Id = $object.Id
        $this.Name = $object.Name
        $this.Description = $object.Description
        $this.CPUCount = $object.CPUCount
        $this.CPUHotPlugEnabled = $object.CPUHotPlugEnabled
        $this.CPUExecutionCap = $object.CPUExecutionCap
        $this.MemorySize = $object.MemorySize
        $this.State = [VBoxMachineState]$object.State
        $this.LogFolder = $object.LogFolder
        $this.ClipboardMode = $object.ClipboardMode
        $this.FirmwareType = $object.FirmwareType
        $this.ChipsetType = $object.ChipsetType
        $this.SnapshotFolder = $object.SnapshotFolder
        $this.SettingsFilePath = $object.SettingsFilePath
        $this.DragNDropMode = $object.DnDMode
    }

    [guid]$Id
    [string]$Name
    [string]$Description
    [int]$CPUCount
    [bool]$CPUHotPlugEnabled
    [ValidateRange(0, 100)]
    [int]$CPUExecutionCap
    [int]$MemorySize
    [VBoxMachineState]$State
    [System.IO.DirectoryInfo]$LogFolder
    [VBoxClipboardMode]$ClipboardMode
    [VBoxFirmwareType]$FirmwareType
    [VBoxChipsetType]$ChipsetType
    [System.IO.DirectoryInfo]$SnapshotFolder
    [System.IO.FileInfo]$SettingsFilePath
    [VBoxDragNDropMode]$DragNDropMode
}

enum VBoxChipsetType
{
    PIIX3 = 1
    ICH9 = 2
}

enum VBoxClipboardMode
{
    Disabled = 0
    HostToGuest = 1
    GuestToHost = 2
    Bidirectional = 3
}

enum VBoxDragNDropMode
{
    Disabled = 0
    HostToGuest = 1
    GuestToHost = 2
    Bidirectional = 3
}

enum VBoxFirmwareType
{
    BIOS = 0
    EFI = 1
    EFI32 = 2
    EFI63 = 3
    EFIDual = 4
}

enum VBoxLockType
{
    Shared = 1
    Write = 2
    VM = 3
}

enum VBoxMachineState
{
    Stopped = 1
    Saved = 2
    Teleported = 3
    Aborted = 4
    Running = 5
    Paused = 6
    Stuck = 7
    Snapshotting = 8
    Starting = 9
    Stopping = 10
    Restoring = 11
    TeleportingPausedVM = 12
    TeleportingIn = 13
    FaultTolerantSync = 14
    DeletingSnapshotOnline = 15
    DeletingSnapshot = 16
    SettingUp = 17
}
#endregion

#region Functions
function Get-VirtualBox
{
    [CmdletBinding()]
    param ()
    process
    {
        Write-Verbose "Creating VirtualBox COM object"
        $vbox = New-Object -ComObject VirtualBox.VirtualBox
        return $vbox
    }
}

function Get-Machines
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string[]]$Name
    )
    process
    {
        $machines = [System.Collections.Generic.List[VBoxMachine]]::new()
        if ($Name)
        {
            Write-Verbose "Getting virtual machines by name"
            foreach ($n in $Name)
            {
                $m = $vbox.FindMachine($n);
                $machines.Add([VBoxMachine]::new($m))
            }
        }
        else
        {
            Write-Verbose "Getting all virtual machines"
            foreach ($m in $vbox.Machines)
            {
                $machines.Add([VBoxMachine]::new($m))
            }
        }
        return $machines
    }
}

function Suspend-Machine
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [guid[]]$Id
    )
    process
    {
        foreach ($i in $Id)
        {
            Write-Verbose "Finding machine $i"
            $machine = $vbox.FindMachine($i)

            if ($machine)
            {
                Write-Verbose "Found {$machine.Name}"
                if ($PSCmdlet.ShouldProcess($machine.Name))
                {
                    Write-Verbose "Creating VBox session"
                    $session = New-Object -ComObject "VirtualBox.Session"

                    Write-Verbose "Locking machine"
                    $machine.LockMachine($session, [VBoxLockType]::Shared)

                    Write-Verbose "Saving state"
                    $session.Machine.SaveState() | Out-Null

                    Write-Verbose "Unlocking machine"
                    $session.UnlockMachine()
                }
                else
                {
                }
            }
        }
    }
}
#endregion

$vbox = Get-VirtualBox