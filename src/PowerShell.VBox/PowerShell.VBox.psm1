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
        $this.CPUCount = $object.CPUCount
        $this.CPUHotPlugEnabled = $object.CPUHotPlugEnabled
        $this.CPUExecutionCap = $object.CPUExecutionCap
        $this.MemorySize = $object.MemorySize
        $this.SnapshotFolder = $object.SnapshotFolder
    }

    [guid]$Id
    [string]$Name
    [int]$CPUCount
    [bool]$CPUHotPlugEnabled
    [ValidateRange(0, 100)]
    [int]$CPUExecutionCap
    [int]$MemorySize
    [System.IO.DirectoryInfo]$SnapshotFolder
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
#endregion

$vbox = Get-VirtualBox