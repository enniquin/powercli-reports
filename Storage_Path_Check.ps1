# Borrowed / Inspired from: VMware/Storage/Storage_path_Check_v2.ps1
#
# Excel is needed : https://github.com/dfinke/ImportExcel
# Install ImportExcel by running:
# Install-Module ImportExcel -scope CurrentUser

# Get all VMHost
$VMHost = Get-Cluster | Get-VMHost
     

$lunReport = foreach ($VH in $VMHost) {
    
        foreach ($LUNMP in $VH.ExtensionData.config.StorageDevice.MultipathInfo.lun) {
    
            
            $LUN = $VH.ExtensionData.config.StorageDevice.ScsiLun | where uuid -eq $LUNMP.Id

    
            $PathNotActive = $LUNMP.path | where state -ne "Active" | Measure-Object | select -ExpandProperty count
            $NumberOfPaths =  $LUNMP.path | Measure-Object | select -ExpandProperty count
            $LUNID = $LUNMP.Path.name | ForEach-Object {$_.split("L") | select -last 1} | select -Unique
    
            [pscustomobject]@{
                Cluster             = $VH.Parent
                Host                = $VH.name
                NAA                 = $LUN.canonicalname
                OperationalState    = [string]$LUN.OperationalState
                CapacityGB          = [math]::round($LUN.capacity.BlockSize * $LUN.capacity.Block / 1GB,1)
                MultipathingPolicy  = $LUNMP.Policy.policy
                SATP                = $LUNMP.StorageArrayTypePolicy.Policy
                Description         = "$($LUN.Vendor) $($Lun.Model)"
                NonActivePaths      = $PathNotActive
                NumberOfPaths       = $NumberOfPaths
                Type                = $LUN.DeviceType
                Local               = $LUN.LocalDisk
                LunID               = [int]$LUNID
                
            }
    
        }
    
} # Foreach VMHost

# Generate Storage path and multipath configuration
    $lunReport | Export-Excel | Sort-Object Cluster -Descending

# Would like to narrow the script down to onyl show SAN Paths / Fibrechannel / vmhba
