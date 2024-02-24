###++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++###
###--------------------------------------------------------------------###

#Excel is need : https://github.com/dfinke/ImportExcel
#install ImportExcel by running:
# Install-Module ImportExcel -scope CurrentUser

#Get All Clusters
$ClusterName = get-cluster -Name 'ClusterName'

#Stil til Fil
$PathFile = ('C:\temp\DRS - Groups And Rules\' + "$VMCluster" + "_DRS_Rules.xlsx")


#Look through all Clusters
foreach($VMCluster in $ClusterName){
	
	#Get all VMHostGroups and VMGroups
	$ClusterGroup = get-cluster -Name $VMCluster | Get-DrsClusterGroup | Sort-Object GroupType -Descending
	#Generate ClusterGroup Report
	$Report = foreach ($DRSGroup in $ClusterGroup){
    	[pscustomobject]@{
        	GroupName = $DRSGroup
        	Cluster = $DRSGroup.Cluster
        	GroupType = $DRSGroup.GroupType
        	Member = (@($DRSGroup.Member) -join ',')
    	}
	}
	#Print ClusterGroup Report
	$Report | Export-Excel -WorksheetName 'DrsClusterGroups' -Path $PathFile

	#Get all Affinity Rules From Cluster
	$AffinityRules = get-cluster -Name $VMCluster | Get-DrsRule
	#Generate Affinity Report
	$AffinityReport = foreach ($Rule in $AffinityRules){
    	[pscustomobject]@{
    	RuleName = $Rule
    	Cluster = $Rule.Cluster
    	VMName = (@((Get-View -id $Rule.ExtensionData.Vm).name) -join ',')
    	}
	}
	#Append and Print Affinity Report to file
	$AffinityReport | Export-Excel -WorksheetName 'DRSRules' -Append -Path $PathFile

	$DRSHostRule = get-cluster -Name $VMCluster | Get-DrsVMHostRule | Select Name,VMGroup,Type,VMHostGRoup,Enabled | Sort-Object -Property Name
	#Generate DRSHostRule Report
	$DRSHostRuleReport = foreach($DRSHRule in $DRSHostRule){
		[PSCustomObject]@{
			Name 		= $DRSHRule.Name
			VMGroup 	= $DRSHRule.VMGroup
			Type		= $DRSHRule.Type
			VMHostGroup = $DRSHRule.VMHostGroup
			Enabled		= $DRSHRule.Enabled 
			VMGroupMembers = (@($DRSHRule.VMGroup.Member) -join ',')
			VMHostGroupMembers = (@($DRSHRule.VMHostGroup.Member) -join ',')
		}
	}
	#Append and Print Affinity Report to file
	$DRSHostRuleReport | Export-Excel -WorksheetName 'DrsVMHostRule' -Append -Path $PathFile
}


