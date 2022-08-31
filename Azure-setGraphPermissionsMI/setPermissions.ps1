<#
.SYNOPSIS
    This script covers the situation when you want to give Graph permissions to a managed identity.
.DESCRIPTION
    Script need inputs for the identity name (for the managed identity), resource group name (of the managed identity) and all the role assignments, that you want to give
    to your identity, as an array.
.EXAMPLE
    RUN: ./setPermission.ps1
.NOTES
    REMEMBER to put in the identity name ($identityName), resource group name ($rgName) and the permissions you want to delegate ($permissions)
#>

$identityName = "id-templatetool-d-01"
$rgName = "rg-templatetool-development"
$permissions = @("ChannelMember.Read.All")

az login

# Use this if you have a system-assigned managed identity (e.g. in a web app or a logic app) 
# $spId = (az resource list -n $identityName --query [*].identity.principalId --out tsv)

# Use this if you have a user-assigned managed identity (to cover multiple resources)
$spId = (az identity show -n $identityName -g $rgName --query principalId --out tsv)

# Get's the resourceId for Microsoft graph
$graphResourceId = $(az ad sp list --display-name "Microsoft Graph" --query [0].objectId --out tsv)

#The url for graph permissions for the defined service principal (managed identity)
$uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$spId/appRoleAssignments"

Write-Host $uri

$appRoleIds = [System.Collections.ArrayList]@()

foreach ($role in $permissions) {
    try {
        #Get's the role id from Microsoft Graph permissions for the permission you want to delegate
        $roleId = $(az ad sp list --display-name "Microsoft Graph" --query "[0].appRoles[?value=='$role' && contains(allowedMemberTypes, 'Application')].id" --output tsv)
        $appRoleIds.Add($roleId)
        Write-Host "Added role id successfully for role: $role, with id: $roleId"
    }
    catch {
        Write-Host "An error ocurred when getting roleId"
        Write-Host $_
    } 
}

foreach ($roleId in $appRoleIds) {
    try {
        #Create the body for the az request to add the permissions for the identity
        $body = "{'principalId':'$spId','resourceId':'$graphResourceId','appRoleId':'$roleId'}"
        Write-Host $body

        # Taking advantage of Azure cli rest so no need to get access tokens Posts to the Graph api to add defined role to your identity's permissions
        az rest --method post --uri $uri --body $body --headers "Content-Type=application/json"
    }
    catch {
        Write-Host 'An error ocurred when adding a permission'
        Write-Host $_
    }
    finally {
        Start-Sleep -s 2
    }

}

Write-Host "The script has ended"
