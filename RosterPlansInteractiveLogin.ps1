# Scripts to try out the rosterPlan APIs

# MSAL.PS added to the function to support the MSAL libraries
# Available from https://github.com/AzureAD/MSAL.PS or https://www.powershellgallery.com/packages/MSAL.PS
# Or Install-Module MSAL.PS -AcceptLicense
Import-Module MSAL.PS

# Interactive login
# Client ID is created in Azure AD under App Registration - requires Group.Read.All and the default User.Read
# Redirect Url is Mobile and Desktop applications - https://login.microsoftonline.com/common/oauth2/nativeclient
# Change TenantId to your own tenant 

# $graphToken = Get-MsalToken -ClientId "47d72fc3-8478-42d0-b5fd-be49463eddb7" -TenantId "01ba1a71-c58f-48a6-bc02-5e697e4298e5" `
# -Interactive -Scope 'https://graph.microsoft.com/Group.Read.All', 'https://graph.microsoft.com/User.Read' `
# -LoginHint brismith@brismith.onmicrosoft.com

$graphToken = Get-MsalToken -ClientId "53ec3fb2-5051-415f-b84d-abc1c04abbe2" -TenantId "01ba1a71-c58f-48a6-bc02-5e697e4298e5" `
-Interactive -Scope 'https://graph.microsoft.com/Tasks.ReadWrite', 'https://graph.microsoft.com/User.Read' `
-LoginHint brismith@brismith.onmicrosoft.com

#################################################
# Create Roster
#################################################

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('Content-Type', "application/json")

$setRequest =@{}
$setRequest.Add("@odata.type", "#microsoft.graph.plannerRoster")

$request = @" 
$($setRequest | ConvertTo-Json)
"@

$uri = "https://graph.microsoft.com/beta/planner/rosters"

# Create roster container

$rosterRequest = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $request

$rosterId=$rosterRequest.id

#################################################
# Create Plan with the Roster as container
#################################################

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('Content-Type', "application/json")


$container=@{}
$container.Add("url", "https://graph.microsoft.com/beta/planner/rosters/" + $rosterId)
$setRequest =@{}
$setRequest.Add("container",$container)
$setRequest.Add("title", "Yes another plan in a roster container " + $rosterId)

$request = @" 
$($setRequest | ConvertTo-Json)
"@

$uri = "https://graph.microsoft.com/beta/planner/plans"

# Create plan in roster container

$rosterRequest = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $request

#################################################
# Populate Roster with accounts
# userId will take either UPN or object Id
#################################################

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('Content-Type', "application/json")

$setRequest =@{}
$setRequest.Add("@odata.type", "#microsoft.graph.plannerRosterMember")
$setRequest.Add("userId", "ceaa49d3-5dd7-425d-b247-404ccc98e6f0")



$request = @" 
$($setRequest | ConvertTo-Json)
"@

$uri = "https://graph.microsoft.com/beta/planner/rosters/" + $rosterId + "/members"

# Create roster container

$rosterRequest = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $request

#################################################
# Read members
#################################################

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('Content-Type', "application/json")

$uri = "https://graph.microsoft.com/beta/planner/rosters/" + $rosterId + "/members"

# Read members

$rosterRequest = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers

$members = $rosterRequest.value
$members.Count
$members[0].userId

#################################################
# Read plan in roster container
#################################################

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('Content-Type', "application/json")

$uri = "https://graph.microsoft.com/beta/planner/rosters/" + $rosterId + "/plans"

# Read plan (there can only be one plan per roster container I think)

$rosterRequest = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
$plans = $rosterRequest.value
$plans.Count
$plans[0]

#################################################
# Read users plans in roster containers
#################################################

$userId = "52ab5038-1fc7-48bc-a092-9efb7577450c"

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('Content-Type', "application/json")

$uri = "https://graph.microsoft.com/beta/users/" + $userId + "/planner/rosterPlans"

# Read members

$rosterRequest = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
$plans = $rosterRequest.value
$plans.Count
$plans[0]



#################################################
# Create plan in a Group container
# This just needs Tasks.ReadWrite if the Group exists (I think)
################################################# 

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('Content-Type', "application/json")


$container=@{}
$container.Add("url", "https://graph.microsoft.com/beta/groups/452a9428-7dbe-4511-a211-13fffaa8bf7c")
$setRequest =@{}
$setRequest.Add("container",$container)
$setRequest.Add("title", "Yet another Plan in a group container")

$request = @" 
$($setRequest | ConvertTo-Json)
"@

$uri = "https://graph.microsoft.com/beta/planner/plans"

# Create plan in Group container

$rosterRequest = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $request
