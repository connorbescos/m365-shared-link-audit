#Parameters
$SiteUrl = "https://org-my.sharepoint.com/personal/connorb_company_com"
$ReportOutput = ".\SharedLinks.csv"
$ListName = "documents"
    
#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -Interactive
$Ctx = Get-PnPContext
$Results = @()
$global:counter = 0
  
#Get all list items in batches
$ListItems = Get-PnPListItem -List $ListName -PageSize 2000
$ItemCount = $ListItems.Count
   
#Iterate through each list item
ForEach($Item in $ListItems)
{
    Write-Progress -PercentComplete ($global:Counter / ($ItemCount) * 100) -Activity "Getting Shared Links from '$($Item.FieldValues["FileRef"])'" -Status "Processing Items $global:Counter to $($ItemCount)";
 
    #Check if the Item has unique permissions
    $HasUniquePermissions = Get-PnPProperty -ClientObject $Item -Property "HasUniqueRoleAssignments"
    If($HasUniquePermissions)
    {       
        #Get Shared Links
        $SharingInfo = [Microsoft.SharePoint.Client.ObjectSharingInformation]::GetObjectSharingInformation($Ctx, $Item, $false, $false, $false, $true, $true, $true, $true)
        $ctx.Load($SharingInfo)
        $ctx.ExecuteQuery()
         
        ForEach($ShareLink in $SharingInfo.SharingLinks)
        {
            If($ShareLink.Url)
            {           
                If($ShareLink.IsEditLink)
                {
                    $AccessType="Edit"
                }
                ElseIf($shareLink.IsReviewLink)
                {
                    $AccessType="Review"
                }
                Else
                {
                    $AccessType="ViewOnly"
                }
                #Get sharedlink recepients
                try{
                    $link = get-pnpfilesharinglink -fileurl $Item.FieldValues["FileRef"]
                    $SharedWith = $link.grantedToIdentitiesV2.user.email -join ", "
                }
                catch{
                    $SharedWith = $null
                }

                #Collect the data
                $Results += New-Object PSObject -property $([ordered]@{
                Name  = $Item.FieldValues["FileLeafRef"]           
                RelativeURL = $Item.FieldValues["FileRef"]
                FileType = $Item.FieldValues["File_x0020_Type"]
                ShareLink  = $ShareLink.Url
                ShareLinkAccess  =  $AccessType
                ShareLinkType  = $ShareLink.LinkKind
                AllowsAnonymousAccess  = $ShareLink.AllowsAnonymousAccess
                IsActive  = $ShareLink.IsActive
                Expiration = $ShareLink.Expiration
                ExternalGuestInvitees = $Sharelink.HasExternalGuestInvitees
                Recpients = $SharedWith
                Created = $ShareLink.Created
                })
            }
        }
    }
    $global:counter++
}
$Results | Export-CSV $ReportOutput -NoTypeInformation
Write-host -f Green "Sharing Links Report Generated Successfully!"



