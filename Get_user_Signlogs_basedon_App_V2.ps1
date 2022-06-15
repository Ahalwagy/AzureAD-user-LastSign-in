$aadmodule = get-module | select Name | Where-Object {$_.Name -match "AzureAD"}

if($aadmodule -eq $null)

{Import-Module AzureADPreview}

if($session -eq $null)
{$session= Connect-AzureAD -TenantDomain "Tenantid"}


#Connect to SharePoint Teams in case of uploading the report to Teams or Sharepoint

$SharepointURL = "SharePoint URl"

Connect-PnPOnline $SharepointURL -UseWebLogin

#Graph Login

$ApplicationID = "Application Appid"
$TenatDomainName = "TenantId"
$AccessSecret = "Secretkey"


$Body = @{    
Grant_Type    = "client_credentials"
Scope         = "https://graph.microsoft.com/.default"
client_Id     = $ApplicationID
Client_Secret = $AccessSecret
} 

$ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenatDomainName/oauth2/v2.0/token" `
-Method POST -Body $Body

$token = $ConnectGraph.access_token

$inputfile= "Path"

$invitations = import-csv $inputfile -Encoding Default

#Create outputfile

$date = Get-Date -Format yyyy_dd_MM_hh_mm_tt

$filename = "Teams_Signin_"+ $date + ".csv"

$Path = "Path\$filename"

$StartDate = (Get-Date).AddDays(-30).GetDateTimeFormats()[114]

$Date = (Get-Date).AddDays(-30).ToString('yyyy-MM-dd')


foreach($item in $invitations)

{
        $kid = $item.KID

        $account = $kid + "@eon.com"

        $obj = $account.ToLower()

      try {  
      
            $user = Get-AzureADUser -ObjectId $account


            if($user -ne $null)

            {

                    $objectid = $user.ObjectId

                    $results = Get-AzureADAuditSignInLogs -Filter "UserId eq '$objectid' and startswith(appDisplayName,'Microsoft Teams') and status/errorCode eq 0 and createdDateTime gt $Date" -Top 1 | Select CreatedDateTime ,UserDisplayName , UserPrincipalName ,AppDisplayName

                    Start-Sleep -s 2


                    if($results -ne $null)

                        {
                            $record = $results[0]

                                     $csvValue = New-Object psobject -Property @{
                                                                                  CreatedDateTime  = $record.CreatedDateTime
                                                                                   UserDisplayName = $record.UserDisplayName
                                                                                   KID = $kid
                                                                                   UserPrincipalName = $record.UserPrincipalName
                                                                                   AppDisplayName = $record.AppDisplayName
                                                                                    Comments = "Interactive"
                                                                                             }

                                    $csvValue | Select CreatedDateTime, UserDisplayName, KID, UserPrincipalName, AppDisplayName, Comments |Export-Csv $Path  -NoTypeInformation -Append -Encoding Default

                                    $output= $csvValue | Select CreatedDateTime, UserDisplayName, KID, UserPrincipalName, AppDisplayName, Comments

                                    Write-Host $output -ForegroundColor DarkGreen
            
                        }

                    else

                        {  
                           try{ 
                            
                                $LoginUrl = "https://graph.microsoft.com/beta/auditLogs/signIns/?filter=userPrincipalName eq '$obj' and startswith(appDisplayName,'Microsoft Teams') and signInEventTypes/any(t: t eq 'nonInteractiveUser') and createdDateTime ge $Date &`$top=1"

                                $non = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $LoginUrl -Method Get).value[0]

                                Start-Sleep -s 3

                              }

                            catch

                            {
                                Write-Host "Token Expired-->" $_.Exception.Message -ForegroundColor White

                                
                                    
                                    $Body = @{    
                                    Grant_Type    = "client_credentials"
                                    Scope         = "https://graph.microsoft.com/.default"
                                    client_Id     = $ApplicationID
                                    Client_Secret = $AccessSecret
                                    } 

                                    $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenatDomainName/oauth2/v2.0/token" `
                                    -Method POST -Body $Body

                                    $token = $ConnectGraph.access_token

                                    $LoginUrl = "https://graph.microsoft.com/beta/auditLogs/signIns/?filter=userPrincipalName eq '$obj' and startswith(appDisplayName,'Microsoft Teams') and signInEventTypes/any(t: t eq 'nonInteractiveUser') and createdDateTime ge $Date &`$top=1"

                                    $non = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $LoginUrl -Method Get).value[0]
                                
                                    
                                    Start-Sleep -s 3
                                
                                
                            
                            }

                            if($non -ne $null)

                            {
                                $csvValue = New-Object psobject -Property @{
                                                                                  CreatedDateTime  = $non.createdDateTime
                                                                                   UserDisplayName = $non.userDisplayName
                                                                                   KID = $kid
                                                                                   UserPrincipalName = $non.userPrincipalName
                                                                                   AppDisplayName = $non.appDisplayName
                                                                                    Comments = "Non-Interactive"
                                                                                             }

                                    $csvValue | Select CreatedDateTime, UserDisplayName, KID, UserPrincipalName, AppDisplayName, Comments |Export-Csv $Path  -NoTypeInformation -Append -Encoding Default

                                    $output= $csvValue | Select CreatedDateTime, UserDisplayName, KID, UserPrincipalName, AppDisplayName, Comments

                                    Write-Host $output -ForegroundColor DarkGreen
                               
                            }

                            else

                            {
                            
                                     $csvValue = New-Object psobject -Property @{
                                                                                  CreatedDateTime  = ""
                                                                                   UserDisplayName = $user.DisplayName
                                                                                   KID = $kid
                                                                                   UserPrincipalName = $user.UserPrincipalName
                                                                                   AppDisplayName = ""
                                                                                    Comments = "No teams successfull sign in activity since 30 days"
                                                                                             }

                                    $csvValue | Select CreatedDateTime, UserDisplayName, KID, UserPrincipalName, AppDisplayName, Comments |Export-Csv $Path  -NoTypeInformation -Append -Encoding Default

                                    $output= $csvValue | Select CreatedDateTime, UserDisplayName, KID , UserPrincipalName, AppDisplayName, Comments

                                    Write-Host $output -ForegroundColor Yellow
                            
                            
                            }
                
                        }

                }


            else

            {
                 $csvValue = New-Object psobject -Property @{
                                                                                  CreatedDateTime  = ""
                                                                                   UserDisplayName = ""
                                                                                   KID= $kid
                                                                                   UserPrincipalName = ""
                                                                                   AppDisplayName = ""
                                                                                    Comments = "No Azure AD Account exists"
                                                                                             }

                                    $csvValue | Select CreatedDateTime, UserDisplayName, KID, UserPrincipalName, AppDisplayName, Comments |Export-Csv $Path  -NoTypeInformation -Append -Encoding Default

                                    $output= $csvValue | Select CreatedDateTime, UserDisplayName, kID,  UserPrincipalName, AppDisplayName, Comments

                                    Write-Host $output -ForegroundColor Red
        
        
        
            }

        }

    catch

      {
         $csvValue = New-Object psobject -Property @{
                                                                                  CreatedDateTime  = ""
                                                                                   UserDisplayName = ""
                                                                                   KID= $kid
                                                                                   UserPrincipalName = ""
                                                                                   AppDisplayName = ""
                                                                                    Comments = "No Azure AD Account exists"
                                                                                             }

                                    $csvValue | Select CreatedDateTime, UserDisplayName, KID, UserPrincipalName, AppDisplayName, Comments |Export-Csv $Path  -NoTypeInformation -Append -Encoding Default

                                    $output= $csvValue | Select CreatedDateTime, UserDisplayName, kID,  UserPrincipalName, AppDisplayName, Comments

                                    Write-Host $output -ForegroundColor Red

                                    Write-Host "Account Exception" $_.Exception.Message -ForegroundColor Yellow
    
      }

  }


try
{

  $excel = New-Object -ComObject Excel.Application 
  $excel.Visible = $false

    # change thread culture

    [System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'

    $excel.Workbooks.Open("$Path").SaveAs("$($Path.Replace('.csv','')).xlsx",51)
    $excel.Quit()

    $workbook= $Path.Replace(".csv",".xlsx")

  if (Test-Path $Path) { Remove-Item $Path}

     #$SharepointURL = "https://eonos.sharepoint.com/sites/RegITTeamsSecurityGroups-Mapping"

    #Connect-PnPOnline $SharepointURL -UseWebLogin

    try
    {
       Add-PnPFile -Folder "Shared Documents/General/Enviam/Sign-in Activity Output" -Path $workbook
    
    }

    catch

    {
        Write-Host "Excel File Upload Error-->" $_.Exception.Message -ForegroundColor Magenta
    }

  }

catch

{
    Write-Host "Excel File Creation Error-->" $_.Exception.Message -ForegroundColor Yellow
  }

    #Upload to Teams



#Get-AzureADAuditSignInLogs -Filter "UserId eq '$objectid' and appDisplayName eq 'Microsoft Teams' and status/errorCode eq 0 " | Where-Object {$_.createdDateTime -gt "$StartDate"}| Select CreatedDateTime ,UserDisplayName , UserPrincipalName ,AppDisplayName
