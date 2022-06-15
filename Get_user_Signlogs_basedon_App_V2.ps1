$aadmodule = get-module | select Name | Where-Object {$_.Name -match "AzureAD"}

if($aadmodule -eq $null)

{Import-Module AzureADPreview}

$TenatDomainName = "TenantID"

if($session -eq $null)
{$session= Connect-AzureAD -TenantDomain $TenatDomainName }


#Graph Login

$ApplicationID = "ApplicationID"
$AccessSecret = "Secret Key"


$Body = @{    
Grant_Type    = "client_credentials"
Scope         = "https://graph.microsoft.com/.default"
client_Id     = $ApplicationID
Client_Secret = $AccessSecret
} 

$ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenatDomainName/oauth2/v2.0/token" `
-Method POST -Body $Body

$token = $ConnectGraph.access_token

#$inputfile= "Path"
# Get All the users in the tenant , or optionally you can read from file

$users = Get-AzureADUser -All $true

#Create outputfile

$date = Get-Date -Format yyyy_dd_MM_hh_mm_tt

$filename = "Signin_"+ $date + ".csv"

$Path = "Path\$filename"

foreach($user in $users)

{

 $objectid = $user.ObjectId
  
 
       try
       { 

            $LoginUrl = "https://graph.microsoft.com/beta/users/$objectid/?`$select=userPrincipalName,signInActivity"

            $signin = Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $LoginUrl -Method Get | select signInActivity 


             $csvValue = New-Object psobject -Property @{
                                                                               
                                                                                   UserDisplayName = $user.DisplayName
                                                                                   UserPrincipalName = $user.UserPrincipalName
                                                                                   ObjectId=$user.ObjectId
                                                                                   lastSignInDateTime = $signin.signInActivity.lastSignInDateTime
                                                                                   lastNonInteractiveSignInDateTime=$signin.signInActivity.lastNonInteractiveSignInDateTime
                                                                                             }

                                    $csvValue | Select UserDisplayName,UserPrincipalName,ObjectId,lastSignInDateTime,lastNonInteractiveSignInDateTime |Export-Csv $Path  -NoTypeInformation -Append -Encoding Default

                                    $output= $csvValue | Select UserDisplayName,UserPrincipalName,ObjectId,lastSignInDateTime,lastNonInteractiveSignInDateTime

                                    Write-Host $output -ForegroundColor DarkGreen

             

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

                                     $LoginUrl = "https://graph.microsoft.com/beta/users/$objectid/?`$select=userPrincipalName,signInActivity"

                                     $signin = Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $LoginUrl -Method Get | select signInActivity 

                                     $signin.signInActivity.lastSignInDateTime
                                     $signin.signInActivity.lastNonInteractiveSignInDateTime

                                     $csvValue = New-Object psobject -Property @{
                                                                               
                                                                                   UserDisplayName = $user.DisplayName
                                                                                   UserPrincipalName = $user.UserPrincipalName
                                                                                   ObjectId=$user.ObjectId
                                                                                   lastSignInDateTime = $signin.signInActivity.lastSignInDateTime
                                                                                   lastNonInteractiveSignInDateTime=$signin.signInActivity.lastNonInteractiveSignInDateTime
                                                                                             }

                                    $csvValue | Select UserDisplayName, UserPrincipalName, ObjectId, lastSignInDateTime, lastNonInteractiveSignInDateTime |Export-Csv $Path  -NoTypeInformation -Append -Encoding Default

                                    $output= $csvValue | Select UserDisplayName, UserPrincipalName, ObjectId, lastSignInDateTime, lastNonInteractiveSignInDateTime

                                    Write-Host $output -ForegroundColor DarkGreen
                                
                            
            }

                           

                            
                
}

               


#Create a report in excel and upload to teams

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

   

  }

catch

{
    Write-Host "Excel File Creation Error-->" $_.Exception.Message -ForegroundColor Yellow
  }






