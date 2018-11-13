﻿$csvPath = "c:\temp\users.csv" #CSV File that contains a single column of user email addresses. Heading = Email
$tempFolder = "c:\temp\photos\" #temporarily location to store image for resizing and re-uploading.
$orgName = "contoso" #org prefix; e.g., contoso

$adminURL = "https://" + $orgName  + "-admin.sharepoint.com"
$mySiteURL = "https://" + $orgName + "-my.sharepoint.com"

#Install O365 Client Side Libraries...
Add-Type -Path "C:\Program Files\SharePoint Client Components\16.0\Assemblies\Microsoft.Online.SharePoint.Client.Tenant.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll"

$UserCredential = Get-Credential

# SharePoint Admin Session
$ctxAdmin = New-Object Microsoft.SharePoint.Client.ClientContext($adminURL)
$ctxAdmin.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserCredential.UserName, $UserCredential.Password)
$peopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($ctxAdmin)

# SharePoint MySite Session
$ctxMySite = New-Object Microsoft.SharePoint.Client.ClientContext($mySiteURL)
$ctxMySite.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserCredential.UserName, $UserCredential.Password)

# Import CSV and Report Users
$csvData = Import-Csv $csvPath
$itemCount = ($csvData | Measure-Object).Count
if ($itemCount -gt 0) {

    $library = $ctxMySite.Web.Lists.GetByTitle("User Photos")
    $folder = $library.RootFolder.Folders.GetByUrl("Profile Pictures")
    $ctxMySite.Load($folder)
    $ctxMySite.ExecuteQuery()

    $pictureSizes = @{"_SThumb" = "48"; "_MThumb" = "72"; "_LThumb" = "200"}

    $csvData | ForEach-Object {
        $userEmail = $_.Email
        Try {

            $tempFilePath = $tempFolder + $userEmail + ".jpg"
            $imagePrefix = $userEmail.Replace("@", "_").Replace(".", "_");

            $Param =@{
                Uri = "https://outlook.office365.com/ews/Exchange.asmx/s/GetUserPhoto?email=" + $userEmail + "&size=HR240x240"
                Credential = $UserCredential
                OutFile = $tempFilePath
            }
            Invoke-WebRequest @Param

            $file = Get-ChildItem  -LiteralPath $tempFilePath
            $stream = $file.OpenRead()
            $img = [System.Drawing.Image]::FromStream($stream)

            Foreach($size in $pictureSizes.GetEnumerator())
            {
                [int32]$new_width = $size.Value
                [int32]$new_height = $size.Value
                $img2 = New-Object System.Drawing.Bitmap($new_width, $new_height)
                $graph = [System.Drawing.Graphics]::FromImage($img2)
                $graph.DrawImage($img, 0, 0, $new_width, $new_height)

                #Covert image into memory stream
                $stream = New-Object -TypeName System.IO.MemoryStream
                $format = [System.Drawing.Imaging.ImageFormat]::Jpeg
                $img2.Save($stream, $format)
                $streamseek = $stream.Seek(0, [System.IO.SeekOrigin]::Begin)

                #Upload image into sharepoint online
                $fileName = $imagePrefix + $size.Name + ".jpg"
                
                $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                $FileCreationInfo.Overwrite = $true
                $FileCreationInfo.ContentStream = $stream
                $FileCreationInfo.URL = $fileName
                $fileUpload = $folder.Files.Add($FileCreationInfo)
                $ctxMySite.Load($fileUpload)
                $ctxMySite.ExecuteQuery()
            }

            $stream.Close()

            #Set PictureURL property in SP User Profile
            $pictureURL = $mySiteURL + "/User Photos/Profile Pictures/" + $imagePrefix + "_MThumb.jpg"
            $userAccount = "i:0#.f|membership|" + $_.Email
            $peopleManager.SetSingleValueProfileProperty($userAccount,"PictureURL",$pictureURL)
            $ctxAdmin.ExecuteQuery()

        }
        Catch {
            Write-Output "$('Error |', $userEmail, $_)"
        }
    }
}
