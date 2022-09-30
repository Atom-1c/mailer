param([string]$infoPath = (get-location).Path, 
        [string] $fileName = "CustomerInfo.xlsx")

#constructs $customerInfo from the excel spreadsheet
function Import-CustomerInfo($filepath) {
    import-module psexcel #it wasn't auto loading on my machine
    return (Import-XLSX -Path $filepath)
    }

#constructs an outlook connection
function Create-OutlookConnection {
    return (New-Object -ComObject Outlook.Application)
    }

#constructs a file path to a given product
function Get-Path($productName, $msgPath) {
    if($productName){
        return $msgPath + $productName.trim() + ".msg"
        }
    }

#checks if the file specified as the source file is read to receive IO from this process
function Wait-ForCloseFile($filePath) {
    $openFile = New-Object -TypeName System.IO.FileInfo -ArgumentList $filePath
    $ErrorActionPreference = "SilentlyContinue"
    while ($true) { 
        [System.IO.FileStream] $fs = $openFile.OpenWrite(); 
        if (!$?) {
            read-host "Please save and close the Source File before continuing"
            }
        else {
            $fs.Dispose()
            break
            }
        }
    }

#randomly chooses a product from the approved list of products that has a corresponding mail object in the /mail folder
function Get-MultiProductDecision($customerInfo, $activeMail) {
    $approvedProducts = @()
    foreach($productType in $activeMail.product){
        if($customerInfo.$productType -contains "y"){
            $approvedProducts+=$productType
            }
        }
    if($approvedProducts.Length -gt 0){
        return $approvedProducts[(Get-Random -Maximum $approvedProducts.Length)]
        }
    }

#gets the content from a given .msg file with a specified outlook session
function Get-MessageContent($msgPath, $outlookSession) {
    $item = $outlookSession.Session.OpenSharedItem($msgPath)
    return $item.subject, $item.HTMLbody
    }

#this function imports .msg files from $msgPath with a name that matches a field heading from the customer info excel doc and returns an array of products and .msg html content
function Import-MessageContent($msgPath, $customerInfo){
    $productList = (($customerInfo[0] | Get-Member -MemberType Properties).Name).Trim()
    $emailContent = @()
    $outlook = Create-OutlookConnection
    if(Test-Path -Path $msgPath){
        foreach($product in $productList){
            $contentPath = (Get-Path $product $msgPath)
            if(Test-Path -Path $contentPath){
                $subject, $content = Get-MessageContent $contentPath $outlook
                $emailContent += New-Object -TypeName psobject -Property @{Product=$product; Content=$content; Subject=$subject}
            }
        }
    }
    $Outlook.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)
    return $emailContent
}


#test code block
#$testSubject, $testBody = Get-MessageContent (Get-Path "intro" ((Get-Location).path + "\mail\")) (Create-OutlookConnection)
#Write-Host $testSubject
#Write-Host $testBody
#$testCustomers = Import-CustomerInfo ($infoPath + "\" + $fileName)
#write-host $testCustomers
#$testMessageContent = Import-MessageContent ((get-location).Path + "\mail\") $testCustomers
#write-host $testMessageContent.product $testMessageContent.subject
#$testMailRelay = Create-OutlookConnection
#Write-Host $testMailRelay.Name
#$testProductDecision = Get-MultiProductDecision $testCustomers[0]
#Write-Host $testProductDecision
#end test code block

######-BEGIN ACTUAL EXEC CODE-##########

Wait-ForCloseFile ($infoPath + "\" + $fileName)

$customerInfo = Import-CustomerInfo ($infoPath + "\" + $fileName)

$mailCampaigns = Import-MessageContent ($infoPath + "\mail\") $customerInfo

$delay = Read-Host "Defer the Delivery time or Date? Y/N"

if($delay.ToLower() -match "y"){

Add-Type -AssemblyName System.Windows.Forms

# Main Form
$mainForm = New-Object System.Windows.Forms.Form
$font = New-Object System.Drawing.Font(“Consolas”, 13)
$mainForm.Text = ” Pick Date-Time”
$mainForm.Font = $font
$mainForm.ForeColor = “White”
$mainForm.BackColor = “DarkBlue”
$mainForm.Width = 350
$mainForm.Height = 225

# DatePicker Label
$datePickerLabel = New-Object System.Windows.Forms.Label
$datePickerLabel.Text = “Date”
$datePickerLabel.Location = “15, 10”
$datePickerLabel.Height = 22
$datePickerLabel.Width = 90
$mainForm.Controls.Add($datePickerLabel)

# MinTimePicker Label
$minTimePickerLabel = New-Object System.Windows.Forms.Label
$minTimePickerLabel.Text = “Time”
$minTimePickerLabel.Location = “15, 45”
$minTimePickerLabel.Height = 22
$minTimePickerLabel.Width = 90
$mainForm.Controls.Add($minTimePickerLabel)

# DatePicker
$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = “110, 7”
$datePicker.Width = “150”
$datePicker.Format = [windows.forms.datetimepickerFormat]::custom
$datePicker.CustomFormat = “dd/MM/yyyy”
$mainForm.Controls.Add($datePicker)

# MinTimePicker
$minTimePicker = New-Object System.Windows.Forms.DateTimePicker
$minTimePicker.Location = “110, 42”
$minTimePicker.Width = “150”
$minTimePicker.Format = [windows.forms.datetimepickerFormat]::custom
$minTimePicker.CustomFormat = “HH:mm:ss”
$minTimePicker.ShowUpDown = $TRUE
$mainForm.Controls.Add($minTimePicker)

# OD Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = “15, 130”
$okButton.ForeColor = “Black”
$okButton.BackColor = “White”
$okButton.Text = “OK”
$okButton.add_Click({
    $mainForm.close()
    })
$mainForm.Controls.Add($okButton)

[void] $mainForm.ShowDialog()

$deferredDeliveryTime = ($datePicker.Value.Date.Add($minTimePicker.Value.TimeOfDay))
write-host "Mail will be sent " $deferredDeliveryTime

}

$Preview = Read-Host "Preview messages before sending? (Recommended if sending now) Y/N"

foreach ($customer in $customerInfo){

    $mailRelay = Create-OutlookConnection
    while(!$mailRelay){
        Write-host "waiting for relay"
        }
    $mailType = Get-MultiProductDecision $customer $mailCampaigns
    $bodyContent = $mailCampaigns | Where-Object{$_.product -contains $mailType}
    if(!$bodyContent.content){
        continue
        }

    $ntnxMail = $mailRelay.CreateItem(0)
    $ntnxMail.HTMLBody = "Hey " + $customer.Customer + " Team," + $bodyContent.content
    $ntnxMail.To = $customer.Emails
    $ntnxMail.Subject = $bodyContent.subject

    if ($deferredDeliveryTime) {
        $ntnxMail.DeferredDeliveryTime = $deferredDeliveryTime
        }    

    if ($preview.toLower() -match "y") {
        $ntnxMail.Display()
        read-host “Press ENTER to continue...”
        } 
    else {
        $ntnxMail.Send()
        }

    }