#Print lable and save CSV-file


Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

$Logfile = ".\utlamning-$(gc env:username)-$(gc env:computername).log"

LogWrite "name,samaccountname,pnr,class,school,serialnumber,date"


Add-Type -AssemblyName System.Web          
$name = ""
$school = ""
$class = ""
$pnr = "" 
$samaccountname = ""
 $hash = @{}
 $hash.'CRLF' = '%0D%0A'
 $hash.'+' = '%20'
 $hash.' ' = '%20'

#$printerNameURL = [System.Web.HttpUtility]::UrlEncode($printerName).Replace('+', '%20') 
#Write-Host "This is the Encoded URL" $printerNameURL -ForegroundColor Green
$labelSetXml = ''


# Write-Host $labelSetXmlURL


$labelXml = '<?xml version="1.0" encoding="utf-8"?><DieCutLabel Version="8.0" Units="twips"><PaperOrientation>Landscape</PaperOrientation><Id>Address</Id><PaperName>30252 Address</PaperName><DrawCommands><RoundRectangle X="0" Y="0" Width="2025" Height="5020" Rx="270" Ry="270"/></DrawCommands><ObjectInfo><AddressObject><Name>Address</Name><ForeColor Alpha="255" Red="0" Green="0" Blue="0"/><BackColor Alpha="0" Red="255" Green="255" Blue="255"/><LinkedObjectName/><Rotation>Rotation0</Rotation><IsMirrored>False</IsMirrored><IsVariable>True</IsVariable><HorizontalAlignment>Left</HorizontalAlignment><VerticalAlignment>Middle</VerticalAlignment><TextFitMode>ShrinkToFit</TextFitMode><UseFullFontHeight>True</UseFullFontHeight><Verticalized>False</Verticalized><StyledText/><ShowBarcodeFor9DigitZipOnly>False</ShowBarcodeFor9DigitZipOnly><BarcodePosition>BelowAddress</BarcodePosition><LineFonts/></AddressObject><Bounds X="332" Y="150" Width="4455" Height="1260"/></ObjectInfo></DieCutLabel>'
$labelXmlURL = [System.Web.HttpUtility]::UrlEncode($labelXml)


 Foreach ($key in $hash.Keys) {
    $labelXmlURL = $labelXmlURL.Replace($key, $hash.$key)
 }
 
# Write-Host $labelXmlURL

$printParamsXml =""

# kontrollera om det finns en ansluten skdymo skrivare och hämta namnet

[XML]$xml = Invoke-RestMethod -Method Get -Uri https://localhost:41951/DYMO/DLS/Printing/GetPrinters
#Write-Host $xml.Printers.LabelWriterPrintes
$printerName = $xml.printers.LabelWriterPrinter.Name
Write-Host "Ansluten skrivare: " $printerName -ForegroundColor Green
$printerNameURL = [System.Web.HttpUtility]::UrlEncode($printerName)


 Foreach ($key in $hash.Keys) {
    $printerNameURL = $printerNameURL.Replace($key, $hash.$key)
 }

pause

# skapa grundmall för XML från template (XML-fil eller $string)

# WHILE pnr != x (CTRL+c)
# START
while($pnr -ne "x") {

#clear
write-host ""
write-host 'Fyll i personnummer som YYYYMMDDNNNN'
write-host 'Avsluta med att fylla i "x" som personnummer' -ForegroundColor Green
write-host ""

$PNR = Read-Host 'Personnummer' 
if ($pnr -eq "x") {break}
write-host ""
write-host ""
write-host 'PNR: '$PNR
$strFilter = "(&(objectCategory=User)(employeeid=$PNR))"

$objDomain = New-Object System.DirectoryServices.DirectoryEntry

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1
$objSearcher.Filter = $strFilter
$objSearcher.SearchScope = "Subtree"

$colProplist = "displayname", "department", "physicaldeliveryofficename", "mail", "samaccountname"

foreach ($i in $colPropList){
$dummy = $objSearcher.PropertiesToLoad.Add($i)
}

$colResults = $objSearcher.FindAll()

foreach ($objResult in $colResults)
    {
    $objItem = $objResult.Properties
         $name = $objItem.displayname
         $school = $objItem.department
         $class = $objItem.physicaldeliveryofficename
         $mail = $objItem.mail
         $samaccountname = $objItem.samaccountname
         }
write-host "Namn:  $name"
write-host "Skola: $school"
write-host "Klass: $class"
write-host "Mail:  $mail"
write-host "ID:  $samaccountname"

write-host ""
write-host ""

pause
write-host ""
write-host 'Fyll i serienummer'
write-host 'Avsluta med att fylla i "x" som personnummer' -ForegroundColor Green
write-host ""
$serialnumber = ""
while ($serialnumber.Length -ne 12){
$serialnumber = Read-Host 'Serienummer' 
if ($serialnumber -eq "x") {break}
write-host ""
write-host ""

$serialnumber = $serialnumber.ToUpper().TrimStart("S")
write-host 'Serienummer: '$serialnumber
}
$date = get-date -format u
LogWrite "$name,$samaccountname,$pnr,$class,$school,$serialnumber,$date"

# END

# fånga inmatning av serienummer (från sträckkod)

# skapa xml för output
#$text = 'Any old text string'
$date = get-date -format u 
# $labelSetXml = '<LabelSet><LabelRecord><ObjectData Name="Address">'+$date+'CRLF'+$pnr+'CRLF'+$name+'CRLF'+$school+'CRLF'+$class+'CRLF'+$serialnumber+'</ObjectData></LabelRecord></LabelSet>'
 $labelSetXml = '<LabelSet><LabelRecord><ObjectData Name="Address">'+$name+'CRLF'+$school+'CRLF'+$class+'CRLF'+$serialnumber+'</ObjectData></LabelRecord></LabelSet>'

# barcodetest
# $labelSetXml = '<LabelSet><LabelRecord><ObjectData Name="Address">'+$name+'CRLF'+$school+'CRLF'+$class+'</ObjectData><ObjectData Name="Barcode">'+$pnr+'</ObjectData></LabelRecord></LabelSet>'

#$labelSetXmlURL = [System.Web.HttpUtility]::UrlEncode($labelSetXml).Replace('+', '%20').Replace('CRLF', '%0D%0A') 
$labelSetXmlURL = [System.Web.HttpUtility]::UrlEncode($labelSetXml)


 Foreach ($key in $hash.Keys) {
    $labelSetXmlURL = $labelSetXmlURL.Replace($key, $hash.$key)
 }
 


$body = 'printerName='+$printerNameURL+'&printParamsXml=&labelXml='+$labelXmlURL+'&labelSetXml='+$labelSetXmlURL

# Write-Host $body

# skriv ut via Webservice
$uri="https://localhost:41951/DYMO/DLS/Printing/PrintLabel2"
Invoke-RestMethod -Uri $uri -Method POST -ContentType "application/x-www-form-urlencoded" -Body $body
Invoke-RestMethod -Uri $uri -Method POST -ContentType "application/x-www-form-urlencoded" -Body $body



# WEND

}