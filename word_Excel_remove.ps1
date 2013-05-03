##Script for removing metadata on Excel/Word data files
##Cameron McCoy / Jeff Flagg 
##Computer Forensics Tool Project May 3 2013
##
##Sourced from Ed Wilson (Microsoft 8/10/2008)
##             Scriptingguy1 (technet.com 9/7/2010)
####################################################

#enter selection
#ask user location of files to be sanitized
$path = Read-Host 'Enter folder location'

while ($continue -eq 100)
{
$continue = 100
#Adds microsoft office excel/word assembly
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName Microsoft.Office.Interop.Word

# start variable declarations 
$xlRemoveDocType = "Microsoft.Office.Interop.Excel.XlRemoveDocInfoType" -as [type] 
$wrdRemoveDocType = "Microsoft.Office.Interop.Word.WdRemoveDocInfoType" -as [type]

#Find all files in chosen path
$excelFiles = Get-ChildItem -Path $path -include *.xls, *.xlsx -recurse
$wordFiles = Get-ChildItem -Path $path -include *.doc, *.docx -recurse

#creates object to open 
$objExcel = New-Object -ComObject excel.application 
$objWord = New-Object -ComObject word.application 

#makes sure new application does not appear on screen
$objExcel.visible = $false 
$objWord.visible = $false

###Function Section#############################################

#Creates different colored lines for better readablity 
Function funLine($strIN)  
{ 
  $strLine = "=" * $strIn.length 
  Write-Host -ForegroundColor Yellow "`n$strIN" 
  Write-Host -ForegroundColor Cyan $strLine 
} 

#Print function for options menu
Function funMetaData() 
{ 
 foreach($sFolder in $path) 
  { 
   $a = 0 
   $objShell = New-Object -ComObject Shell.Application 
   $objFolder = $objShell.namespace($sFolder) 
 
   foreach ($strFileName in $objFolder.items()) 
    { FunLine( "$($strFileName.name)") 
      for ($a ; $a  -le 266; $a++) 
       {  
         if($objFolder.getDetailsOf($strFileName, $a)) 
           { 
             $hash += @{ ` 
                   $($objFolder.getDetailsOf($objFolder.items, $a))  =` 
                   $($objFolder.getDetailsOf($strFileName, $a))  
                   } #end hash 
            $hash 
            $hash.clear() 
           } #end if 
       } #end for  
     $a=0 
    } #end foreach 
  } #end foreach 
}
##############################################################################




#Print out Menu
$removal = Read-Host "
    Are the files Excel Documents? type 1
    Are the files Word Documents? type 2
    Show all metadata associated with found files 3
    Change path 4
    To quit type 5

    Please enter a selection number"

    
    #What data should be deleted
    if ($removal -eq 3)
    {
        funMetaData
        continue
    }
    #Excel Option
    if ($removal -eq 1)
    {
        #recursive loop to check evrey excel file 
       foreach($wb in $excelFiles) 
        { 
        $workbook = $objExcel.workbooks.open($wb.fullname) 
        "About to remove document information from $wb" 

        #to change what metadata is deleted, update "xlRDIALL" parameter (17 options)
        $workbook.RemoveDocumentInformation($xlRemoveDocType::xlRDIAll) 
        $workbook.Save() 
        $objExcel.Workbooks.close() 
        } 
        $objExcel.Quit()
    }
        
    #Word Option
    if ($removal -eq 2)
    {#recursive loop to check evrey word file 
    foreach($wb in $WordFiles) 
        { 
        $document = $objWord.documents.open($wb.fullname) 
        "About to remove document information from $wb" 

        #to change what metadata is deleted, update "xlRDIALL" parameter (17 options)
        $document.RemoveDocumentInformation($wrdRemoveDocType::wdRDIAll) 
        $document.Save() 
        $objWord.documents.close() 
        } 
        $objWord.Quit()

    }
    #Change $path option
    if ($removal -eq 4)
    {
        $path = Read-Host 'Enter folder location'
        continue
    }

    #Quit 
    if ($removal -eq 5)
    {
                $continue = 25
                
                exit
    }
 }
    
    
    




