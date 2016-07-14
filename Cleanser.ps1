#Global Variable for masking file
$global:destfile="NULL"
# Function to write output in the csv file
function writecsv{
param([string]$link,[string]$filename,[string]$change,[int]$value)
$project=$global:destfile
$output   = "$PSScriptRoot\Config\projectxyz.csv"
$myobj = "" | Select "Project_Name","Link","File_Name","Change","Value"
        #fill the object
        $myobj.Project_Name = $project
        $myobj.Link = $link
        $myobj.File_Name = $filename
        $myobj.Change = $change
        $myobj.Value = $value
        
if($output)
{
$myobj | Export-Csv $output -NoTypeInformation -Append
}
else
{
$myobj | Export-Csv $output -NoTypeInformation
}
}
###########################################################################################
# Function to acquire path of the 7z.exe
function Get-7ZipExecutable
{
    $7zipExecutable = "C:\Program Files\7-Zip\7z.exe"
    return $7zipExecutable
}

# Function to zip folders where the destination is set to
function 7Zip-ZipDirectories
{
    param
    (
        [CmdletBinding()]
        [Parameter(Mandatory=$true)]
        [System.IO.DirectoryInfo[]]$include,
        [Parameter(Mandatory=$true)]
        [System.IO.FileInfo]$destination
             )

    $7zipExecutable = Get-7ZipExecutable
  
     # All folders in the destination path will be zipped in .7z format
     foreach ($directory in $include)
    {
        $arguments = "a","$($destination.FullName)","$($directory.FullName)"
    (& $7zipExecutable $arguments)
        
        $7ZipExitCode = $LASTEXITCODE
        
        if ($7ZipExitCode -ne 0)
        {
            $destination.Delete()
            throw "An error occurred while zipping [$directory]. 7Zip Exit Code was [$7ZipExitCode]."
        }
    }

    return $destination
}

# Function to unzip files
function 7Zip-Unzip
{
    param
    (
        [CmdletBinding()]
        [Parameter(Mandatory=$true)]
        [System.IO.FileInfo]$archive,
        [Parameter(Mandatory=$true)]
        [System.IO.DirectoryInfo]$destinationDirectory
    )

    $7zipExecutable = Get-7ZipExecutable
    $archivePath = $archive.FullName
    $destinationDirectoryPath = $destinationDirectory.FullName

    (& $7zipExecutable x "$archivePath" -o"$destinationDirectoryPath" -aoa -r)

    $7zipExitCode = $LASTEXITCODE
    if ($7zipExitCode -ne 0)
    {
        throw "An error occurred while unzipping [$archivePath] to [$destinationDirectoryPath]. 7Zip Exit Code was [$7zipExitCode]."
    }

    return $destinationDirectory
}
########################################################################################
#Replaces according to masking file input
function replaceFileName
{
	param([string]$str1,[string]$str2,[string]$path)

    echo "Replacing the File Name"
    get-childItem -Path $path -Filter “*$str1*” -Recurse |
	rename-item -newname { $_.name -replace $str1,$str2}
    	writecsv Link FileNameChange $str1 1
	
}
########################################################################################
#Function to remove Images from Headers And Footers 

function HeaderFooterImage
{    param($path)
	      [Console]::Out.Flush()
          $list=$null      	
    $list = gci -Path $path -Include *.docx,*.doc -Force -recurse -Exclude ~$*
    Add-Type -AssemblyName Microsoft.Office.Interop.Word
    $count=0	

    foreach ($foo in $list) 
	{
        echo "Deleting Images from $foo"
    	$counter = 0
        $objWord = New-Object -ComObject word.application
		$objWord.Visible = $False

		$objDoc = $objWord.Documents.Open("$foo")
		$objSelection = $objWord.Selection 
        
        foreach ($Section in $objDoc.Sections)
        {
            # Update Header
            $Header = $Section.Headers.Item(1)
             foreach ( $obj in $Header.Range.InlineShapes)
             {
                $obj.Delete();
                $count=$count+1
            }
           }
 foreach ($Section in $objDoc.Sections)
        {
            # Update Header
            $Footer = $Section.Footers.Item(1)
             foreach ( $obj in $Footer.Range.InlineShapes)
             {
                $obj.Delete();
                 $count=$count+1
            }
           }

      #  writecsv Link $foo.PSChildName ImagesRemoved $count
      echo "Link $foo.PSChildName ImagesRemoved $count"
        $objDoc.save()
	$objDoc.close()
	$objWord.quit()

	}

}


########################################################################################
#Function to work for Visio files

function Visio
{
param($str1,[string]$str2,$path,$id="")
Add-type -AssemblyName office
$vsoFind="$str1"
$vsoReplace="$str2"
$count = 0
$visio_app = New-Object -ComObject Visio.Application 
Get-ChildItem -Path $path -Include *.vsd -Recurse |
ForEach-Object {

            echo "Doing Replace operation on VISIO FILE $_.fullname"                                   
            #This will open Visio vsd File
            $doc = $visio_app.Documents.Open($_.fullname)
            
                foreach  ($vsoPage in $visio_app.ActiveDocument.Pages)
                {
                    foreach ($vsoShape in  $vsoPage.Shapes)
                  {
                       $text = $vsoShape.Text
                       $checktext=$text
                      $text=$text -replace $vsoFind,$vsoReplace
                      $vsoShape.Text=$text
                      if($text -eq $checktext){
                                                $count=$count+1
                      }
    
                  } 
                }
                $doc.SaveAs($_.fullname)

                  if($id -eq ""){ 
                                    writecsv Link $_.name $str1 $count
		                        }
                                    else
                                    {
                                        writecsv Link $_.name $id $count
		                            }


 $count = 0
 $visio_app.Quit()
 }
 [gc]::collect()
[gc]::WaitForPendingFinalizers()
}

########################################################################################
#Function to work for Regular Expression 

function awordRegEx
{
	param([string]$str1,[string]$str2,$path,$id)
	      [Console]::Out.Flush()
          $list=$null      	
    $list = gci -Path $path -Include *.docx,*.doc -Force -recurse -Exclude ~$*
    Add-Type -AssemblyName Microsoft.Office.Interop.Word
	
    foreach ($foo in $list) 
	{
        echo "Doing Replace operation on Word File $foo"
    	$counter = 0
        $objWord = New-Object -ComObject word.application
		$objWord.Visible = $False
		$objDoc = $objWord.Documents.Open("$foo")
		$objSelection = $objWord.Selection 
        # FOR HEADERS AND FOOTERS
        $findtext= "$str1"
		$ReplaceText = "$str2"
		$wdReplaceOne = 1
		$FindContinue = 1
		$MatchFuzzy = $False
		$MatchCase = $True
		$MatchPhrase = $false
		$MatchWholeWord = $True
		$MatchWildcards = $True
		$MatchSoundsLike = $False
		$MatchAllWordForms = $False
		$Forward = $True
		$Wrap = $FindContinue
		$Format = $False
	    while($true){	
    
            [bool]$rec = $objSelection.Find.execute(
				$FindText,
			   $MatchCase,
				$MatchWholeWord,
				$MatchWildcards,
				$MatchSoundsLike,
				$MatchAllWordForms,
				$Forward,
				$Wrap,
				$Format,
				$ReplaceText,
				$wdReplaceOne
			)
            
            if($rec){
            $counter ++
            }
            else
            {
             Break

            }

        }
        $count=0
        foreach ($Section in $objDoc.Sections)
        { 
             # Update Header
             $Header = $Section.Headers.Item(1)
            
            while($true){	
    
            [bool]$rec = $Header.Range.find.Execute($FindText, 
              $true, #match case
              $true, #match whole word
              $true, #match wildcards
              $false, #match soundslike
              $false, #match all word forms
              $true,  #forward
              $findWrap, 
              $null,      #format
              $ReplaceText,
				$wdReplaceOne)
                  
            if($rec){
            $count ++
            }
            else {
             Break

            }

            }
        

            # Update Footer
            $Footer = $Section.Footers.Item(1)
            while($true){	
    
            [bool]$rec = $Footer.Range.find.Execute($FindText, 
              $true, #match case
              $true, #match whole word
              $true, #match wildcards
              $false, #match soundslike
              $false, #match all word forms
              $true,  #forward
              $findWrap, 
              $null,      #format
              $ReplaceText,
				$wdReplaceOne)
                
            if($rec){
            $count ++
            }
            else {
             Break

            }

            }
        }

        #Even if there is no replacements done it will update , need to rectify the whole code 
        writecsv Link $foo.PSChildName $id $counter
        writecsv Link $foo.PSChildName "$id headerfooter" $count
        $objDoc.save()
	$objDoc.close()
	$objWord.quit()

	}

}

# Function to Find and Replace a given Regular Expression in PowerPoint
function apptRegEx
{
param($str1,[string]$str2,$path,$id)

Add-type -AssemblyName office

$find = "$str1"
$replace = "$str2"


function ReplaceText {
  param(
      [object]$shape,
      [string]$find,
      [string]$replace
      )
$counting=0

  if ($shape.HasTextFrame)
  {
     $counting= $counting + 1
     $textFrame = $shape.TextFrame
     $textRange = $textFrame.TextRange
       if ($shape.TextFrame.HasText)
       {
       $strInput = $shape.TextFrame.TextRange.Text
       $strInput = $strInput -replace $str1,$str2
       $shape.TextFrame.TextRange = $strInput
       }
       <#
     $paragraphs = $textRange.Paragraphs()
     foreach ($paragraph in $paragraphs)
     {
        $text = $paragraph.Text
        if($text.Contains($find)) {
            $text = $text.Replace($find, $replace)
            $paragraph.Text = $text
            
        }
        #>
     
  }
  $counting
}

$countppt = 0
$Application = New-Object -ComObject powerpoint.application
$msoTrue = [Microsoft.Office.Core.MsoTriState]::msoTrue
$msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
			#$application.visible = $msoTrue
Get-ChildItem -Path $path -Include *.pptx,*.ppt -Recurse |
ForEach-Object {

            echo "Doing Replace operation on PowerPoint File $_.fullname"                                   

  #Open presentation with ReadOnly:False, Untitled:False, Visible:True
 $presentation = $application.Presentations.Open($_.fullname, $msoFalse, $msoFalse, $msoFalse)

         foreach ($slide in $presentation.Slides) {
             foreach ($shape in $slide.Shapes) {
              
                if ($shape.Type -eq 6) {
                    foreach ($item in $shape.GroupItems) { 
                     $co =   ReplaceText $item $find $replace
                        }
                } else {
                 $co = ReplaceText $shape $find $replace
                 
               }
               $countppt = $countppt +$co
            }
         } 
writecsv Link $_.name $id $countppt
 
 $countppt = 0
 $presentation.Save()
 $presentation.Close()
} 

$Application.Quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
}

####################################################################################################

# Function to Find and Replace a given string in PowerPoint
function appt
{
param([string]$str1,[string]$str2,$path)

Add-type -AssemblyName office

$find = "$str1"
$replace = "$str2"


function ReplaceText {
  param(
      [object]$shape,
      [string]$find,
      [string]$replace
      )
$counting=0

  if ($shape.HasTextFrame)
  {
     $counting= $counting + 1
     $textFrame = $shape.TextFrame
     $textRange = $textFrame.TextRange
     <#
     $textRange.Replace($find, $replace, 0, $msoFalse, $msoFalse) 
     #>
     #<#
     #Use this if above replacement causes formatting to be lost.
 
     $paragraphs = $textRange.Paragraphs()
     foreach ($paragraph in $paragraphs)
     {
        $text = $paragraph.Text
        if($text.Contains($find)) {
            $text = $text.Replace($find, $replace)
            $paragraph.Text = $text
            
        }
        
     }
    
  }
  $counting
}

$countppt = 0
$Application = New-Object -ComObject powerpoint.application
$msoTrue = [Microsoft.Office.Core.MsoTriState]::msoTrue
$msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
			#$application.visible = $msoTrue

                                  
Get-ChildItem -Path $path -Include *.pptx,*.ppt -Recurse |
ForEach-Object {
echo "Doing Replace operation on PowerPoint Files $_.fullname" 
  #Open presentation with ReadOnly:False, Untitled:False, Visible:True
 $presentation = $application.Presentations.Open($_.fullname, $msoFalse, $msoFalse, $msoFalse)

         foreach ($slide in $presentation.Slides) {
             foreach ($shape in $slide.Shapes) {
                # [Microsoft.Office.Core.MsoShapeType]::msoGroup
                if ($shape.Type -eq 6) {
                    foreach ($item in $shape.GroupItems) { 
                     $co =   ReplaceText $item $find $replace
                        }
                } else {
                 $co = ReplaceText $shape $find $replace
                 
               }
               $countppt = $countppt +$co
            }
         } 
    writecsv Link $_.name $str1 $countppt
 
 $countppt = 0
 $presentation.Save()
 $presentation.Close()
} 

$Application.Quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
}

# Function to Find and Replace a given string in Word Documents

function aword{
param([string]$str1,[string]$str2,$path)
	[Console]::Out.Flush()
    $list=$null      	
    $list = gci -Path $path -Include *.docx,*.doc -Force -recurse -Exclude ~$*
    Add-Type -AssemblyName Microsoft.Office.Interop.Word
	echo "Processing Find And Replace Operation in word"	
    foreach ($foo in $list) 
	{
	    echo "Processing Find And Replace Operation in word $foo"	
    
    	$counter = 0
        $objWord = New-Object -ComObject word.application
		$objWord.Visible = $False
		$objDoc = $objWord.Documents.Open("$foo")
		$objSelection = $objWord.Selection 
        # FOR HEADERS AND FOOTERS
        $findtext= "$str1"
		$ReplaceText = "$str2"
		$wdReplaceOne = 1
		$FindContinue = 1
		$MatchFuzzy = $False
		$MatchCase = $False
		$MatchPhrase = $false
		$MatchWholeWord = $True
		$MatchWildcards = $False
		$MatchSoundsLike = $False
		$MatchAllWordForms = $False
		$Forward = $True
		$Wrap = $FindContinue
		$Format = $False
	    while($true){	
    
            [bool]$rec = $objSelection.Find.execute(
				$FindText,
			   $MatchCase,
				$MatchWholeWord,
				$MatchWildcards,
				$MatchSoundsLike,
				$MatchAllWordForms,
				$Forward,
				$Wrap,
				$Format,
				$ReplaceText,
				$wdReplaceOne
			)
            
            if($rec){
            $counter ++
            }
            else
            {
             Break

            }

        }
        $count=0
        foreach ($Section in $objDoc.Sections)
        { 
             # Update Header
             $Header = $Section.Headers.Item(1)
            
            while($true){	
    
            [bool]$rec = $Header.Range.find.Execute($FindText, 
              $false, #match case
              $true, #match whole word
              $false, #match wildcards
              $false, #match soundslike
              $false, #match all word forms
              $true,  #forward
              $findWrap, 
              $null,      #format
              $ReplaceText,
				$wdReplaceOne)
                  
            if($rec){
            $count ++
            }
            else {
             Break

            }

            }
        

            # Update Footer
            $Footer = $Section.Footers.Item(1)
            while($true){	
    
            [bool]$rec = $Footer.Range.find.Execute($FindText, 
              $false, #match case
              $true, #match whole word
              $false, #match wildcards
              $false, #match soundslike
              $false, #match all word forms
              $true,  #forward
              $findWrap, 
              $null,      #format
              $ReplaceText,
				$wdReplaceOne)
                
            if($rec){
            $count ++
            }
            else {
             Break

            }

            }
        }
        writecsv Link $foo.PSChildName $str1 $counter
        writecsv Link $foo.PSChildName "$str1 headerfooter" $count
        $objDoc.save()
	$objDoc.close()
	$objWord.quit()

	}
}


# Function to find and replace in excel 
function aexcel
{
    param([string]$str1,[string]$str2,$path,[string]$id="")
    
    
    $list = gci -Path $path -Include *.xlsx,*.xls -Force -recurse
    
    #Loop to move acroos the files
	foreach ($foo in $list) 
	{
		echo "Doing Replace operation on Excel Files $foo"
        [Console]::Out.Flush() 
        $exApp = New-Object -ComObject Excel.Application
	    $exApp.DisplayAlerts = $false
    	$exBook = $exApp.Workbooks.Open("$foo")
        $counter=0 
			foreach ($sheet in $exBook.Sheets)
			{
				$cells = $sheet.UsedRange.Value2
				if($cells -ne "")
				{
				if ($cells.GetType().Name -eq "String")
				{
					if ($cell -eq "$str1")
						{    
			                        $counter=$counter+1                          
						$cell = $cells -replace ("$str1", "$str2")
                                 		}
					
					$sheet.UsedRange.Value2 = $cell
				}
	
				else
				{
					for ($i = 1; $i -le $cells.GetUpperBound(0); $i++)
					{
						for ($j = 1; $j -le $cells.GetUpperBound(1); $j++)
						{
                                $cell = $null
                                $cell = $cells[$i,$j]
                                
                                $cellmatch = $cells[$i,$j]
                                
                                $cells[$i,$j] = $cell -replace ("$str1", "$str2") 
                               
                                 if ($cells[$i,$j] -ne "" )
                                 {
                                 if ($cells[$i,$j] -ne $cellmatch)
                                   {
                                    $counter=$counter+1
							        
                                    }
                                  }
						}
					}
	
					$sheet.UsedRange.Value2 = $cells
				}
			}
			}
        if($id -eq ""){ 
        writecsv Link $foo.PSChildName $str1 $counter
		}
        else
        {
        writecsv Link $foo.PSChildName $id $counter
		}

        	$exBook.Save()
		$exBook.Close()
		$exApp.Quit()
        	
	}

}
###########################################################################################
# FUnction to remove all file properties from PPT
function pptmeta
{
            
    $path = "$PSScriptRoot\Working"

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Add-type -AssemblyName office

$PpRemoveDocType = "Microsoft.Office.Interop.PowerPoint.PpRemoveDocInfoType" -as [type] 
$PPTFiles = Get-ChildItem -Path $path -include *.ppt, *.pptx -recurse 
$objPPT = New-Object -ComObject powerpoint.application
$msoTrue = [Microsoft.Office.Core.MsoTriState]::msoTrue
$msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
		#$objPPT.visible = $msoTrue

foreach($obj in $PPTFiles) 
{ 
    
    #Open presentation with ReadOnly:False, Untitled:False, Visible:True
    $presentations = $objPPT.Presentations.Open($obj.fullname, $msoFalse,$msoFalse,$msoFalse) 
    "Removing document information from $obj" 
    $presentations.RemoveDocumentInformation($PpRemoveDocType::ppRDIAll) 
    $presentations.Save() 
    $presentations.Close()
    writecsv Link $obj.PSChildName FILE-PROPERTY 1

} 
$objPPT.Quit()
}


# FUnction to remove all file properties from Word
function wordmeta
{ 
   $path = "$PSScriptRoot\Working"
Add-Type -AssemblyName Microsoft.Office.Interop.Word
$WdRemoveDocType = "Microsoft.Office.Interop.Word.WdRemoveDocInfoType" -as [type] 
$WdAlertsNone = "Microsoft.Office.Interop.Word.WdAlertsNone" -as [type]
$wordFiles = Get-ChildItem -Path $path -include *.doc, *.docx -recurse 
$objword = New-Object -ComObject word.application 
#$objword.DisplayAlerts =$wdAlertsNone
$objword.visible = $false 

foreach($obj in $wordFiles) 
{ 
    $documents = $objword.Documents.Open($obj.fullname) 
    "Removing document information from $obj" 
    # 99 = WdRDIAll
    #$documents.RemoveDocumentInformation(99)
    $documents.RemoveDocumentInformation($WdRemoveDocType::wdRDIAll) 
    $documents.Save() 
    $objword.documents.close() 
    writecsv Link $obj.PSChildName FILE-PROPERTY 1
} 
$objword.Quit()
}


# FUnction to remove all file properties from Excel

function excelmeta
{
$path = "$PSScriptRoot\Working"
Add-Type -AssemblyName Microsoft.Office.Interop.Excel 

$xlRemoveDocType = “Microsoft.Office.Interop.Excel.XlRemoveDocInfoType” -as [type] 

$excelFiles = Get-ChildItem -Path $path -include *.xls, *.xlsx -recurse 
$objExcel = New-Object -ComObject excel.application 
$objExcel.DisplayAlerts = $false
$objExcel.visible = $false 

foreach($wb in $excelFiles) 
{ 
$workbook = $objExcel.workbooks.open($wb.fullname) 
“Removing document information from $wb” 
$workbook.RemoveDocumentInformation($xlRemoveDocType::xlRDIAll) 
$workbook.Save() 
$objExcel.Workbooks.close()
writecsv Link $wb.PSChildName FILE-PROPERTY 1 
} 
$objExcel.Quit()
}

#######################################################################################
#Function to Count Images in Excel
function excelImageCount{
    param($path)
    $list = gci -Path $path -Include *.xlsx,*.xls -Force -recurse
        #Loop to move acroos the files
	foreach ($foo in $list) 
	{
        echo "Counting Images In $foo"
		[Console]::Out.Flush() 
        $count=0;
        $exApp = New-Object -ComObject Excel.Application
		$exBook = $exApp.Workbooks.Open("$foo")
        $exApp.DisplayAlerts = $false
        
			foreach ($sheet in $exBook.Sheets)
			{
				$count=$count + $sheet.Shapes.Count
                
               }
        writecsv Link $foo.PSChildName Image $count
	    $exBook.Save()
		$exBook.Close()
		$exApp.Quit()
    }
}
#Function to Count Images in PPT
function pptImageCount
{
param($path)
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Add-type -AssemblyName office
$countppt = 0
$Application = New-Object -ComObject powerpoint.application
$msoTrue = [Microsoft.Office.Core.MsoTriState]::msoTrue
$msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
$msoMedia = [Microsoft.Office.Core.MsoShapeType]::msoMedia;
$msoGroup = [Microsoft.Office.Core.MsoShapeType]::msoGroup;
$msoPicture = [Microsoft.Office.Core.MsoShapeType]::msoPicture;
		#$application.visible = $msoTrue
Get-ChildItem -Path $path -Include *.pptx,*.ppt -Force -Recurse |
ForEach-Object {
    echo "Counting Images In $_.fullname"
    #Open presentation with ReadOnly:False, Untitled:False, Visible:True
    $presentation = $application.Presentations.Open($_.fullname, $msoFalse, $msoFalse, $msoFalse)
           foreach ($slide in $presentation.Slides) {
             foreach ($shape in $slide.Shapes) {
                        if ($shape.Type -eq $msoMedia) {
                        $countppt = $countppt + 1
                        }
                        if ($shape.Type -eq $msoPicture) {
                        $countppt = $countppt + 1
                        }
                     } 
            }
            writecsv Link $_.name Image $countppt
            $presentation.Save()
            $presentation.Close()
         } 
  $Application.Quit()
#$application = $null
#[gc]::collect()
#[gc]::WaitForPendingFinalizers()
}

#Function to Count Images in Word

function wordImageCount
{
	param($path)
	[Console]::Out.Flush()
    $list=$null      	
    $list = gci -Path $path -Include *.docx,*.doc -Force -recurse
    foreach ($foo in $list) 
	{
        echo "Counting Images In $foo"
        $objWord = New-Object -ComObject word.application
        $objWord.Visible = $False
        $objDoc = $objWord.Documents.Open("$foo")
        $count = $objDoc.Inlineshapes.count
        writecsv Link $foo.PSChildName Image $count
 		$objDoc.save()
		$objWord.quit()
	}

}
##########################################################################################
#Function to replace email occurences in Excel and Word  
function emailReplacer()
{
param($path)
echo "We are replacing Email ID's now"
$str1="[A-z,0-9,.,_,-]{1,}\@[A-z,0-9,\.]{1,}"
$str2="email"
visio $str1 $str2 $path "Email"
awordRegEx $str1 $str2 $path "Email"
aexcel $str1 $str2 $path "Email"
apptRegEx $str1 $str2 $path "Email"
}

###########################################################################################
#Function to replace phone nos.
 function phoneReplacer()
 {

 echo "Will replace Phone Numbers Now" 
  $path = "$PSScriptRoot\Working"

$str1ex="(?:(?:\+|0{0,2})91(\s*[\-]\s*)?|[0]?)?[789]\d{9}"#mobile number all format
$str2ex="(\d+[ \-]+\d+)"#local all format
$str3ex="(\+\d)*\s*(\(\d{3}\)\s*)*\d{3}(-{0,1}|\s{0,1})\d{2}(-{0,1}|\s{0,1})\d{2}"#for intermational format
$rep="phoneno"
$str4ex="\+(?:[0-9] ?){6,14}[0-9]"#it is for those like+1 1234567890123 ,+12 123456789 : +123 123456 format
$str5ex="\+?\(?\d{2,4}\)?[\d\s-]{3,}"# validates +61 8 9650 5000

 aexcel $str1ex $rep $path "PhoneNo"
 aexcel $str2ex $rep $path "PhoneNo"
 aexcel $str3ex $rep $path "PhoneNo"
 aexcel $str4ex $rep $path "PhoneNo"
 aexcel $str5ex $rep $path "PhoneNo"
 
 visio $str1ex $rep $path "PhoneNo"
 visio $str2ex $rep $path "PhoneNo"
 visio $str3ex $rep $path "PhoneNo"
 visio $str4ex $rep $path "PhoneNo"
 visio $str5ex $rep $path "PhoneNo"
 

 apptRegEx $str1ex $rep $path "PhoneNo"
 apptRegEx $str2ex $rep $path "PhoneNo"
 apptRegEx $str3ex $rep $path "PhoneNo"
 apptRegEx $str4ex $rep $path "PhoneNo"
 apptRegEx $str5ex $rep $path "PhoneNo"
  
# Word Function is creating a little problem with temporary files and script gets stopped so it is commented out for now
 
$str1wo="[0]*?[789][0-9]{9}"#it works for those like 09422630995 ,0 9422630995
$str2wo="[0-9]{5}?[0-9]{6}"#03595-259506 basically all standard std phone numbers
$str3wo="[789][0-9]{9}"#all normal phone numbers
$str4wo="[+]*[91]?[789][0-9]{9}"#all starting with +91 or 91
$str5wo="?*[0-9]{3}?*[0-9]{3}?*[0-9]{4}"#standard format for international 800-555-5555 | 
 #333-444-5555 | 212-666-1234|000-000-0000 | 123-456-7890 |(123) 456-7890 | 123-456-7890
#<#
 awordRegEx $str1wo $rep $path "PhoneNo"
 awordRegEx $str2wo $rep $path "PhoneNo"
 awordRegEx $str3wo $rep $path "PhoneNo"
 awordRegEx $str4wo $rep $path "PhoneNo"
 #awordRegEx $str5wo $rep $path "PhoneNo"
 #>
 
 }    
        
 #########################################################################################
 #Function to handle all the Word , PPt , Excel replace functions , this function picks up values from masking file
 
 function replaceHandler{
       param($path)
       	
        #Function to read values from the excel file and invoke find replace of all other files

 
    	$list = gci "$PSScriptRoot\Config"  -Include *.xlsx,*.xls -Force -recurse
        
        foreach ($file in $list) 
	    {
		$global:destfile=[io.path]::GetFileNameWithoutExtension($file)
        $destfile
        $sheetName = "Sheet1"
		
		#Create an instance of Excel.Application and Open Excel file
		$objExcel = New-Object -ComObject Excel.Application
	    $objExcel.DisplayAlerts = $false
		$workbook = $objExcel.Workbooks.Open($file)
		$sheet = $workbook.Worksheets.Item($sheetName)
		
		$objExcel.Visible=$false
	
		#Count max row
		$rowMax = ($sheet.UsedRange.Rows).count
	
		#Declare the starting positions
		$rowA,$colA = 1,1
		$rowB,$colB = 1,2

		#loop to get values and store it
		for ($i=1; $i -le $rowMax-1; $i++)
			{
				[string]$var1 = $sheet.Cells.Item($rowA+$i,$colA).text
				[string]$var2 = $sheet.Cells.Item($rowB+$i,$colB).text
				
                visio $var1 $var2 $path
				aexcel $var1 $var2 $path
				aword $var1 $var2 $path
				appt $var1 $var2 $path
				replaceFileName $var1 $var2 "$path"
			}
	        #close excel file
            $workbook.Save()
            $workbook.Close()
            $objExcel.Quit()
	    }
    }     
 function filePropertyHandler{
        wordmeta
        excelmeta
        pptmeta
                }
 function imageCountHandler{
                param($path)
                wordImageCount $path
                excelImageCount $path
                pptImageCount $path
                
                }
#Function to unzip files and also unzip the files present inside the zip file if any and write output on csv
 function unzip{
$path = "$PSScriptRoot\Working"
$i = 0
        


    while ($i -lt 10) {
        $list = Get-ChildItem -Path $path -Include *.rar,*.zip,*.7z -recurse
        foreach ($file in $list) {
             #7Zip-Unzip "D:\PROJECT\Working\zipperomg.7z" "D:\PROJECT\Working\zipperomg1"
             #$destfile= [io.path]::GetFileNameWithoutExtension($path)
             $dest=Join-Path([System.IO.Path]::GetDirectoryName($file))([System.IO.Path]::GetFileNameWithoutExtension($file))
             7Zip-Unzip $file $dest
             writecsv Link $file.PSChildName UNZIPED 1
             Remove-Item $file
             replaceHandler $dest
              emailReplacer $path
                phoneReplacer
             imageCountHandler $path
             filePropertyHandler
             } 
       if ($list -eq $null) {break}
    }
}

#Function to zip the folders present in the given directory will invoke the zip-directory function
 function zip{
$path = "$PSScriptRoot\Working"

        $list = Get-ChildItem -Path $path -dir

        foreach ($folder in $list) {
             #7Zip-Unzip "D:\PROJECT\Working\zipperomg.7z" "D:\PROJECT\Working\zipperomg1"
             #$destfile= [io.path]::GetFileNameWithoutExtension($path)
             $dest=$folder.FullName
             7Zip-ZipDirectories $dest $dest
             Remove-Item $dest -Recurse
             writecsv Link $folder.PSChildName ZIPPED 1
            } 
}
 
 #Main function to handle all functionalities
 function mainHandler{
     	echo "-- This takes time generally, please be patient, in the mean time you can keep checking on the changes done"
	echo "Process is Starting"
	
     $path = "$PSScriptRoot\Working"
     replaceHandler $path
      emailReplacer $path
      phoneReplacer
      imageCountHandler $path
      HeaderFooterImage $path
    
     filePropertyHandler
     unzip
     zip
        
   }
   
 #Calling the main 
        mainHandler
        write-host EndOfTask
