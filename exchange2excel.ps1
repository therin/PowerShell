# Sergey Borodich, 2016
# This PowerShell script does the following:

# 1. Reads specified Exchange mailbox and saves all xls attachments to a predefined folder
# 2. Moves mailbox items to specified folder
# 3. Merges all found xls files into one
# 4. Replacing one of the column names in a resulting file to avoid duplicates

# Prior to running script, create a target box for processed mail under Inbox named "Processed"

###################################################################
# Config
###################################################################

# Set Exchange Version (Exchange2010, Exchange2010_SP1, Exchange2010_SP2, etc.)
$MSExchangeVersion = "Exchange2010_SP3"

# Set URL for EWS
$EWSUrl = "https://mail.some.com/ews/Exchange.asmx"

# Set mailbox credentials
$user = "some.login"

# No plain text passwords in here, use "somecoolpass" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "C:\SiSense\Password.txt" to update

$password = Get-Content "C:\SiSense\Password.txt" | ConvertTo-SecureString
$domain = "some.com"
$boxname = "some.user@some.com"

# Download directory
$downloadDirectory = "c:\SiSense\CodeBlue"

###################################################################
# Connect to Exchange Web Services
###################################################################

# Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

# Create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

# Load EWS API and attach to CAS & EWS

Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"

# Create Exchange Service Object
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

$creds = New-Object System.Net.NetworkCredential($user,$password,$domain) 
$service.Credentials = $creds 

$MailboxName = $boxname
$uri=[system.URI] $EWSUrl
$service.Url = $uri 

###################################################################
# Process mailbox
###################################################################

# Bind to CodeBlue folder

$PathToSearch = "\CodeBlue"  
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,"some@some.com")   
$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
#Split the Search path into an array  
$fldArray = $PathToSearch.Split("\") 
#Loop through the Split Array and do a Search for each level of folder 
for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
        $fldArray[$lint] 
        #Perform search based on the displayname of each folder level 
        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$fldArray[$lint]) 
        $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
        if ($findFolderResults.TotalCount -gt 0){ 
            foreach($folder in $findFolderResults.Folders){ 
                $tfTargetFolder = $folder                
            } 
        } 
        else{ 
            "Error Folder Not Found"  
            $tfTargetFolder = $null  
            break  
        }     
}  

$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$tfTargetFolder.Id)  

# Find attachments, copy to download directory

$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
$findItemsResults = $Inbox.FindItems($Sfha,$ivItemView)
foreach($miMailItems in $findItemsResults.Items){
	$miMailItems.Load()
	foreach($attach in $miMailItems.Attachments){

		# Only extract XLS attachments. If you need additional filetypes, include them as an OR in the second if below. To extract all attachments, remove these two
        # if loops
		If($attach -is[Microsoft.Exchange.WebServices.Data.FileAttachment]){
			if($attach.Name.Contains(".xls")){  
				$attach.Load()

				# Add random # to filename to ensure unique

				$prefix = Get-Random	
				$fiFile = new-object System.IO.FileStream(($downloadDirectory + "\" + $prefix + $attach.Name.ToString()), [System.IO.FileMode]::Create)    		

				$fiFile.Write($attach.Content, 0, $attach.Content.Length)
				$fiFile.Close()
			}
		}
	}
}


# Get the ID of the folder to move to  
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)  
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"Processed")
$findFolderResults = $Inbox.FindFolders($SfSearchFilter,$fvFolderView)  

# Define ItemView to retrive (max 100 items)
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)  
$fiItems = $null    
do{    
    $fiItems = $Inbox.FindItems($Sfha,$ivItemView)   
    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
        foreach($Item in $fiItems.Items){      
            # Move   
            $Item.Move($findFolderResults.Folders[0].Id)  
        }    
        $ivItemView.Offset += $fiItems.Items.Count    
    }while($fiItems.MoreAvailable -eq $true)  


#Get a list of files to copy from
$Files = GCI 'C:\SiSense\CodeBlue\' | ?{$_.Extension -Match "xls?"} | select -ExpandProperty FullName
echo $Files

# Stop script execution if there are no files to work with:
If ($Files -eq $null) {
    Write-Host "File Not Found"
    Exit
}

#Launch Excel, and make it do as its told (supress confirmations)
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $True
$Excel.DisplayAlerts = $False

#Open up a new workbook
$Dest = $Excel.Workbooks.Add()

#Loop through files, opening each, selecting the used range, and only grabbing the first 6 columns of it. Then find next available row on the destination worksheet
# and paste the data
ForEach($File in $Files){
    $Source = $Excel.Workbooks.Open($File,$true,$true)
    If(($Dest.ActiveSheet.UsedRange.Count -eq 1) -and ([String]::IsNullOrEmpty($Dest.ActiveSheet.Range("A1").Value2))){ #If there is only 1 used cell and it is blank 
    #select A1
        [void]$source.ActiveSheet.Range("A1","AF$(($Source.ActiveSheet.UsedRange.Rows|Select -Last 1).Row)").Copy()
        [void]$Dest.Activate()
        [void]$Dest.ActiveSheet.Range("A1").Select()
    }Else{ #If there is data go to the next empty row and select Column A
        [void]$source.ActiveSheet.Range("A2","AF$(($Source.ActiveSheet.UsedRange.Rows|Select -Last 1).Row)").Copy()
        [void]$Dest.Activate()
        [void]$Dest.ActiveSheet.Range("A$(($Dest.ActiveSheet.UsedRange.Rows|Select -last 1).row+1)").Select()
    }
    [void]$Dest.ActiveSheet.Paste()
    $Source.Close()
    Remove-Item $File

}
# Replacing one of the column names to avoid duplicates
$ws=$Dest.WorkSheets.item(1)
$ws.Cells.Item(1,14).Value2 = "Follow-Up Comments"
$Dest.SaveAs("C:\SiSense\CodeBlue\Combined\code_blue.xls",51)
$Dest.close()
$Excel.Quit()


