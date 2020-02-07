########################################################################################################################## 
# Author: Zewwy (Aemilianus Kehler)
# Date:   Feb 4, 2020
# Script: Change-FilesName
# This script allows to remove/change characters from file names
# Required parameters: 
#   A File Path and selected characters of users choice
##########################################################################################################################

##########################################################################################################################
#   Static Variables
##########################################################################################################################

#MyLogoArray
$MylogoArray = @("#####################################","# This script is brought to you by: #","#                                   #","#             Zewwy                 #","#                                   #","#####################################"," ")
#Static Variables
$ScriptName = "Change-FilesName; Cause some user are.... creative. ¯\_(°_o)_/¯`n"

$pswheight = (get-host).UI.RawUI.MaxWindowSize.Height
$pswwidth = (get-host).UI.RawUI.MaxWindowSize.Width

##########################################################################################################################
#   Functions
##########################################################################################################################

#function takes in a name to alert confirmation of deletion returns true or false
function confirm($name)
{
    #function variables, generally only the first two need changing
    $title = "Confirm SharePoint Feature Removal!"
    $message = "You are about to remove a SharePoint Feature: $name"

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "This means Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "This means No"

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $Options, 0)
    Write-Host " "
    Switch ($result)
        {
              0 { Return $true }
              1 { Return $false }
        }
}

#Function to Centeralize Write-Host Output, Just take string variable parameter and pads it. Nerd Level over 9000!!! Ad-hoc Polymorphic power time!!
function Centeralize()
{
  param(
  [Parameter(Position=0,Mandatory=$true)]
  [string]$S,
  [Parameter(Position=1,Mandatory=$false,ParameterSetName="color")]
  [string]$C
  )
    $sLength = $S.Length
    $padamt =  "{0:N0}" -f (($pswwidth-$sLength)/2)
    $PadNum = $padamt/1 + $sLength #the divide by one is a quick dirty trick to covert string to int
    $CS = $S.PadLeft($PadNum," ").PadRight($PadNum," ") #Pad that shit
    if ($C) #if variable for color exists run below
    {    
        Write-Host $CS -ForegroundColor $C #write that shit to host with color
    }
    else #need this to prevent output twice if color is provided
    {
        $CS #write that shit without color
    }
}

#Ask for a Path, verify, return it
function GetPath()
{
    Write-host "Folder Path (C:\Temp): " -ForegroundColor Magenta -NoNewline
    $SubmittedPath = Read-Host
    if ( !$SubmittedPath -or !(Test-Path $SubmittedPath) )
    {
        Write-Host "`nThe path you provided is not valid.`n" -ForegroundColor Yellow
        GetPath
    }
    else
    {
        return $SubmittedPath
    }
}

#Ask for search characters, if null use defaults, return it
function GetChars()
{
    Write-host "Characters to search for !$'&{^,}~#%][)( " -ForegroundColor Magenta -NoNewline
    $UnsupportedChars = Read-Host
    if ( !$UnsupportedChars ){ $UnsupportedChars = "\]!$'&{^,+}~#%[)(" }
    return $UnsupportedChars
}

#Creates an array of objects based on the search characters criteria, and returns it, also extends the object to hold the matched character value
function FilterItems($Items, $RegEx)
{
    $shit = @()
    foreach ($item in $Items)
    {   
        if ($v = $item.Name | Select-String -AllMatches $RegEx | Select-Object -ExpandProperty Matches)
        {
            $item | Add-Member -MemberType NoteProperty -Name MatchedValue -Value $v.Value
            $shit = $shit + $item
        }
    }
    return $shit
}

#Yup function says it all
function ListTurds()
{
    foreach ($turd in $shit){Centeralize "$turd has an illegal character $($turd.MatchedValue)" "Red"}
}

#I'd like to say this function is more dynamic then it really is, it is not as dynamic as it sounds, but can be altered to suit such needs
function AskHowToList($Question, $color)
{
    Centeralize "$Question" $color;$answer = Read-Host;Write-Host " "
    Switch($answer)
    {
        List{$result=0}
        L{$result=0}
        Export{$result=1}
        E{$result=1}
        X{$result=1}
        Rename{$result=2}
        R{$result=2}
        default{AskHowToList $Question $color}
    }
    Switch ($result)
        {
              0 { Return "l" }
              1 { Return "x" }
              2 { Return "r" }
        }
}

#Clearly Export the results
function ExportList()
{
    $ValidPath = GetPath
    Write-Host "Provide a filename: " -NoNewLine;$answer = Read-Host; Write-Host ""
    $CSVFile = "$ValidPath\$answer"
    Centeralize "You have entered: $CSVFile ... But are you good to write?" "green"
    Try { [io.file]::OpenWrite($CSVFile).close() }
    Catch { Write-Warning "Unable to write to output file $CSVFile" }
    (Get-Item $CSVFile).Delete()
    Centeralize "Congrats, You can write to the specified file. Please Wait, creating export file..." "green"
    $shit | Select-Object Name, Directory, MatchedValue| Export-Csv -Path $CSVFile -NoTypeInformation

}

function ReplaceCharsInItems()
{
    #These chacaters cause -match to Error and need to be escaped
    #[ + ( )
    #These don't cause -match to error but list all? needs to be escaped
    #^ $
    #all the above need to be esaped with \ .... fuckers
    $Chars = $Chars -replace '[\\]' #\ is used as an escape character in RegEx, but we can't have file names with them, so it should not be in this loop
    $Fuckers = '[+($)^\[]' #The list of special characters in RegEx that need to be escaped
    for ($i=0; $i -lt $Chars.length; $i++)
    {
        Write-Host "You have selected: "$Chars[$i]" What do you want to replace this character with? " -NoNewline;$Replacement = Read-Host
        $SelectedCharacter = $Chars[$i]
        if($SelectedCharacter | Select-String -AllMatches $Fuckers){Write-Host "We found a special character... fixing...";$SelectedCharacter = "\"+$SelectedCharacter}
        foreach ($turd in $shit)
        {
            $MV = $turd.MatchedValue
            if ($MV -match $SelectedCharacter)
            {
                $OldName = $turd.name
                $NewName = $OldName.Replace($MV, $Replacement)
                Write-Host "Replacing $MV in "$turd.name" with $Replacement;" $NewName
            }
        }
    }
}

##########################################################################################################################
#   Main
##########################################################################################################################

foreach($L in $MylogoArray){Centeralize $L "green"}
Centeralize $ScriptName "White"
Centeralize "This script helps change file names in bulk for specific folders.`n"
$Path = GetPath
$Items = Get-ChildItem -Path $Path -Recurse
$Chars = GetChars
$RegEx = "[$Chars]"
$shit = FilterItems $Items $RegEx
if ($shit.Count -gt 0)
{
    $ugh = "We have found a total of "+ $shit.count +" turds... what ya want to do? ((L)ist\(R)ename\(E)xport)"
    switch(AskHowToList $ugh "Red")
    {
        l{ListTurds}
        x{ExportList}
        r{ReplaceCharsInItems}
    }
}
else
{
    Write-Host ""
    Centeralize "We found no files within your criteria. Whomp whom womomomomomo." "Yellow"
}
