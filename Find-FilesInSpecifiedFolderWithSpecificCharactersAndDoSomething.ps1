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
$ScriptName = "Find-FilesInSpecifiedFolderWithSpecificCharactersAndDoSomething ¯\_(°_o)_/¯`n"
$ScriptDescription = "This script helps change file names in bulk for specific folders.`n"

$pswheight = (get-host).UI.RawUI.MaxWindowSize.Height
$pswwidth = (get-host).UI.RawUI.MaxWindowSize.Width

##########################################################################################################################
#   Functions
##########################################################################################################################

#function takes in a name to alert confirmation of deletion returns true or false
function confirm($OldName, $NewName)
{
    #function variables, generally only the first two need changing
    $title = "Confirm File Rename!"
    $message = "You are about to rename the file: $OldName to $NewName"

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "This means Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "This means No"

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $Options, 0)

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
  [Parameter(Position=1,Mandatory=$false)]
  [string]$C,
  [Parameter(Position=2,Mandatory=$false)]
  [string]$N
  )
    $sLength = $S.Length
    $padamt =  "{0:N0}" -f (($pswwidth-$sLength)/2)
    $PadNum = $padamt/1 + $sLength #the divide by one is a quick dirty trick to covert string to int
    $CS = $S.PadLeft($PadNum," ").PadRight($PadNum," ") #Pad that shit
    if ($C -and $N) #if variable for color exists run below
    {    
        Write-Host $CS -ForegroundColor $C -NoNewline #write that shit to host with color
    }
    elseif ($C)
    {
        Write-Host $CS -ForegroundColor $C
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
    Write-host "Characters to search for !$&'``^,~#%)(}{][@=; " -ForegroundColor Magenta -NoNewline
    $UnsupportedChars = Read-Host
    if ( !$UnsupportedChars ){ $UnsupportedChars = "\]!$'&{^,+}~#%[)@=;(``" }
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
            $item | Add-Member -MemberType NoteProperty -Name NameToChange -Value $item.name
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
    Centeralize "$Question " -C $color -N "NoNewLine";$answer = Read-Host;Write-Host ""
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
    Centeralize "You have entered: $CSVFile ... But are you good to write?" "yellow"
    Try { [io.file]::OpenWrite($CSVFile).close() }
    Catch { Centeralize "Unable to write to output file $CSVFile" "Red"}
    (Get-Item $CSVFile).Delete()
    Centeralize "Congrats, You can write to the specified file. Please Wait, creating export file..." "green"
    $shit | Select-Object Name, Directory, @{Expression={$_.MatchedValue -join '\'}} | Export-Csv -Path $CSVFile -NoTypeInformation
}

function ReplaceCharsInItems()
{
    #These chacaters cause -match to Error and need to be escaped
    #[ + ( )
    #These don't cause -match to error but list all? needs to be escaped
    #^ $
    #all the above need to be esaped with \ .... SpecialCharacters
    $Chars = $Chars -replace '[\\]' #\ is used as an escape character in RegEx, but we can't have file names with them, so it should not be in this loop
    $SpecialCharacters = '[+($)^\[]' #The list of special characters in RegEx that need to be escaped
    for ($i=0; $i -lt $Chars.length; $i++)
    {
        Write-Host "You have selected: "$Chars[$i]" What do you want to replace this character with? " -NoNewline;$Replacement = Read-Host
        $SelectedCharacter = $Chars[$i]
        if($SelectedCharacter | Select-String -AllMatches $SpecialCharacters){$SelectedCharacter = "\"+$SelectedCharacter} #;Write-Host "We found a special character... fixing..."
        foreach ($turd in $shit)
        {
            if ($turd.MatchedValue -match $SelectedCharacter)
            {
                $OldName = $turd.NameToChange
                $NewName = $OldName -Replace $SelectedCharacter, $Replacement
                $turd.NameToChange=$NewName
            }
        }
    }
    Write-Host ""
    switch(AskHowToRename "Would you like to manually go through each file, or bulk rename with report? (M)anual\(B)ulk" "Yellow")
    {
        m {Manual}
        a {Auto}
    }

}

function AskHowToRename($Question, $color)
{
    Centeralize "$Question " -C $color -N "NoNewLine";$answer = Read-Host;Write-Host ""
    Switch($answer)
    {
        Manual{$result=0}
        M{$result=0}
        Bulk{$result=1}
        B{$result=1}
        A{$result=1}
        default{AskHowToRename $Question $color}
    }
    Switch ($result)
        {
              0 { Return "m" }
              1 { Return "a" }
        }
}

function Manual()
{
    foreach($turd in $shit)
    {
        #$turd | Get-Member
        if(confirm $turd.name $turd.NameToChange)
        {
            
            $ThatOldName = $turd.FullName
            if (($ThatOldName).ToCharArray() -ccontains "[") {Write-host "We found a bad char";$ThatOldName -replace "\[", "`["}
            Try { Rename-Item -literalPath $ThatOldName -NewName ($turd.NameToChange) -ErrorAction Stop }
            Catch [System.UnauthorizedAccessException]
            {
                Centeralize "You do not have permission to change this file's name" "Red"
            }
            Catch 
            { 
                $ErMsg = "Unable to change file's name "+$turd.name
                Centeralize $ErMsg "Red"
                $ErrMsg = $_.Exception.Message
                Centeralize $ErrMsg "Red"
            }
        }
        else
        {
            Write-Host "Maybe Another time..."
        }
    }
}

function confirmAuto()
{
    #function variables, generally only the first two need changing
    $title = "Confirm Automatic File Rename!"
    $message = "You are about to rename all the files! Make sure this is what you want to do!"

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "This means Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "This means No"

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $Options, 0)

    Switch ($result)
        {
              0 { Return $true }
              1 { Return $false }
        }
}

function Auto()
{
    Centeralize "Provide a report output path."
    $ValidPath = GetPath
    Write-Host "Provide a filename: " -NoNewLine;$answer = Read-Host; Write-Host ""
    $CSVFile = "$ValidPath\$answer"
    $shit | Select-Object Name, Directory, @{N='New File Name';Expression={$_.NameToChange}},@{N='Matched Value';Expression={$_.MatchedValue -join '\'}} | Export-Csv -Path $CSVFile -NoTypeInformation
    Centeralize "A report has been generated at $CSVFile if everything is ok, then continue. This is your final warning." "Red"
    if(confirmAuto)
    {
        Write-Host "This is where we write the final rename code..."
            foreach($turd in $shit)
            {
            Try { Rename-Item $turd.name -NewName ($turd.NameToChange) -ErrorAction Stop }
            Catch [System.UnauthorizedAccessException]
            {
                Centeralize "You do not have permission to change this file's name" "Red"
            }
            Catch 
            { 
                $ErMsg = "Unable to change file's name "+$turd.name
                Centeralize $ErMsg "Red"
                $ErrMsg = $_.Exception.Message
                Centeralize $ErrMsg "Red"
            }
        }
    }
}

##########################################################################################################################
#   Main
##########################################################################################################################

foreach($L in $MylogoArray){Centeralize $L "green"}
Centeralize $ScriptName "White"
Centeralize $ScriptDescription "Green"
$Path = GetPath
$Items = Get-ChildItem -Path $Path -Recurse
$Chars = GetChars
$RegEx = "[$Chars]"
$shit = FilterItems $Items $RegEx
if ($shit.Count -gt 0)
{
    Write-Host ""
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
