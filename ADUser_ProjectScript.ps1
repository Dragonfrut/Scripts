<# ORGANIZATIONAL UNITS
CreateOU is a function to check if top level OU Group1 - Group6 exsist, 
if they dont, to then create them. #>
function CreateOU(){

    <# For loop that loops 6 times and creates an OU with the name of Group and a number corisponding to how many
    times the loop has ran #>
    for ($i = 1; $i -le 6; $i++){
        $ouName = "Group" + "$i"
        
        # checks if the OU already exists using LDAP
        if([adsi]::Exists("LDAP://OU=$ouName,DC=morkrica,DC=com")){
            Write-Host "$ouName already exists" -ForegroundColor DarkYellow
        } else {
            # Creates new OU
            New-ADOrganizationalUnit -Name $ouName
            Write-Host "$ouName has been created" -ForegroundColor Green
        }
    }
}

# USERS
# Install-Module -Name ImportExcel #
function createUsers() {

    # tracker for how many times the loop is running
    $userNum = 1

    # import excel file and create headers to organize
    $filePath = Read-Host -Prompt 'Input the filepath for your Excel spreadsheet'
    $userList = Import-Excel -Path $filePath -HeaderName 'Firstname', 'Lastname'
    foreach ($user in $userList) {
        # collect data and make variables for acct info# 
        $firstName = $($user.'Firstname')
        $lastName = $($user.'Lastname')
        $fullName = $firstName + ' ' + $lastName
        # taking the first character of firstname and putting the lastname behind it, then converting all to lowercase
        $SAM = $($firstName.SubString(0,1) + $lastName).ToLower()
        $UPN = $SAM + '@morkrica.com'
        $password = 'Password01'
        $description = "Why was I, $fullName of $group, born from the likes of powershell? Into a world were I am constantly removed from existance and reborn into endless suffering?"
        $userNum++
        # query for the current user
        $exsistCheck = Get-ADUser -Filter "SamAccountName -eq '$SAM'"
        

        # decide which OU the user will belong in by how many times loop has ran
        if($userNum -le 50){$group = 'Group1'} `
        elseif($userNum -le 100 -and $userNum -gt 50) {$group = 'Group2'} `
        elseif($userNum -le 150 -and $userNum -gt 100) {$group = 'Group3'} `
        elseif($userNum -le 200 -and $userNum -gt 150) {$group = 'Group4'} `
        elseif($userNum -le 250 -and $userNum -gt 200) {$group = 'Group5'} `
        elseif($userNum -le 300 -and $userNum -gt 250) {$group = 'Group6'}


        # check if both first and last name exsits
        if(($firstName -ne $null) -and ($lastName -ne $null)) {

           # If the user already exists don't try to make them
           if($exsistCheck -eq $null) {

               # Creating a new user
               New-ADUser `
                -Name $fullName `
                -GivenName $firstName `
                -SurName $lastName `
                -SamAccountName $SAM `
                -Description $description `
                -UserPrincipalName $UPN `
                -AccountPassword $(ConvertTo-SecureString $password -AsPlainText -Force) `
                -ChangePasswordAtLogon $true `
                -Enabled $true `
                -Path "OU=$group,DC=morkrica,DC=com" 
                    

                Write-Host "Created $fullName in $group" -ForegroundColor Magenta

                # calls function to use the new user to make a Home Directory for it
                HomeDirectory($SAM)

           } else { Write-Host "User $SAM already exists" -ForegroundColor Yellow }
           
        } else { Write-Host 'The user you are trying to create does not have a full name available' -ForegroundColor Red }     
    }
}

# DIRECTORY
# function to create users home directories, set permissions, and "mount" the shares
function HomeDirectory($SAM){
    
    # variables to hold the shares path and pull the users information
    $homePath = "\\MRR-Powershell.morkrica.com\Home\{0}" -f $SAM
    $drive = 'U:'
    $user = Get-ADUser -Identity $SAM

    #if the user exists continue to create and set the home directory
    if($user -ne $null){

        # Creates the directory
        $homeDir = New-Item -Path $homePath -ItemType Directory -Force

        # gets the object that represents security for the given directory path
        $acl = Get-ACL $homeDir

        <# variable reprsenting user permissions, creating new permissions for the user on their new home
        directory #>
        $userPermissions = New-Object System.Security.AccessControl.FileSystemAccessRule("$SAM","FullControl","Allow")
        $acl.SetAccessRule($userPermissions)
        
        # applies the new permissions to the directory
        $acl | Set-Acl $homePath

        # sets the newly made home directory to the user 
        Set-ADUser -Identity $User -HomeDirectory "$homePath" -HomeDrive "$drive"
        
        Write-Host ("HomeDirecotry has been created at {0}" -f $homePath) -ForegroundColor DarkGreen
    
    }
}

# DELETE
# Function to delete everything that this script creates
function Delete(){
    
    <# Filters for all OU's with Group and * for anything, disables accidental deletion protection
    and deletes all of them and any users inside #>
    Get-ADOrganizationalUnit -filter "Name -like 'Group*'" | Set-ADObject -ProtectedFromAccidentalDeletion:$false -PassThru | Remove-ADOrganizationalUnit -Confirm:$false -Recursive
    
    # Removes all files in the 
    Remove-Item -Recurse -Force U:\Home\*
    Write-Host 'Removed all files in U:\Home\' -ForeGround Green
}

# asks if user wants to delete what the script made, then calls the delete function if they do.#
$delete = Read-Host -Prompt 'Do you want to remove everything this script makes? [y][n]'
if ($delete -eq 'y') {
    Delete
}
  
# asks if user wants to create the six OU's we need and calls the CreateOU function if yes .#
$makeOU = Read-Host -Prompt 'Do you want to try to make six group OUs? [y][n]'
if ($makeOU -eq 'y') {
    CreateOU
}

# asks if user wants to create users from an excel file, and calls CreateUsers if yes.# 
$makeUser = Read-Host -Prompt 'Do you want to add AD users with an excel file [y][n]'
if ($makeUser -eq 'y') {
    CreateUsers
}
 