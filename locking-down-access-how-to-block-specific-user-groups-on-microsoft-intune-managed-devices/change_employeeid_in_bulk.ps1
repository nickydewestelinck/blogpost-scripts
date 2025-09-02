#Step 1: Check if the Microsoft Graph Module is installed; if not, install it
 
$graphModule = Get-Module -ListAvailable Microsoft.Graph
if ($graphModule) {
    Write-Host "The Microsoft Graph PowerShell Module is already installed." -ForegroundColor Green
} else {
    Write-Host "The Microsoft Graph PowerShell Module is NOT installed." -ForegroundColor Red
    Write-Host "Installing now..." -ForegroundColor Yellow
 
        try {
            Install-Module Microsoft.Graph -Scope CurrentUser -Force
            Write-Host "The Microsoft Graph PowerShell Module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install the Microsoft Graph PowerShell Module: $_" -ForegroundColor Red
        }
    }
 
#Step 2: Connect to Microsoft Graph and update employeeId in bulk from CSV
##Step 2.1: Connect to Microsoft Graph with the required scopes
Connect-MgGraph -Scopes "User.ReadWrite.All"
 
##Step 2.2: Confirm connection to Microsoft Graph
$context = Get-MgContext
Write-Host "Connected as $($context.Account)" -ForegroundColor Cyan
 
##Step 2.3: Import CSV
$csvPath = "Path To Your .csv File"
$userList = Import-Csv -Path $csvPath
 
##Step 2.4: Loop through users and update employeeId for all users in the CSV
foreach ($user in $userList) {
    $upn = $user.UserPrincipalName
    $employeeId = $user.EmployeeID
 
    try {
        Update-MgUser -UserId $upn -EmployeeId $employeeId
        Write-Host "Updated $upn with EmployeeID $employeeId" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to update $upn $_" -ForegroundColor Red
    }
}
 
#Step 3: Disconnect from Microsoft Graph
Disconnect-MgGraph
