# Load the Active Directory module (if not already loaded)
if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host "Active Directory module loaded successfully."
    } catch {
        Write-Error "Failed to load Active Directory module. Ensure RSAT is installed." 
        exit
    }
}

# Define the path to the CSV file (Update this with the actual file path)
$CSVPath = "C:\Path\To\Your\File.csv"

# Define the domain controller (Update this with your domain controller FQDN)
$DC = "YourDomainController.yourdomain.com"

# Validate if the CSV file exists
if (-Not (Test-Path -Path $CSVPath)) {
    Write-Error "CSV file not found at: $CSVPath. Ensure the file path is correct."
    exit
}

# Import users from CSV file
try {
    $Users = Import-Csv -Path $CSVPath
    if ($Users.Count -eq 0) {
        Write-Error "CSV file is empty. Please ensure it contains valid data."
        exit
    }
} catch {
    Write-Error "Failed to read CSV file. Ensure the file is properly formatted."
    exit
}

# CSV Expected Fields:
# - Email (UserPrincipalName in AD)
# - EmployeeID (New Employee ID to assign)

# Iterate through each user in the CSV
foreach ($user in $Users) {
    # Validate required fields
    if (-not $user.Email -or -not $user.EmployeeID) {
        Write-Warning "Skipping row due to missing Email or EmployeeID."
        continue
    }

    # Store UPN from CSV (UserPrincipalName in Active Directory)
    $UPN = $user.Email

    Write-Host "`nProcessing UPN: $UPN"

    # Define AD Query Filter
    $Filter = "userPrincipalName -eq '$UPN'"

    try {
        # Retrieve AD user object
        $Result = Get-ADUser -Filter $Filter -Properties DistinguishedName -Server $DC -ErrorAction Stop
        $Count = @($Result).Count

        if ($Count -eq 1) {
            Write-Host "Found 1 AD user for UPN: $UPN"

            # Update employeeID attribute
            Set-ADUser -Identity $Result.DistinguishedName -Replace @{ employeeID = $user.EmployeeID } -Server $DC
            Write-Host "Successfully updated EmployeeID: $($user.EmployeeID) for user: $UPN"

        } elseif ($Count -gt 1) {
            Write-Warning "Multiple AD accounts found for UPN '$UPN'. Skipping."
        } else {
            Write-Warning "No AD user found with UPN '$UPN'. Skipping."
        }

    } catch {
        Write-Warning "Error processing UPN '$UPN': $_"
    }
}

Write-Host "`nScript execution completed."
