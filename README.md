# PowerShell Script for Managing GPO and Retrieving Services

This PowerShell script is designed to pull a list of services froma remote server, display the list in a selectable GUI, then create a Group Policy Object (GPO) that gives the specified AD group start/stop/restart permissions to the service(s). Optionally you can link it to the closest Organizational Unit (OU) of the specified server during the GPO creation

## IMPORTANT

The script sets the permissions to the specified service by writing the SDDL (Security Descriptor Definition Language) directly into the GPO .inf file. The default permissions created by this GPO are as follows:
  - System : Full Control
  - Builtin Administrators: Full Control
  - Interactive: Read
  - "AD Group": Start/stop/pause

The service(s) are also set to start automatically.
In order to change the permissions given by the script, you can either edit the text at line 129, or create the GPO and edit the permissions afterwards. Please test thoroughly after creating and linking the GPO, though removing it should set the permissions back to the default values

## Parameters

- **`$gpoName`** (Mandatory): The name of the GPO to be created.
- **`$serverName`** (Mandatory): The name of the computer from which services will be gathered.
- **`$AdGroup`** (Mandatory): The name of the Active Directory group to which access should be granted.
- **`$domainName`** (Optional): The Fully Qualified Domain Name (FQDN) of the domain (e.g., domain.local). If not provided, the domain will be retrieved from the running computer.
- **`$AutoLinkGPO`** (Switch): If specified, the new policy will be linked to the closest OU of the server.

## Usage

1. **Set Parameters:**
   - Provide the necessary parameters such as `gpoName`, `serverName`, and `AdGroup`.
   - Optionally, provide the `domainName` or let the script retrieve it automatically.

2. **Run the Script:**
   - Execute the script to create the GPO, link it to the closest OU, and retrieve the list of services from the specified server.

3. **View Services:**
   - The script will display a GUI with a list of services from the remote computer.

4. **Requirements:**
- PowerShell 5.1 or later
- Active Directory module
- Administrative privileges

5. **Notes**
- Ensure that the script is run with appropriate permissions to create GPOs and retrieve services from the remote computer.
The GUI is created using Windows Forms, so it requires a Windows environment to run.

## License
This script is provided “as-is” without any warranty. Use at your own risk.
## Example

```powershell
.\New-ServicesGPO.ps1 -gpoName "MyGPO" -serverName "Server01" -AdGroup "MyAdGroup" -AutoLinkGPO
```
The above command will create a GPO in Active Directory called "MyGPO", connected to a remote server called "Server01" and will give the group "MyAdGroup" stop/start/pause permissions to the specified services.
After the GPO is created it will link the GPO to the OU of the server.

```powershell
.\New-ServicesGPO.ps1 -gpoName "MyGPO2" -serverName "Server02" -AdGroup "MyAdGroup2"
```
The above command will create a GPO in Active Directory called "MyGPO2", connected to a remote server called "Server02" and will give the group "MyAdGroup2" stop/start/pause permissions to the specified services.
The GPO will need to be manually linked to the required OU(s) after creation.
