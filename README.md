# Defender for Office 365 API Functions App Connector

This Functions App Connector provides HTTP-triggered endpoints to manage Microsoft Defender for Office 365 and Exchange Online security configurations using Managed Identity authentication.

## Technology Stack

* **Runtime**: Azure Functions v4 (PowerShell)
* **PowerShell Module**: ExchangeOnlineManagement v3.x (automatically managed)
* **Extension Bundle**: Microsoft.Azure.Functions.ExtensionBundle v4.x
* **Authentication**: Azure Managed Identity (System-Assigned or User-Assigned)

## Prerequisites

* Azure subscription with permissions to create Function Apps and assign Azure AD roles
* Microsoft 365 tenant with Exchange Online
* Global Administrator or Privileged Role Administrator access (for role assignment)
* PowerShell with Microsoft.Graph module (for post-deployment configuration)

## Actions Supported

| **Component** | **Description** |
| --------- | -------------- |
| **ConnectExchangeOnline** | Establish authenticated session to Exchange Online using Managed Identity |
| **DisconnectExchangeOnline** | Terminate the Exchange Online session |
| **ListSpamPolicy** | View existing spam filter policies |
| **CreateSpamPolicy** | Create a spam filter policy |
| **CreateSpamRule** | Create a spam filter rule |
| **TenantAllowBlockList** | View entries for email addresses in the Tenant Allow/Block List |
| **CreateAllowBlockList** | Create block entries for email addresses in the Tenant Allow/Block List |
| **UpdateAllowBlockList** | Modify entries for email addresses in the Tenant Allow/Block List |
| **RemoveAllowBlockListItems** | Remove entries for email addresses from the Tenant Allow/Block List |
| **ListMalwarePolicy** | View existing malware filter policies |
| **BlockMalwareFileExtension** | Add file extensions to malware policy block list |
| **GetInboxRule** | View list of existing rules created in a mailbox (supports IncludeHidden parameter) |
| **SetInboxRule** | Enable or disable inbox rules for a mailbox |
| **RemoveInboxRule** | Remove inbox rule from a particular mailbox |

---

## Deployment Instructions

### Step 1: Deploy the ARM Template

You can deploy this Function App using the Azure Portal:

1. Download the `azuredeploy.json` file from this repository
2. Log in to the [Azure Portal](https://portal.azure.com)
3. Search for **"Deploy a custom template"** in the top search bar
4. Click **"Build your own template in the editor"**
5. Click **"Load file"** and select the downloaded `azuredeploy.json`
6. Click **"Save"**
7. Fill in the deployment parameters:
   - **Subscription**: Select your Azure subscription
   - **Resource Group**: Create new or select existing
   - **Region**: Choose your deployment region
   - **Function App Name**: Enter a unique name (e.g., `o365defender`)
   - **Organization Name**: Your Microsoft 365 organization (e.g., `contoso.onmicrosoft.com`)
   - **User Assigned Identity Resource Id** (Optional): Leave empty for System-Assigned MSI only
8. Click **"Review + create"** → **"Create"**

The deployment will create:
- Azure Function App (PowerShell runtime)
- Storage Account
- App Service Plan (Consumption tier)
- Application Insights
- System-Assigned Managed Identity (enabled by default)

**Important**: After deployment completes, note down the **Outputs** section values:
- `functionAppName`: The generated function app name
- `systemAssignedIdentityPrincipalId`: The Principal ID for role assignment
- `functionAppUrl`: The function app endpoint URL

---

### Step 2: Deploy Function Code

After the ARM template deployment completes, upload the function code:

#### Option A: Using Azure Portal (Easiest)

1. Download `O365DefenderFunctionApp.zip` from the repository releases
2. Go to your Function App in Azure Portal
3. Click **"Deployment Center"** in the left menu
4. Select **"Local Git"** or **"External Git"** → Or use the simpler method below:
5. Alternative: Go to **"Advanced Tools (Kudu)"** → Click **"Go"**
6. In Kudu, go to **"Tools"** → **"Zip Push Deploy"**
7. Drag and drop `O365DefenderFunctionApp.zip` into the upload area
8. Wait for extraction to complete (~30 seconds)
9. Verify: Go back to Function App → **"Functions"** → You should see all 14 functions listed

#### Option B: Using Azure CLI

```bash
# First, create the zip file
cd O365DefenderFunctionApp
zip -r ../O365DefenderFunctionApp.zip .
cd ..

# Deploy to Function App
az functionapp deployment source config-zip \
  --resource-group <your-resource-group> \
  --name <your-function-app-name> \
  --src O365DefenderFunctionApp.zip
```

#### Option C: Using Azure Functions Core Tools

```bash
# Install Azure Functions Core Tools
# Then publish directly:
cd O365DefenderFunctionApp
func azure functionapp publish <your-function-app-name>
```

---

## Post-Deployment Configuration

### Step 1: Assign Exchange Online Role to Managed Identity

The Managed Identity needs an Exchange Online administrator role to execute cmdlets.

#### Option A: Using Azure Portal

1. Go to **Azure Active Directory** (or **Microsoft Entra ID**)
2. Click **Roles and administrators**
3. Search for and select **Exchange Administrator**
4. Click **+ Add assignments**
5. Search for your Function App name (e.g., `o365defxxxxxxx`)
6. Select it and click **Add**

#### Option B: Using PowerShell

```powershell
# Install Microsoft Graph module if not already installed
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory"

# Get the Managed Identity Principal ID (from deployment outputs)
$principalId = "PASTE_YOUR_PRINCIPAL_ID_HERE"

# Get Exchange Administrator role
$roleDefinitionId = (Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq 'Exchange Administrator'").Id

# Assign role
New-MgRoleManagementDirectoryRoleAssignment -PrincipalId $principalId -RoleDefinitionId $roleDefinitionId -DirectoryScopeId "/"

Write-Host "Exchange Administrator role assigned successfully!"
```

**Alternative Roles** (choose based on your needs):
- **Global Administrator**: Full access (not recommended for production)
- **Security Administrator**: Security-focused operations
- **Compliance Administrator**: Compliance and policy management

---

### Step 2: Grant API Permissions

The Managed Identity requires the `Exchange.ManageAsApp` API permission. Managed Identities cannot use the Azure Portal UI for permissions, so use PowerShell:

```powershell
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"

# Get the Managed Identity Principal ID (from deployment outputs or Step 1)
$principalId = "PASTE_YOUR_PRINCIPAL_ID_HERE"

# Get the Managed Identity service principal
$msi = Get-MgServicePrincipal -Filter "id eq '$principalId'"

# Get Office 365 Exchange Online service principal
$exchangeOnline = Get-MgServicePrincipal -Filter "displayName eq 'Office 365 Exchange Online'"

# Get the Exchange.ManageAsApp role
$appRole = $exchangeOnline.AppRoles | Where-Object { $_.Value -eq "Exchange.ManageAsApp" }

# Grant the permission
New-MgServicePrincipalAppRoleAssignment `
    -ServicePrincipalId $msi.Id `
    -PrincipalId $msi.Id `
    -ResourceId $exchangeOnline.Id `
    -AppRoleId $appRole.Id

Write-Host "Exchange.ManageAsApp permission granted successfully!"
```

**Important**: This permission grants admin-level access to Exchange Online. Wait 5-15 minutes for propagation before testing.

---

### Step 3: Wait for Role Propagation

Azure AD role assignments can take **5-15 minutes** to propagate. Wait before testing the connection.

---

### Step 4: Test the Connection

Use the Azure Portal or a REST client to test:

```bash
# Get the function key from Azure Portal: Function App → Functions → ConnectExchangeOnline → Function Keys

curl -X POST "https://<your-function-app>.azurewebsites.net/api/ConnectExchangeOnline?code=<function-key>" \
  -H "Content-Type: application/json" \
  -d '{}'
```

Expected response (HTTP 200):
```json
"Exchange Online connected successfully using Managed Identity"
```

---

## Optional: User-Assigned Managed Identity Setup

If you need to use a User-Assigned Managed Identity:

### 1. Create User-Assigned Managed Identity

```bash
az identity create --name o365-defender-identity --resource-group <your-rg>
```

Note the output:
- `id`: Full resource ID (needed for ARM template parameter)
- `clientId`: Client ID (needed for API requests)

### 2. Assign Roles

Follow the same role assignment steps as above, but use the User-Assigned Identity's Principal ID.

### 3. Attach to Function App

Provide the `id` value as the `UserAssignedIdentityResourceId` parameter during ARM template deployment.

### 4. Use in API Requests

Include the `clientId` in your ConnectExchangeOnline requests:

```json
{
    "ManagedIdentityAccountId": "<client-id-from-step-1>"
}
```

---

## How to Use

### 1. Connect to Exchange Online

**Always call this first** before executing any Exchange Online operations.

**System-Assigned MSI** (simplest - uses environment variable):
```bash
POST /api/ConnectExchangeOnline
Content-Type: application/json

{}
```

**System-Assigned MSI** (with organization override):
```bash
POST /api/ConnectExchangeOnline
Content-Type: application/json

{
    "OrganizationName": "contoso.onmicrosoft.com"
}
```

**User-Assigned MSI**:
```bash
POST /api/ConnectExchangeOnline
Content-Type: application/json

{
    "ManagedIdentityAccountId": "bf6dcc76-4331-4942-8d50-87ea41d6e8a1"
}
```

### 2. Execute Operations

Call any other function endpoints to manage Exchange Online.

### 3. Disconnect from Exchange Online

**Always call this when finished** to clean up the session.

```bash
POST /api/DisconnectExchangeOnline
Content-Type: application/json

{}
```

---

## API Request Body Examples

### 1. ConnectExchangeOnline

**System-Assigned MSI** (uses environment variable):
```json
{}
```

**User-Assigned MSI**:
```json
{
    "ManagedIdentityAccountId": "bf6dcc76-4331-4942-8d50-87ea41d6e8a1"
}
```

### 2. DisconnectExchangeOnline

```json
{}
```

### 3. ListSpamPolicy

```json
{
    "Identity": ""
}
```
*Leave Identity empty to fetch all policies*

### 4. CreateSpamPolicy

```json
{
    "Name": "BlockSuspiciousSenders",
    "HighConfidenceSpamAction": "Quarantine",
    "SpamAction": "Quarantine",
    "BulkThreshold": 6,
    "BlockedSenderDomains": ["suspicious-domain.com", "spam-sender.net"]
}
```

### 5. CreateSpamRule

```json
{
    "Name": "ApplyToSalesDept",
    "HostedContentFilterPolicy": "BlockSuspiciousSenders",
    "RecipientDomainIs": "contoso.com"
}
```

### 6. TenantAllowBlockList

```json
{
    "ListType": "Sender"
}
```

### 7. CreateAllowBlockList

Block sender email addresses, URLs, or file hashes:

**Block email addresses:**
```json
{
    "ListType": "Sender",
    "Entries": ["malicious@attacker.com", "spam@bad-domain.com"]
}
```

**Block URLs** (no protocol, no trailing slash):
```json
{
    "ListType": "Url",
    "Entries": ["malicious-site.com", "phishing-page.net/login"]
}
```

**Block file hashes** (SHA256):
```json
{
    "ListType": "FileHash",
    "Entries": ["9f86d081884c7d659a2feaa0c55ad015a3bf4f1b2b0b822cd15d6c15b0f00a08"]
}
```

*Optional: Add `"ExpirationDate": "2025-12-31"` for temporary blocks (format: YYYY-MM-DD)*

### 8. UpdateAllowBlockList

```json
{
    "ListType": "Sender",
    "Entries": "updated@domain.com",
    "ExpirationDate": "12/31/2025"
}
```

### 9. RemoveAllowBlockListItems

```json
{
    "ListType": "Sender",
    "Entries": "remove@domain.com"
}
```

### 10. ListMalwarePolicy

```json
{
    "Identity": ""
}
```

### 11. BlockMalwareFileExtension

```json
{
    "MalwarePolicyName": "Default",
    "FileExtensions": [".exe", ".vbs", ".js"]
}
```

### 12. GetInboxRule

**Get visible rules only:**
```json
{
    "Mailbox": "user@contoso.com"
}
```

**Get all rules (including hidden):**
```json
{
    "Mailbox": "user@contoso.com",
    "IncludeHidden": true
}
```

### 13. SetInboxRule

**Disable a specific rule:**
```json
{
    "Mailbox": "user@contoso.com",
    "Identity": "Suspicious Rule Name",
    "Enabled": false
}
```

**Enable a specific rule:**
```json
{
    "Mailbox": "user@contoso.com",
    "Identity": "Legitimate Rule Name",
    "Enabled": true
}
```

**Disable ALL rules for a mailbox:**
```json
{
    "Mailbox": "user@contoso.com",
    "DisableAll": true
}
```

### 14. RemoveInboxRule

```json
{
    "Mailbox": "user@contoso.com",
    "Identity": "RuleName or RuleIdentity"
}
```

---

## Getting the Deployment Package

You have two options to get the function code:

### Option 1: Download from Repository

Download the pre-built `O365DefenderFunctionApp.zip` from the GitHub repository releases page.

### Option 2: Create Your Own

If you've made modifications, create the zip file:

**Windows (PowerShell)**:
```powershell
Compress-Archive -Path .\O365DefenderFunctionApp\* -DestinationPath O365DefenderFunctionApp.zip -Force
```

**Linux/macOS/WSL**:
```bash
cd O365DefenderFunctionApp
zip -r ../O365DefenderFunctionApp.zip .
cd ..
```

**Important**: The zip must contain the function app files at the root level (not in a subfolder):
```
O365DefenderFunctionApp.zip
├── host.json                    ← At root level!
├── profile.ps1
├── requirements.psd1
├── ConnectExchangeOnline/
│   ├── function.json
│   └── run.ps1
└── (other function folders...)
```

**Verify structure**:
```bash
# Linux/WSL
unzip -l O365DefenderFunctionApp.zip | head -20

# PowerShell
Expand-Archive O365DefenderFunctionApp.zip -DestinationPath temp -Force; ls temp; rm -r temp
```

---

## Troubleshooting

### Connection Fails with "Access Denied"

**Cause**: Managed Identity doesn't have Exchange Administrator role or API permission.

**Solution**:
1. Verify role assignment in Azure AD → Roles and administrators
2. Check API permissions in Azure AD → App registrations
3. Wait 5-15 minutes for propagation

### "OrganizationName is required" Error

**Cause**: Organization name not provided and environment variable not set.

**Solution**:
- Provide `OrganizationName` in request body, OR
- Ensure ARM template was deployed with `OrganizationName` parameter

### Functions Not Found (404)

**Cause**: Deployment package URL is incorrect or zip structure is wrong.

**Solution**:
1. Verify `WEBSITE_RUN_FROM_PACKAGE` URL in Function App settings
2. Download the zip manually to verify structure
3. Check that `host.json` is at the root of the zip

### Module Version Errors

**Cause**: Old ExchangeOnlineManagement module version.

**Solution**:
- Restart the Function App to trigger cold start
- Check `requirements.psd1` shows `'ExchangeOnlineManagement' = '3.*'`
- Note: PowerShell managed dependencies only support major version wildcards (e.g., `3.*`) not minor versions (e.g., `3.9.*`)
- Clear Function App cache: Platform features → Advanced Tools → Restart

### Extension Bundle Version Errors

**Cause**: Using deprecated extension bundle v3.

**Solution**:
- Check `host.json` shows `"version": "[4.*, 5.0.0)"`
- Extension bundle v3 is deprecated; v4 is the current stable version
- Rebuild and redeploy if using old package

### URL Blocking Not Working

**Cause**: Invalid URL format in CreateAllowBlockList entries.

**Solution**:
- **Remove** protocol (`https://` or `http://`)
- **Remove** trailing slashes
- **Valid formats**:
  - `malicious-site.com` - Blocks entire domain
  - `www.malicious-site.com` - Blocks specific subdomain
  - `malicious-site.com/path` - Blocks specific path
- **Invalid formats**:
  - ❌ `https://malicious-site.com`
  - ❌ `malicious-site.com/`
  - ❌ `http://www.malicious-site.com/page.html`

---

## Security Considerations

* Function-level authentication is enabled (requires function key in URL)
* Managed Identity credentials are managed automatically by Azure
* No secrets or certificates need to be stored or rotated
* Use Azure Key Vault for any additional sensitive configuration
* Enable Application Insights for monitoring and auditing
* Consider enabling Azure Private Link for enhanced network security

---

## References

* [Exchange Online PowerShell Managed Identity Documentation](https://learn.microsoft.com/en-us/powershell/exchange/connect-exo-powershell-managed-identity)
* [Defender for Office 365 Documentation](https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/anti-spam-policies-configure)
* [Azure Functions PowerShell Developer Guide](https://learn.microsoft.com/en-us/azure/azure-functions/functions-reference-powershell)
* [Azure Managed Identities Overview](https://learn.microsoft.com/en-us/azure/active-directory/managed-identities-azure-resources/overview)
