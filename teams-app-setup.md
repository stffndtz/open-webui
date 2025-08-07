# Teams App Setup for ai.nordholding.de

## Overview
This guide explains how to set up the Microsoft Teams app to work with the Open WebUI instance at `ai.nordholding.de`.

## Prerequisites
1. Access to Microsoft Teams admin center
2. Azure AD application registration
3. Teams app manifest file

## Steps to Deploy

### 1. Azure AD App Registration
1. Go to Azure Portal > Azure Active Directory > App registrations
2. Create a new registration:
   - Name: "Open WebUI Teams App"
   - Supported account types: "Accounts in this organizational directory only"
   - Redirect URI: Web > `https://ai.nordholding.de/auth/callback`

### 2. Configure Authentication
1. In the app registration, go to "Authentication"
2. Add platform: "Single-page application"
3. Add redirect URI: `https://ai.nordholding.de`
4. Enable "Access tokens" and "ID tokens"

### 3. API Permissions
1. Go to "API permissions"
2. Add permissions:
   - Microsoft Graph > Delegated > User.Read
   - Microsoft Graph > Delegated > User.ReadBasic.All

### 4. Create Teams App Package
1. Update the `teams-app-manifest.json` file:
   - Replace `{{TEAMS_APP_ID}}` with your app registration ID
   - Replace `{{AAD_APP_ID}}` with your app registration ID
2. Create app icons (color.png and outline.png)
3. Package the manifest and icons into a .zip file

### 5. Upload to Teams
1. Go to Teams admin center
2. Navigate to Teams apps > Manage apps
3. Upload the app package
4. Assign the app to users or groups

## Troubleshooting

### "URL was unable to be opened in embedded browser" Error
This error occurs when:
1. The domain `ai.nordholding.de` is not in the `validDomains` list
2. The Teams app doesn't have proper permissions
3. The authentication flow is trying to redirect to an unauthorized URL

**Solutions:**
1. Ensure `ai.nordholding.de` is listed in `validDomains` in the manifest
2. Check that the Azure AD app has the correct redirect URIs
3. Verify the app has proper API permissions
4. Test the authentication flow in Teams developer portal

### Authentication Flow Issues
1. Check browser console for detailed error messages
2. Verify the Teams SDK is properly initialized
3. Ensure the app is running in a Teams iframe environment
4. Test with different authentication prompts (none, consent, login)

## Configuration Files

### teams-app-manifest.json
The manifest file contains:
- App metadata (name, description, icons)
- Tab configurations
- Authentication settings
- Valid domains list
- Required permissions

### Environment Variables
Set these in your deployment:
```
TEAMS_APP_ID=your-app-registration-id
AAD_APP_ID=your-app-registration-id
WEBUI_BASE_URL=https://ai.nordholding.de
```

## Testing
1. Install the app in Teams
2. Navigate to the app tab
3. Check browser console for authentication logs
4. Verify user context is available
5. Test the chat functionality

## Support
For issues with Teams integration, check:
1. Teams developer documentation
2. Azure AD app registration logs
3. Browser console for JavaScript errors
4. Network tab for failed requests 