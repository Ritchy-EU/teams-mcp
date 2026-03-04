---
description: Authenticate with Microsoft Teams. Use when the user wants to sign in to Microsoft Teams, when auth_status shows not authenticated, or when Teams tools fail with authentication errors.
---

# Microsoft Teams Authentication

Help the user sign in to Microsoft Teams using the device code flow.

## Steps

1. Check current authentication status using the `auth_status` tool
2. If already authenticated, inform the user — no further action needed
3. If not authenticated, call the `start_authentication` tool
4. The tool returns a URL and a one-time code — present them clearly to the user:
   - Ask the user to open the URL in their browser
   - Ask the user to enter the code shown on the page
5. Wait for the user to confirm they have completed the steps in the browser
6. Call `auth_status` again to verify authentication succeeded
7. Confirm to the user that Microsoft Teams is now connected

## Notes
- Authentication is saved locally — the user only needs to do this once
- If `start_authentication` fails, ask the user to check that AZURE_CLIENT_ID and AZURE_TENANT_ID are configured in Claude managed settings
