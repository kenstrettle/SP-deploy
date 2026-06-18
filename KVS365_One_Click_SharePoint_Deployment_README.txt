KVS365 One-Click SharePoint Deployment Script
===========================================

What this script does
---------------------
1. Connects to the tenant admin site using PnP.PowerShell.
2. Creates a modern Team Site if it does not already exist.
3. Applies a simple site sharing baseline.
4. Creates standard SharePoint groups.
5. Optionally creates separate document libraries for Company Admin, Finance, HR, Sales & Marketing, Operations, Projects and Archive.
6. Applies library versioning.
7. Optionally breaks library inheritance and assigns a SharePoint group to that library.
8. If you switch off extra libraries, it builds a folder-based structure inside Shared Documents.

Before you run it
-----------------
- Update the CONFIG block in the .ps1 file.
- Check the tenant name, owner account, site title and alias.
- Decide whether you want separate libraries or one Shared Documents library with folders.
- Add any initial members to the $GroupMembers hashtable.

Recommended first use
---------------------
- Run it in a test tenant first.
- Start with CreateExtraLibraries = $false if you want a simpler first deployment.
- Once you are comfortable, use separate libraries for Finance and HR where unique permissions make sense.

Suggested KVS365 defaults
-------------------------
- Finance and HR: separate libraries with unique permissions.
- Everything else: broad member access.
- Archive: read-only where practical.
- Keep folder depth shallow.

Authentication note
-------------------
PnP PowerShell guidance notes that the old multi-tenant PnP Management Shell app was deleted in September 2024, and environments that depended on it may need their own Entra ID app registration for authentication.
