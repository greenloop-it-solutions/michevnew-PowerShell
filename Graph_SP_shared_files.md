# SharePoint Sites shared items inventory

The script uses the Graph API to enumerate all SharePoint sites in the tenant and enumerates all Document Libraries as drives. All items in the drive are enumerated and if a file is shared, additional information about the permissions is gathered. Running the script without any parameters will include only the items in the "root" directory, if you want to include items in any of the (sub)folders, use the corresponding script parameters as detailed below.

In order to use the script, you will need to first configure some variables and store them in a `secrets.json` file in the same directory. That file should be formatted like:

```json
{
    "Sharepoint-ODFB-tenantID": "tenant.onmicrosoft.com",
    "Sharepoint-ODFB-appID" : "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxx",
    "Sharepoint-ODFB-client_secret" : "xxxxxxlongsecretstringxxxxxx"
}
```
Provide the tenantID and the AppID of an Azure AD application you've registered with the tenant. The application needs the following permissions for the script to run as expected:

\#    User.Read.All to enumerate all users in the tenant

\#    Sites.ReadWrite.All to return all the item sharing details

\#    Files.Read.All

\#    (optional) Directory.Read.All to obtain a domain list and check whether an item is shared externally

After creating the application and granting the permissions, copy the key/secret and use it to configure the $secretsJson.'Sharepoint-ODFB-client_secret' variable. This same application, secret, and secrets file can also be used for `Graph_ODFB_shared_files.ps1` (note that this has one more permission required!).

If you need more help understanding all the concepts mentioned above, start with [this article](https://docs.microsoft.com/en-us/graph/auth/auth-concepts).

The script has two optional parameters you can use. The -ExpandFolders switch instructs it to enumerate files in any (sub)folders found under the root, and the -Depth parameter controls how deep the expansion is. The default value is 0, so only the top-most set of folders will be expanded. Use this parameter with care, while I've tested the script with few thousand items in multiple nested folders, I cannot guarantee it will work in all scenarios.

```PowerShell
.\Graph_SP_shared_files.ps1 -ExpandFolders -depth 2
```
To determine whether a file is externally shared, the script needs to know the list of domains configured in the company. If you have granted the corresponding permissions, it will fetch them automatically. Otherwise, you can populate the list manually at line 272.

By default, the script will return a filtered list of just the items that have been shared, and will also store the output in a global variable called $varSPSharedItems in case you want to reuse it. The unfiltered output will be saved to a CSV file in the user's 'Downloads' directory, which you can then format, sort and filter as needed.

Additional information about the script can be found at: https://practical365.com/clients/onedrive/reporting-on-onedrive-for-business-shared-files/

Questions and feedback are welcome.

/// 03.12.2024 Created script and .md based on Graph_ODFB_shared_files.ps1 -SM