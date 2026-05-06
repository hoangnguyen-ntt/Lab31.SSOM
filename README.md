# Lab Guide: Create a SharePoint 2019 SSOM App

## Objective
Build a simple **SharePoint Server Object Model (SSOM)** application for **SharePoint Server 2019** that connects to a site and reads data from the farm.

## Prerequisites
- SharePoint Server 2019 installed
- Visual Studio installed on a SharePoint server or a machine with SharePoint 2019 client/server assemblies available
- .NET Framework 4.7.2 or later
- Permission to access the SharePoint farm
- A SharePoint site URL, for example: `http://sharepoint2019`

## Important Notes
- SSOM runs **on the SharePoint server only**.
- Use SSOM for **farm-side** solutions, not for remote client apps.
- The application must be compiled with **Full Trust** and run with an account that has access to SharePoint.

## Lab Scenario
Create a console application that:
1. Connects to a SharePoint 2019 web application
2. Reads the site title
3. Displays the title in the console

---

## Step 1: Create a Console Project
1. Open **Visual Studio**
2. Create a new project
3. Select **Console App (.NET Framework)**
4. Name it `SSOMLab`
5. Target **.NET Framework 4.7.2** or later

---

## Step 2: Add SharePoint References
Add references to the SharePoint assemblies from:

```text
C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI
```

Add these references:
- `Microsoft.SharePoint.dll`
- `Microsoft.SharePoint.Security.dll`

## Step 3: Write the Code
Replace `Program.cs` with the following:

```csharp
using System;
using Microsoft.SharePoint;

namespace SSOMLab
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = "http://sharepoint2019";

            try
            {
                using (SPSite site = new SPSite(siteUrl))
                {
                    using (SPWeb web = site.RootWeb)
                    {
                        Console.WriteLine("SharePoint site connected successfully.");
                        Console.WriteLine("Web Title: " + web.Title);
                        Console.WriteLine("Web URL: " + web.Url);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
```

---

## Step 4: Build the Project
1. Set the project platform to **x64** or **Any CPU**
2. Build the solution
3. Fix any missing reference or framework errors

---

## Step 5: Run the Application
Run the console app from Visual Studio or from the output folder.

Expected output:

```text
SharePoint site connected successfully.
Web Title: <your site title>
Web URL: http://sharepoint2019
```

---

## Step 6: Test with a List
To read items from a SharePoint list, use this sample code:

```csharp
using System;
using Microsoft.SharePoint;

namespace SSOMLab
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = "http://sharepoint2019";

            using (SPSite site = new SPSite(siteUrl))
            using (SPWeb web = site.RootWeb)
            {
                SPList list = web.Lists["Documents"];

                Console.WriteLine("List Title: " + list.Title);
                Console.WriteLine("Item Count: " + list.ItemCount);
            }
        }
    }
}
```

---

## Common Problems
### 1. `FileNotFoundException`
- Check that SharePoint 2019 DLL references are added correctly
- Verify the path to the `ISAPI` folder

### 2. `UnauthorizedAccessException`
- Run the app with a user account that has permission to the SharePoint site
- Confirm the account is a farm or site collection administrator if needed

### 3. `PlatformNotSupportedException`
- Ensure the app runs on a SharePoint server
- SSOM is not supported in remote client applications

---

## Cleanup
- Remove temporary code
- Delete unused references
- Confirm the final build is successful

## Result
A working SharePoint 2019 SSOM console application that connects to a site and reads SharePoint data.
