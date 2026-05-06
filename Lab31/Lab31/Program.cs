using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;

namespace Lab31
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== SharePoint SSOM CRUD Operations ===");
            
            // Thay đổi URL thành địa chỉ website SharePoint thực tế của bạn
            string siteUrl = "http://win-oe9a92lkgih/sites/test1"; 
            string listName = "Course";

            try
            {
                using (SPSite site = new SPSite(siteUrl))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList list = web.Lists.TryGetList(listName);
                        if (list != null)
                        {
                            // 1. Create (C)
                            Console.WriteLine("\n[1] Creating a new item...");
                            SPListItem newItem = list.Items.Add();
                            
                            // Note: The default column "Name" that you see usually has the Internal Name "Title"
                            newItem["Title"] = "SharePoint SSOM Programming"; 
                            newItem["Description"] = "Course on Server-Side Object Model in SharePoint";
                            newItem.Update();
                            Console.WriteLine($"-> Successfully created item with ID: {newItem.ID}");

                            // 2. Read (R)
                            Console.WriteLine("\n[2] Reading list of items...");
                            foreach (SPListItem item in list.Items)
                            {
                                Console.WriteLine($"- ID: {item.ID} | Name: {item["Title"]} | Description: {item["Description"]}");
                            }

                            // 3. Update (U)
                            Console.WriteLine($"\n[3] Updating item with ID {newItem.ID}...");
                            SPListItem itemToUpdate = list.GetItemById(newItem.ID);
                            itemToUpdate["Title"] = "Advanced SharePoint SSOM Programming";
                            itemToUpdate.Update();
                            Console.WriteLine("-> Update successful.");

                            // 4. Delete (D)
                            Console.WriteLine($"\n[4] Deleting item with ID {itemToUpdate.ID}...");
                            SPListItem itemToDelete = list.GetItemById(itemToUpdate.ID);
                            itemToDelete.Delete();
                            Console.WriteLine("-> Deletion successful.");
                        }
                        else
                        {
                            Console.WriteLine($"Error: Cannot find List '{listName}' on the site.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
    }
}
