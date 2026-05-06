using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;

namespace SSOM
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== SSOM Hello World cho SharePoint 2019 ===");

            // Thay đổi URL dưới đây bằng URL site SharePoint của bạn
            string siteUrl = "http://localhost"; 

            try
            {
                using (SPSite site = new SPSite(siteUrl))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        Console.WriteLine("Kết nối thành công!");
                        Console.WriteLine("Tên Web (Title): " + web.Title);
                        Console.WriteLine("URL: " + web.Url);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Có lỗi xảy ra: " + ex.Message);
                Console.WriteLine("Lưu ý: Mã SSOM phải được chạy trực tiếp trên máy chủ cài đặt SharePoint và ứng dụng phải được build ở chế độ x64.");
            }

            Console.WriteLine("Nhấn phím bất kỳ để thoát...");
            Console.ReadKey();
        }
    }
}
