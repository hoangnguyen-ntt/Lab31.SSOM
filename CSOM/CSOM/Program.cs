using System; // package System để sử dụng Console và Exception
using System.Security; // package System.Security để sử dụng SecureString nếu cần thiết
using Microsoft.SharePoint.Client; // package Microsoft.SharePoint.Client để sử dụng ClientContext, Web và các lớp liên quan đến SharePoint Client Object Model (CSOM)

namespace CSOM
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Thay đổi URL site của bạn ở đây
            string siteUrl = "http://WIN-OE9A92LKGIH/sites/test1";

            // Thay đổi thông tin đăng nhập của bạn
            string userName = "Administrator";
            string password = "123456Aa@";
            string domain = "smaclink.local"; // Bỏ trống nếu không dùng domain

            try
            {
                using (ClientContext ctx = new ClientContext(siteUrl))
                {
                    // Chuyển lại về xác thực Default
                    ctx.AuthenticationMode = ClientAuthenticationMode.Default;

                    // Sử dụng luôn tài khoản Windows mặc định (tài khoản đang login vào máy)
                    ctx.Credentials = System.Net.CredentialCache.DefaultCredentials;

                    // Lấy đối tượng Web hiện tại
                    Web web = ctx.Web;

                    // Chỉ tải thuộc tính Title của Web để tối ưu hóa context
                    ctx.Load(web, w => w.Title);

                    // Thực thi yêu cầu lên máy chủ
                    ctx.ExecuteQuery();

                    Console.WriteLine("Connect successfully to SharePoint 2019!");
                    Console.WriteLine("Name of site: " + web.Title);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Có lỗi xảy ra: " + ex.Message);
            }

            Console.WriteLine("Nhấn phím bất kỳ để thoát...");
            Console.ReadKey();
        }
    }
}
