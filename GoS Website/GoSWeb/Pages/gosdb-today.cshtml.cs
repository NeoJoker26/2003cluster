using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;
using static System.Data.Entity.Infrastructure.Design.Executor;

namespace GoSWeb.Pages
{
    public class gosdb_todayModel : PageModel
    {
        public void OnGet()
        {
            //This page is run every date at 02:00 by a Plesk trigger, to provide a simple indication of the date of the database backup on Steve's PC 
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";

            string sql = "update today set date = GETDATE()";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //we not expecting a answer
                    }
                }
            }
        }
    }
}
