using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;

namespace GoSWeb.Pages
{
    public class IndexModel : PageModel
    {
        public List<Matches> MatchesTable = new List<Matches>();
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet()
        {
            try
            {
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";;
            using(SqlConnection cnn = new SqlConnection(connetionString))
            {
                    cnn.Open();
                    string sql = "SELECT * FROM [greenson_greensonscreen].[dbo].[country]";
                    using(SqlCommand command = new SqlCommand(sql, cnn))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Matches table = new Matches();
                                table.country = reader.GetString(0);

                                MatchesTable.Add(table);
                            }
                        }
                    }
            }
             

           

            }
            catch(Exception ex)
            {

            }


           
        }
        
        public class Matches
        {
            public string country;
        }
    }
}
