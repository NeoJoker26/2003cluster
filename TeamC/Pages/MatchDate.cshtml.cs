using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;
using System.Net;
using Microsoft.AspNetCore;
using System.Web;

namespace _2003v5.Pages;

public class MatchDateModel : PageModel
{
    public string yearsSeasons;
    public string yearsMatches;
    
    public void OnGet()
    {
      
    }
    public class Matchyears
    {
        
            public int season_no { get; set; }
            public DateTime date_start { get; set; }
            public DateTime date_end { get; set; }
            public string years { get; set; }
            public string decade { get; set; }
            public string promrel { get; set; }
            public int endpos { get; set; }
            public int tier { get; set; }
            public int teams_in_div { get; set; }
            public int teams_above_div { get; set; }
            public string division { get; set; }
            public string divistion_short { get; set; }
            public int id { get; set; }
            public int pos_promote { get; set; }
            public int pos_promote_playoff { get; set; }
            public int pos_relegate { get; set; }
            public int pos_relegate_playoff { get; set; }
            public DateTime matchdate { get; set; }
            public int matchyear { get; set; }
            public int matchmon { get; set; }
            public DateTime matchday { get; set; }
            public int matchdecade { get; set; }
            public int matchseason { get; set; }
            public int phase { get; set; }
            public int shortrange { get; set; }
            public int lastid { get; set; }
    }
}
//   string connetionString =
//     @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
// using (SqlConnection cnn = new SqlConnection(connetionString))
// {
//     cnn.Open();
//     string sql = "";
//     using (SqlCommand command = new SqlCommand(sql, cnn))
//     {
//         using (SqlDataReader reader = command.ExecuteReader())
//         {
//             while (reader.Read())
//             {
//                         
//             }
//         }
//     }
// } 