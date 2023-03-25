using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;

namespace GoSWeb.Pages
{
    public class youtubeviewModel : PageModel
    {
        public List<Video> VideoTable = new List<Video>();
        public string parm;
        public void OnGet()
        {
            parm = Request.Query["parm"];
            string sqlqual1;
            string sqlqual2;
            string sqlorder;

            if (parm == "1") //review this sqlqual1 code with client
            {
                sqlqual1 = "where event_published = 'Y' and event_type = 'V' ";
                sqlqual2 = "and datediff(" + "dd" + ", publish_timestamp, getdate()) < 10 ";
                sqlorder = "order by cast(publish_timestamp as date) desc, date ";
            }
            else
            {
                sqlqual1 = "where event_published = 'Y' and ((event_type = 'M' and material_type = 'Y') or event_type = 'V') ";
                sqlqual2 = "";
                sqlorder = "order by date ";
            }
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                string sql = "select date, publish_timestamp, opposition, goalsfor, goalsagainst, homeaway, material_details1, straight_to_youtube from match join event_control on date = event_date " + sqlqual1 + sqlqual2 + sqlorder;
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Video table = new Video();
                            table.date = reader.GetDateTime(0);
                            table.publish_timestamp = reader.GetDateTime(1);
                            table.oppositon = reader.GetString(2);
                            table.goalfor = Convert.ToInt32(reader.GetInt16(3));
                            table.goalagainst = Convert.ToInt32(reader.GetInt16(4));
                            table.homeaway = reader.GetString(5);
                            table.material_detailes1 = reader.GetString(6);
                            try
                            {
                                table.straight_to_youtube = reader.GetString(7);
                            }
                            catch
                            {
                                table.straight_to_youtube = "N";
                            }
                            
                            VideoTable.Add(table);
                        }
                    }


                }
            }
        }
        public string MonthName(DateTime date)
        {            
            switch (date.Month)
            {
                case 1:
                    return "Jan";                   
                case 2:
                    return "Feb";                  
                case 3:
                    return "Mar";                  
                case 4:
                    return "Apr";
                case 5:
                    return "May";                    
                case 6:
                    return "Jun";                   
                case 7:
                    return "Jul";
                case 8:
                    return "Aug";                
                case 9:
                    return "Sep";
                case 10:
                    return "Oct";
                case 11:
                    return "Nov";
                case 12:
                    return "Dec";
            }
                       
                return Convert.ToString(date.Month);           
        }
        public class Video
        {
            public DateTime date;
            public DateTime publish_timestamp;
            public string oppositon;
            public int goalfor;
            public int goalagainst;
            public string homeaway;
            public string material_detailes1;
            public string straight_to_youtube;
        }
    }
}
