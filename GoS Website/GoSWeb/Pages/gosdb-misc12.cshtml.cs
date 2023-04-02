using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using static GoSWeb.Pages.gosdb_misc3Model;

namespace GoSWeb.Pages
{
    public class gosdb_misc12Model : PageModel
    {
        public List<Appearances> AppearancesTable = new List<Appearances>();
        public List<Appearances> AppearancesTable2 = new List<Appearances>();
        public List<Appearances2> AppearancesTable3 = new List<Appearances2>();
        public void OnGet()
        {
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                string sql = "select rank() over (order by consec_count desc) as rank, a.player_id, surname, forename, consec_count, start_date, end_date from consecutive_appears a join player b on a.player_id = b.player_id where consec_count >= 50 order by consec_count desc, start_date ";
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Appearances GAppearancestable = new Appearances(); //  G stand for Get 
                            GAppearancestable.rank = (int)reader.GetInt64(0);
                            GAppearancestable.player_id = reader.GetInt16(1);
                            GAppearancestable.surname = reader.GetString(2);
                            GAppearancestable.forename = reader.GetString(3);
                            GAppearancestable.consec_count = reader.GetInt16(4);
                            GAppearancestable.start_date = reader.GetDateTime(5);
                            GAppearancestable.end_date = reader.GetDateTime(5);
                            AppearancesTable.Add(GAppearancestable);
                        }
                    }
                }
                sql = "select rank() over (order by l_consec_count desc) as rank, a.player_id, surname, forename, l_consec_count, l_start_date, l_end_date from consecutive_appears a join player b on a.player_id = b.player_id where l_consec_count >= 50 order by l_consec_count desc, l_start_date ";
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Appearances GAppearancestable = new Appearances(); //  G stand for Get 
                            GAppearancestable.rank = (int)reader.GetInt64(0);
                            GAppearancestable.player_id = reader.GetInt16(1);
                            GAppearancestable.surname = reader.GetString(2);
                            GAppearancestable.forename = reader.GetString(3);
                            GAppearancestable.consec_count = reader.GetInt16(4);
                            GAppearancestable.start_date = reader.GetDateTime(5);
                            GAppearancestable.end_date = reader.GetDateTime(5);
                            AppearancesTable2.Add(GAppearancestable);
                        }
                    }
                }
                sql = "select rank() over (order by goal_consec_count desc) as rank, a.player_id, surname, forename, goal_consec_count, goal_count, goal_start_date, goal_end_date from consecutive_appears a join player b on a.player_id = b.player_id where goal_consec_count >= 3 order by goal_consec_count desc, goal_start_date ";
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Appearances2 GAppearancestable = new Appearances2(); //  G stand for Get 
                        GAppearancestable.rank = (int)reader.GetInt64(0);
                        GAppearancestable.player_id = reader.GetInt16(1);
                        GAppearancestable.surname = reader.GetString(2);
                        GAppearancestable.forename = reader.GetString(3);
                        GAppearancestable.goal_consec_count = reader.GetInt16(4);
                        GAppearancestable.goal_count = reader.GetInt16(5);
                        GAppearancestable.start_date = reader.GetDateTime(6);
                        GAppearancestable.end_date = reader.GetDateTime(7);
                        AppearancesTable3.Add(GAppearancestable);
                    }
                }
                }
            }
            
        }
    public class Appearances
    {
        public int rank;
        public int player_id;
        public string surname;
        public string forename;
        public int consec_count;
        public DateTime start_date;
        public DateTime end_date;
    }
    public class Appearances2
    {
        public int rank;
        public int player_id;
        public string surname;
        public string forename;
        public int goal_consec_count;
        public int goal_count;
        public DateTime start_date;
        public DateTime end_date;
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
    }
    
}
    

