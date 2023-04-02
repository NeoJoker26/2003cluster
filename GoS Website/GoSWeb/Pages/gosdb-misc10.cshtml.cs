using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Collections.Generic;
using System.Data.SqlClient;
using static GoSWeb.Pages.gosdb_misc12Model;

namespace GoSWeb.Pages
{
    public class gosdb_misc10Model : PageModel
    {
        public string competition;
        public string season1;
        public string season2;
        public string heading;
        public List<Goalscorers> GRST = new List<Goalscorers>(); //Goalscorers Ranked by Season Table
        public void OnGet()
        {
            //parameter here:
            competition = Request.Query["competition"];
            season1 = Request.Query["season1"];
            season2 = Request.Query["season2"];

            string tableview = "";

            switch (competition)
            {
                case "LG":                   
                    tableview = "v_match_all_league";
                    heading = "League";
                    break;
                  case "FAC":
                    tableview = "v_match_FA";
                    heading = "FA Cup";
                 break;
                case "CUP":
                    tableview = "v_match_cups";
                    heading = "All Cups";
                        break;
                default:
                    tableview = "v_match_all";
                    heading = "All Competitions";
                    break;
            }

            switch(season1)
            {
                case "":
                    season1 = "112";
                    break;
                case null:
                    season1 = "112";
                    break;
                default:
                    break;
            }
            switch (season2)
            {
                case "":
                    season2 = "112";
                    break;
                case null:
                    season2 = "112";
                    break;
                default:
                    break;
            }

            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "with CTE1 as \r\n( \r\nselect years, player_id_spell1, surname, nullif(forename, initials) as firstname, \r\n count(c.player_id) as goals, count(distinct b.date) as games, round(count(c.player_id)/cast(count(distinct b.date) as dec(7,3)),2) as pergame \r\nfrom "+ tableview +" a join season on date between date_start and date_end \r\njoin match_player b on a.date = b.date \r\nleft outer join match_goal c on b.player_id = c.player_id and b.date = c.date \r\njoin player d on b.player_id = d.player_id \r\nwhere season_no between "+ season1+" and "+ season2+ "\r\ngroup by years, player_id_spell1, forename, surname, initials \r\n), \r\nCTE2 as \r\n( \r\nselect player_id_spell1, count(distinct years) as seasons, \r\n count(c.player_id) as goals, count(distinct b.date) as games, round(count(c.player_id)/cast(count(distinct b.date) as dec(7,3)),2) as pergame \r\nfrom " + tableview +" a join season on date between date_start and date_end \r\njoin match_player b on a.date = b.date \r\nleft outer join match_goal c on b.player_id = c.player_id and b.date = c.date \r\njoin player d on b.player_id = d.player_id \r\ngroup by player_id_spell1 \r\n) \r\nselect a.player_id_spell1, a.years, rank() over (partition by a.years order by a.goals desc) as rank, trim(isnull(rtrim(a.firstname),'') + ' ' + a.surname) as player, \r\n a.goals, a.games, a.pergame, seasons, b.goals as totgoals, b.games as totgames, b.pergame as totpergame \r\nfrom CTE1 a join CTE2 b on a.player_id_spell1 = b.player_id_spell1 \r\nwhere a.goals > 0 \r\norder by a.years, rank, a.surname";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Goalscorers GGRST = new Goalscorers();
                            GGRST.player_id_spell1 = reader.GetInt16(0);
                            GGRST.year = reader.GetString(1);
                            GGRST.rank = (int)reader.GetInt64(2);
                            GGRST.player = reader.GetString(3);
                            GGRST.goals = reader.GetInt32(4);
                            GGRST.games = reader.GetInt32(5);
                            GGRST.pergame = reader.GetDecimal(6);
                            GGRST.seassons = reader.GetInt32(7);
                            GGRST.totgoals = reader.GetInt32(8);
                            GGRST.totgames = reader.GetInt32(9);
                            GGRST.totpergame = reader.GetDecimal(10);
                            GRST.Add(GGRST);
                        }
                    }
                }
            }
        }
        public class Goalscorers
        {
            public int player_id_spell1;
            public string year;
            public int rank;
            public string player;
            public int goals;
            public int games;
            public decimal pergame;
            public int seassons;
            public int totgoals;
            public int totgames;
            public decimal totpergame;
        }
    }
}
