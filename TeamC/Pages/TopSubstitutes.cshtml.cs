using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Collections.Generic;
using System.Data.SqlClient;


namespace _2003v5.Pages
{
    public class TopSubstitutesModel : PageModel
    {
        public int AP1; // All Player 
        public int AP2;
        public int AP3;
        public int AP4;
        public int CS1; // Current Player
        public int CS2;
        public int CS3;
        public int CS4;
        public List<Player> HSA = new List<Player>(); // Highest Substitute Appearances
        public List<Player> TSG = new List<Player>(); // Top Substitute Goalscorers
        public List<Player> HSA2 = new List<Player>();
        public List<Player> TSG2 = new List<Player>();
        string competition;
        public void OnGet()
        {
            //read parameters
            competition = Request.Query["competition"];
            //
            string view;

            switch (competition)
            {
                case "FLG": //Football League
                    view = "v_match_FL";
                    break;
                case "FAC":
                    view = "v_match_FA";
                    break;
                case "FLC":
                    view = "v_match_LC";
                    break;
                default:  //All league 
                    view = "v_match_all";
                    break;
            }

            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "select top 5 player_id_spell1, surname, forename, initials, count(*) as count from " + view + " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id where startpos = 0 group by player_id_spell1, surname, forename, initials order by count desc, surname ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Player GHSA = new Player();
                            GHSA.player_id_spell1 = reader.GetInt16(0);
                            GHSA.surname = reader.GetString(1);
                            GHSA.forename = reader.GetString(2);
                            GHSA.initials = reader.GetString(3);
                            GHSA.count = reader.GetInt32(4);
                            HSA.Add(GHSA);
                        }
                    }
                }
                sql = "select top 5 player_id_spell1, surname, forename, initials, count(*) as count from " + view + " a join match_player b1 on a.date = b1.date join match_goal b2 on a.date = b2.date and b1.player_id = b2.player_id join player c on b1.player_id = c.player_id where startpos = 0 group by player_id_spell1, surname, forename, initials order by count desc, surname ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Player GTSG = new Player();
                            GTSG.player_id_spell1 = reader.GetInt16(0);
                            GTSG.surname = reader.GetString(1);
                            GTSG.forename = reader.GetString(2);
                            GTSG.initials = reader.GetString(3);
                            GTSG.count = reader.GetInt32(4);
                            TSG.Add(GTSG);
                        }
                    }
                }
                sql = "select top 5 player_id_spell1, surname, forename, initials, count(*) as count from " + view + " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id where startpos = 0 and last_game_year = 9999 group by player_id_spell1, surname, forename, initials order by count desc, surname ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Player GHSA2 = new Player();
                            GHSA2.player_id_spell1 = reader.GetInt16(0);
                            GHSA2.surname = reader.GetString(1);
                            GHSA2.forename = reader.GetString(2);
                            GHSA2.initials = reader.GetString(3);
                            GHSA2.count = reader.GetInt32(4);
                            HSA2.Add(GHSA2);
                        }
                    }
                }
                sql = "select top 5 player_id_spell1, surname, forename, initials, count(*) as count from " + view + " a join match_player b1 on a.date = b1.date join match_goal b2 on a.date = b2.date and b1.player_id = b2.player_id join player c on b1.player_id = c.player_id where startpos = 0 and last_game_year = 9999 group by player_id_spell1, surname, forename, initials order by count desc, surname ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Player GTSG2 = new Player();
                            GTSG2.player_id_spell1 = reader.GetInt16(0);
                            GTSG2.surname = reader.GetString(1);
                            GTSG2.forename = reader.GetString(2);
                            GTSG2.initials = reader.GetString(3);
                            GTSG2.count = reader.GetInt32(4);
                            TSG2.Add(GTSG2);
                        }
                    }
                }
                sql = "select count(distinct player_id_spell1) as count from " + view + " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            AP1 = reader.GetInt32(0);
                        }
                    }
                }
                sql = "select count(distinct player_id_spell1) as count from " + view + " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id where startpos > 0 ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            AP2 = reader.GetInt32(0);
                        }
                    }
                }
                sql = "select count(distinct player_id_spell1) as count from " + view + " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id where startpos = 0 ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            AP3 = reader.GetInt32(0);
                        }
                    }
                }
                sql = "select count(*) as count from " + view + " a join match_player b on a.date = b.date where startpos = 0 ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            AP4 = reader.GetInt32(0);
                        }
                    }
                }
                sql = "select count(distinct player_id_spell1) as count from " + view + " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id where last_game_year = 9999 ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            CS1 = reader.GetInt32(0);
                        }
                    }
                }
                sql = "select count(distinct player_id_spell1) as count from " + view + " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id where startpos > 0 and last_game_year = 9999 ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            CS2 = reader.GetInt32(0);
                        }
                    }
                }
                sql = "select count(distinct player_id_spell1) as count from " + view + " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id where startpos = 0 and last_game_year = 9999 ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            CS3 = reader.GetInt32(0);
                        }
                    }
                }
                sql = "select count(*) as count from " + view + " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id where startpos = 0 and last_game_year = 9999 ";

                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            CS4 = reader.GetInt32(0);
                        }
                    }
                }
            }

        }
        public class Player
        {
            public int player_id_spell1;
            public string surname;
            public string forename;
            public string initials;
            public int count;
        }
    }
}