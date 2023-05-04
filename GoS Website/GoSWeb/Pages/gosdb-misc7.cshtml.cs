using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using System.Data.Entity.Core.Common.EntitySql;
using static GoSWeb.Pages.gosdb_misc5Model;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace GoSWeb.Pages
{
    public class gosdb_misc7Model : PageModel
    {
        public string weighting;
        public List<FT> FLRDT = new List<FT>(); //Football League Record by Decade Table
        public void OnGet()
        {
            weighting = Request.Query["weighting"];

            switch (weighting)
            {
                case "1":
                    break;
                case "1.1":
                    break;
                case "1.2":
                    break;
                case "1.3":
                    break;
                case "1.4":
                    break;
                case "1.5":
                    break;
                case "1.6":
                    break;
                case "1.7":
                    break;
                case "1.8":
                    break;
                case "1.9":
                    break;
                case "2":
                    break;
                case null:
                    weighting = "1.2";
                    break;
                case "":
                    weighting = "1.2";
                    break;
                default:
                    weighting = "1.2";
                    break;
            }

            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "WITH CTE1 as \r\n(select date, 1 as matches, \r\nrow_number() over (partition by years order by date) as game, \r\ncase when goalsfor > goalsagainst then 1 else 0 end as wins, \r\ncase when goalsfor = goalsagainst then 1 else 0 end as draws, \r\ncase when goalsfor < goalsagainst then 1 else 0 end as defeats, \r\ncase when goalsfor > goalsagainst then 3 \r\n\t when goalsfor = goalsagainst then 1 \r\n\t else 0  \r\n\t end as modernpoints, \r\ncase when homeaway = 'H' then attendance else NULL end as attendance \r\nfrom [v_match_FL-39]  a join season b \r\non a.date >= b.date_start and a.date <= b.date_end \r\n), \r\nCTE2 AS \r\n(select 1 as matches, decade, \r\ncase when game = 1 and tier = 1 then 1 else 0 end as tier1, \r\ncase when game = 1 and tier = 2 then 1 else 0 end as tier2, \r\ncase when game = 1 and tier = 3 then 1 else 0 end as tier3, \r\ncase when game = 1 and tier = 4 then 1 else 0 end as tier4, \r\nwins, draws, defeats, modernpoints, power(" + weighting + " ,4-tier)*modernpoints as weightedpoints, \r\nattendance \r\nfrom CTE1 a join season b \r\non a.date >= b.date_start and a.date <= b.date_end \r\n) \r\nselect case when grouping(decade) = 1 then 'All' else decade end as decade, \r\nsum(tier1) as tier1, sum(tier2) as tier2, sum(tier3) as tier3, sum(tier4) as tier4, \r\nsum (matches) as matches, sum(wins) as wins, sum(draws) as draws, sum(defeats) as defeats, \r\nsum(modernpoints)*100/sum(matches) as modernpoints, sum(weightedpoints)*100/sum(matches) as weightedpoints, \r\navg(attendance) as attendance \r\nfrom CTE2 \r\ngroup by decade with rollup \r\norder by decade ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            FT GFLRDT = new FT();
                            GFLRDT.decade = reader.GetString(0);
                            GFLRDT.tier1 = reader.GetInt32(1);
                            GFLRDT.tier2 = reader.GetInt32(2);
                            GFLRDT.tier3 = reader.GetInt32(3);
                            GFLRDT.tier4 = reader.GetInt32(4);
                            GFLRDT.matches = reader.GetInt32(5);
                            GFLRDT.wins = reader.GetInt32(6);
                            GFLRDT.draws = reader.GetInt32(7);
                            GFLRDT.defeats = reader.GetInt32(8);
                            GFLRDT.modernpoints = reader.GetInt32(9);
                            try { 
                                GFLRDT.weightedpoints = (float)reader.GetDecimal(10);
                             }
                            catch(Exception ex)
                            {
                                GFLRDT.weightedpoints = (float)reader.GetInt32(10);
                            }
                            GFLRDT.attendance = reader.GetInt32(11);
                            string[] playerinfo = Playerinfo(reader.GetString(0));
                            GFLRDT.PlayerSurname1 = playerinfo[0];
                            GFLRDT.i1 = playerinfo[1];
                            GFLRDT.Goal1 = Int32.Parse(playerinfo[2]);                           
                            GFLRDT.PlayerSurname2 = playerinfo[3];
                            GFLRDT.i2 = playerinfo[4];
                            GFLRDT.Goal2 = Int32.Parse(playerinfo[5]);
                            FLRDT.Add(GFLRDT);

                        }
                    }
                }
            }
        }
        public class FT // footubal table
        {
            public string decade;
            public int tier1;
            public int tier2;
            public int tier3;
            public int tier4;
            public int matches;
            public int wins;
            public int draws;
            public int defeats;
            public int modernpoints;
            public float weightedpoints;
            public int attendance;
            public string PlayerSurname1;
            public string PlayerSurname2;
            public string i1;
            public string i2;
            public int Goal1;
            public int Goal2;
        }
        public string Weightedpointsratio(float weightedpoints)
        {
            return (weightedpoints / 1.5).ToString();
        }
        public string[] Playerinfo(string decades) 
        {
            string[] playerinfo = new string[6];
            string[] decade = new string[2];
            if (decades == "All")
            {
                decade[0] = "1920";
                decade[1] = "9999"; // update this number in the year 9999 
            }
            else           
                 decade = decades.Split('-');
                       
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "select top 1 player_id_spell1, surname, initials, count(distinct b.date) as appears \r\nfrom [v_match_FL-39] a join match_player b on a.date = b.date \r\njoin player d on b.player_id = d.player_id \r\nwhere year(a.date) between '"+ decade[0] + "' and '"+ decade[1] + "' \r\n  and startpos > 0 \r\ngroup by player_id_spell1, surname, initials \r\norder by appears desc ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            playerinfo[0] = reader.GetString(1);
                            playerinfo[1] = reader.GetString(2);
                            playerinfo[2] = reader.GetInt32(3).ToString();
                        }
                    }
                }
                sql = "select top 1 player_id_spell1, surname, initials, count(distinct b.date) as goals \r\nfrom [v_match_FL-39] a join match_goal b on a.date = b.date \r\njoin player d on b.player_id = d.player_id \r\nwhere year(a.date) between '"+ decade[0] + "' and '"+ decade[1] + "' \r\n and player_id_spell1 < 9000 \r\ngroup by player_id_spell1, surname, initials \r\norder by goals desc ";
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            playerinfo[3] = reader.GetString(1);
                            playerinfo[4] = reader.GetString(2);
                            playerinfo[5] = reader.GetInt32(3).ToString();
                        }
                    }
                }
            }




            return playerinfo;
        }
    }
}
