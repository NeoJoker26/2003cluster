using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;

namespace _2003v5.Pages
{
    public class ConsecutiveResultsModel : PageModel
    {
        string competition;//FLG | ALL
        string homeaway;// ho | ao | ha
        public string heading;
        //lists
        public List<Consecutive_Wins> Consecutive_Wins_Table = new List<Consecutive_Wins>();
        public List<Consecutive_Draws> Consecutive_Draws_Table = new List<Consecutive_Draws>();
        public List<Consecutive_Defeats> Consecutive_Defeats_Table = new List<Consecutive_Defeats>();
        public List<Without_Losing> Without_Losing_Table = new List<Without_Losing>();
        public List<Without_Winning> Without_Winning_Table = new List<Without_Winning>();
        public List<Clean_Sheets> Clean_Sheets_Table = new List<Clean_Sheets>();
        public void OnGet()
        {
            //reading parameters
            competition = Request.Query["competition"];
            homeaway = Request.Query["homeaway"];
            //sqlcommands
            string colpref = "";
            string homeaway_val = "";
            switch (competition)
            {
                case "FLG":
                    colpref = "l";
                    heading = "Football League";
                    break;
                case "ALL":
                    colpref = "";
                    heading = "All Competitions";
                    break;
                default:
                    colpref = "";
                    heading = "All Competitions";
                    break;
            }
            switch (homeaway)
            {
                case "ho":
                    homeaway_val = "H";
                    heading += " (home only)";
                    break;
                case "ao":
                    homeaway_val = "A";
                    heading += " (away only)";
                    break;
                case "ha":
                    homeaway_val = " ";
                    break;
                default:
                    homeaway_val = " ";
                    break;

            }
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "select top 10 start_" + colpref + "wins, date, datediff(day, start_" + colpref + "wins, date) as interval, " + colpref + "wins from consecutive_results a where homeawayall = '" + homeaway_val + "'  and not exists (  select * from consecutive_results b  where homeawayall = '" + homeaway_val + "'  and b.start_" + colpref + "wins = a.start_" + colpref + "wins  and b.date > a.date ) order by " + colpref + "wins desc, date desc ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Consecutive_Wins table = new Consecutive_Wins();
                            table.start_wins = reader.GetDateTime(0);
                            table.date = reader.GetDateTime(1);
                            table.interval = reader.GetInt32(2);
                            table.wins = reader.GetByte(3);
                            Consecutive_Wins_Table.Add(table);
                        }
                    }
                }


                //SQL command
                sql = "select top 10 start_" + colpref + "draws, date, datediff(day, start_" + colpref + "draws, date) as interval, " + colpref + "draws from consecutive_results a where homeawayall = '" + homeaway_val + "'  and not exists (  select * from consecutive_results b  where homeawayall = '" + homeaway_val + "'  and b.start_" + colpref + "draws = a.start_" + colpref + "draws  and b.date > a.date ) order by " + colpref + "draws desc, date desc ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Consecutive_Draws table = new Consecutive_Draws();
                            table.start_draws = reader.GetDateTime(0);
                            table.date = reader.GetDateTime(1);
                            table.interval = reader.GetInt32(2);
                            table.draws = reader.GetByte(3);
                            Consecutive_Draws_Table.Add(table);
                        }
                    }
                }

                sql = "select top 10 start_" + colpref + "defeats, date, datediff(day, start_" + colpref + "defeats, date) as interval, " + colpref + "defeats from consecutive_results a where homeawayall = '" + homeaway_val + "'  and not exists (  select * from consecutive_results b  where homeawayall = '" + homeaway_val + "'  and b.start_" + colpref + "defeats = a.start_" + colpref + "defeats  and b.date > a.date ) order by " + colpref + "defeats desc, date desc ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Consecutive_Defeats table = new Consecutive_Defeats();
                            table.start_defeats = reader.GetDateTime(0);
                            table.date = reader.GetDateTime(1);
                            table.interval = reader.GetInt32(2);
                            table.defeats = reader.GetByte(3);
                            Consecutive_Defeats_Table.Add(table);
                        }
                    }
                }
                sql = "select top 10 start_" + colpref + "nodefeats, date, datediff(day, start_" + colpref + "nodefeats, date) as interval, " + colpref + "nodefeats from consecutive_results a where homeawayall = '" + homeaway_val + "'  and not exists (  select * from consecutive_results b  where homeawayall = '" + homeaway_val + "'  and b.start_" + colpref + "nodefeats = a.start_" + colpref + "nodefeats  and b.date > a.date ) order by " + colpref + "nodefeats desc, date desc ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Without_Losing table = new Without_Losing();
                            table.start_nodefeats = reader.GetDateTime(0);
                            table.date = reader.GetDateTime(1);
                            table.interval = reader.GetInt32(2);
                            table.nodefeats = reader.GetByte(3);
                            Without_Losing_Table.Add(table);
                        }
                    }
                }
                sql = "select top 10 start_" + colpref + "nowins, date, datediff(day, start_" + colpref + "nowins, date) as interval, " + colpref + "nowins from consecutive_results a where homeawayall = '" + homeaway_val + "'  and not exists (  select * from consecutive_results b  where homeawayall = '" + homeaway_val + "'  and b.start_" + colpref + "nowins = a.start_" + colpref + "nowins  and b.date > a.date ) order by " + colpref + "nowins desc, date desc ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Without_Winning table = new Without_Winning();
                            table.start_nowins = reader.GetDateTime(0);
                            table.date = reader.GetDateTime(1);
                            table.interval = reader.GetInt32(2);
                            table.nowins = reader.GetByte(3);
                            Without_Winning_Table.Add(table);
                        }
                    }
                }
                sql = "select top 10 start_" + colpref + "cleansheets, date, datediff(day, start_" + colpref + "cleansheets, date) as interval, " + colpref + "cleansheets from consecutive_results a where homeawayall = '" + homeaway_val + "'  and not exists (  select * from consecutive_results b  where homeawayall = '" + homeaway_val + "'  and b.start_" + colpref + "cleansheets = a.start_" + colpref + "cleansheets  and b.date > a.date ) order by " + colpref + "cleansheets desc, date desc ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Clean_Sheets table = new Clean_Sheets();
                            table.start_cleansheets = reader.GetDateTime(0);
                            table.date = reader.GetDateTime(1);
                            table.interval = reader.GetInt32(2);
                            table.cleansheets = reader.GetByte(3);
                            Clean_Sheets_Table.Add(table);
                        }
                    }
                }
            }
        }




        public class Consecutive_Wins
        {
            public DateTime start_wins;
            public DateTime date;
            public int interval;
            public Byte wins;
        }
        public class Consecutive_Draws
        {
            public DateTime start_draws;
            public DateTime date;
            public int interval;
            public Byte draws;
        }
        public class Consecutive_Defeats
        {
            public DateTime start_defeats;
            public DateTime date;
            public int interval;
            public Byte defeats;
        }
        public class Without_Losing
        {
            public DateTime start_nodefeats;
            public DateTime date;
            public int interval;
            public Byte nodefeats;
        }
        public class Without_Winning
        {
            public DateTime start_nowins;
            public DateTime date;
            public int interval;
            public Byte nowins;
        }
        public class Clean_Sheets
        {
            public DateTime start_cleansheets;
            public DateTime date;
            public int interval;
            public Byte cleansheets;
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