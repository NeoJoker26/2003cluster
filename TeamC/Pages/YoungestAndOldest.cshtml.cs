using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace _2003v5.Pages
{
    public class YoungestAndOldestModel : PageModel
    {
        string competition;
        public char mg = '*'; //multiple goals
        public string heading;
        public string tablehead = "";
        public List<count> Playercount = new List<count>();
        public List<PlayerM> YPD = new List<PlayerM>();
        public List<PlayerM> OPD = new List<PlayerM>();
        public List<PlayerM> OPA = new List<PlayerM>();
        public List<PlayerM> YS = new List<PlayerM>();
        public List<PlayerM> YSD = new List<PlayerM>();
        public List<PlayerM> YSSD = new List<PlayerM>();
        public List<PlayerM> OS = new List<PlayerM>();

        public void OnGet()
        {
            //reading parameter
            competition = Request.Query["competition"];
            //
            string LFCvalue = "";
            string compcode = "";
            //compcode needed
            switch (competition)
            {
                case "FLG":
                    LFCvalue = "'F'";
                    heading = "Football League";
                    tablehead = " in the Football League";
                    break;
                case "FAC":
                    LFCvalue = "'C'";
                    compcode = " and compcode = ('FAC') ";
                    heading = "FA Cup";
                    tablehead = " in the FA Cup";
                    break;
                case "CUP":
                    LFCvalue = "'C'";
                    heading = "Any Cup Competition";
                    tablehead = " in any Cup";
                    break;
                default:
                    LFCvalue = "'L','F','C'";
                    heading = "All Competitions";
                    break;
            }
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "select 'A', count(distinct a.player_id) as playercount from player a join match_player b on a.player_id = b.player_id where spell = 1 and dob is not null union all select 'B', count(distinct a.player_id) from player a join match_player b on a.player_id = b.player_id where spell = 1 order by 1  ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            count c = new count();
                            c.letter = reader.GetString(0);
                            c.playercount = reader.GetInt32(1);
                            Playercount.Add(c);
                        }
                    }
                }
                sql = "select top 50 rank() over (order by datediff(day,dob,date)) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age from  player a join match_player b on a.player_id = b.player_id where date = (select min(b1.date) from player a1 join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  where a1.player_id_spell1 = a.player_id_spell1 and LFC in (" + LFCvalue + ")" + compcode + " )and dob is not null ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            PlayerM GYPD = new PlayerM();
                            GYPD.rank = (int)reader.GetInt64(0);
                            GYPD.player_id_spell1 = (int)reader.GetInt16(1);
                            GYPD.surname = reader.GetString(2);
                            GYPD.forename = reader.GetString(3);
                            GYPD.initial = reader.GetString(4);
                            GYPD.dob = reader.GetDateTime(5);
                            GYPD.date = reader.GetDateTime(6);
                            GYPD.age = reader.GetInt32(7);
                            YPD.Add(GYPD);
                        }
                    }
                }
                sql = "select top 50 rank() over (order by datediff(day,dob,date) desc) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age from  player a join match_player b on a.player_id = b.player_id where date = (select min(b1.date) from player a1 join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  where a1.player_id_spell1 = a.player_id_spell1 and LFC in (" + LFCvalue + ") " + compcode + ") and dob is not null ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            PlayerM GOPD = new PlayerM();
                            GOPD.rank = (int)reader.GetInt64(0);
                            GOPD.player_id_spell1 = (int)reader.GetInt16(1);
                            GOPD.surname = reader.GetString(2);
                            GOPD.forename = reader.GetString(3);
                            GOPD.initial = reader.GetString(4);
                            GOPD.dob = reader.GetDateTime(5);
                            GOPD.date = reader.GetDateTime(6);
                            GOPD.age = reader.GetInt32(7);
                            OPD.Add(GOPD);
                        }
                    }
                }
                sql = "select top 50 rank() over (order by datediff(day,dob,date) desc) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age from  player a join match_player b on a.player_id = b.player_id where date = (select max(b1.date) from player a1 join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  where a1.player_id_spell1 = a.player_id_spell1 and LFC in (" + LFCvalue + ") " + compcode + ") and dob is not null ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            PlayerM GOPA = new PlayerM();
                            GOPA.rank = (int)reader.GetInt64(0);
                            GOPA.player_id_spell1 = (int)reader.GetInt16(1);
                            GOPA.surname = reader.GetString(2);
                            GOPA.forename = reader.GetString(3);
                            GOPA.initial = reader.GetString(4);
                            GOPA.dob = reader.GetDateTime(5);
                            GOPA.date = reader.GetDateTime(6);
                            GOPA.age = reader.GetInt32(7);
                            OPA.Add(GOPA);
                        }
                    }
                }
                sql = "select top 50 rank() over (order by datediff(day,dob,date)) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age, count(*) as goalcount from  player a join match_goal b on a.player_id = b.player_id where date = (select min(b1.date) from player a1 join match_goal b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  where a1.player_id_spell1 = a.player_id_spell1 and LFC in (" + LFCvalue + ")" + compcode + " )and dob is not null group by player_id_spell1, surname, forename, initials, dob, date ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            PlayerM GYS = new PlayerM();
                            GYS.rank = (int)reader.GetInt64(0);
                            GYS.player_id_spell1 = (int)reader.GetInt16(1);
                            GYS.surname = reader.GetString(2);
                            GYS.forename = reader.GetString(3);
                            GYS.initial = reader.GetString(4);
                            GYS.dob = reader.GetDateTime(5);
                            GYS.date = reader.GetDateTime(6);
                            GYS.age = reader.GetInt32(7);
                            GYS.goalcount = reader.GetInt32(8);
                            YS.Add(GYS);
                        }
                    }
                }
                sql = "select top 50 rank() over (order by datediff(day,dob,date)) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age, count(*) as goalcount from  player a join match_goal b on a.player_id = b.player_id where date = (select min(b1.date) from player a1 join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  where a1.player_id_spell1 = a.player_id_spell1 and LFC in (" + LFCvalue + ")" + compcode + " )and dob is not null group by player_id_spell1, surname, forename, initials, dob, date ";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            PlayerM GYSD = new PlayerM();
                            GYSD.rank = (int)reader.GetInt64(0);
                            GYSD.player_id_spell1 = (int)reader.GetInt16(1);
                            GYSD.surname = reader.GetString(2);
                            GYSD.forename = reader.GetString(3);
                            GYSD.initial = reader.GetString(4);
                            GYSD.dob = reader.GetDateTime(5);
                            GYSD.date = reader.GetDateTime(6);
                            GYSD.age = reader.GetInt32(7);
                            GYSD.goalcount = reader.GetInt32(8);
                            YSD.Add(GYSD);
                        }
                    }
                }
                sql = "select top 50 rank() over (order by datediff(day,dob,date)) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age, count(*) as goalcount from  player a join match_goal b on a.player_id = b.player_id where date = (select min(b1.date) from player a1 join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  where a1.player_id_spell1 = a.player_id_spell1 and LFC in (" + LFCvalue + ") and startpos > 0 " + compcode + ")and dob is not null group by player_id_spell1, surname, forename, initials, dob, date ";
                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            PlayerM GYSSD = new PlayerM();
                            GYSSD.rank = (int)reader.GetInt64(0);
                            GYSSD.player_id_spell1 = (int)reader.GetInt16(1);
                            GYSSD.surname = reader.GetString(2);
                            GYSSD.forename = reader.GetString(3);
                            GYSSD.initial = reader.GetString(4);
                            GYSSD.dob = reader.GetDateTime(5);
                            GYSSD.date = reader.GetDateTime(6);
                            GYSSD.age = reader.GetInt32(7);
                            GYSSD.goalcount = reader.GetInt32(8);
                            YSSD.Add(GYSSD);
                        }
                    }
                }
                sql = "select top 50 rank() over (order by datediff(day,dob,date) desc) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age, count(*) as goalcount from  player a join match_goal b on a.player_id = b.player_id where date = (select max(b1.date) from player a1 join match_goal b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  where a1.player_id_spell1 = a.player_id_spell1 and LFC in (" + LFCvalue + ") " + compcode + ")and dob is not null group by player_id_spell1, surname, forename, initials, dob, date ";
                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            PlayerM GOS = new PlayerM();
                            GOS.rank = (int)reader.GetInt64(0);
                            GOS.player_id_spell1 = (int)reader.GetInt16(1);
                            GOS.surname = reader.GetString(2);
                            GOS.forename = reader.GetString(3);
                            GOS.initial = reader.GetString(4);
                            GOS.dob = reader.GetDateTime(5);
                            GOS.date = reader.GetDateTime(6);
                            GOS.age = reader.GetInt32(7);
                            GOS.goalcount = reader.GetInt32(8);
                            OS.Add(GOS);
                        }
                    }
                }
                //dy inputs with ex
                DateTime date1 = new DateTime(2021, 8, 31);
                DateTime date2 = new DateTime(2006, 7, 28);
                //uncode
                DateTime workdate = new DateTime(date1.Year, date2.Month, date2.Day); //this one is the one with differente year
                string h = (date1 - workdate).TotalDays.ToString();
                // if < 0 year of date date2.year - 1
            }
        }
        public class PlayerM
        {
            public int rank;
            public int player_id_spell1;
            public string surname;
            public string forename;
            public string initial;
            public DateTime dob;
            public DateTime date;
            public int age;
            public int goalcount;
        }
        public class count
        {
            public string letter;
            public int playercount;
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
        public string dy(DateTime date1, DateTime date2)
        {
            DateTime workdate = new DateTime(date1.Year, date2.Month, date2.Day); //this one is the one with differente year
            if ((date1 - workdate).TotalDays < 0)
                workdate = new DateTime(date1.Year - 1, date2.Month, date2.Day);
            return (date1 - workdate).TotalDays.ToString();
        }
        public string yr(DateTime date1, DateTime date2)
        {
            DateTime workdate = new DateTime(date1.Year, date2.Month, date2.Day);
            if ((date1 - workdate).TotalDays < 0)
                return (date1.Year - 1 - date2.Year).ToString();
            return (date1.Year - date2.Year).ToString();
        }
    }
}
