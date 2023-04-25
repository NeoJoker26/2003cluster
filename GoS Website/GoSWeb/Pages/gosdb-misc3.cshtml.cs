using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static GoSWeb.Pages.gosdb_misc2Model;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using System.Data.SqlClient;
using System;

namespace GoSWeb.Pages
{
    public class gosdb_misc3Model : PageModel
    {
        public List<table> HAHP = new List<table>();
        public List<table> HATR = new List<table>();
        public List<table> LAHP = new List<table>();
        string competition;
        string SN1; //1 -1000
        string SN2; //1-1000
        public void OnGet()
        {
            //reading parameters
            competition = Request.Query["competition"];
            //
            string LFCvalue = "'L','F','C'";
            string heading;
            switch (competition)
            {
                case "FLG":
                    LFCvalue = "'F'";
                    heading = "Football League";
                    break;
                case "CUP":
                    LFCvalue = "'C'";
                    heading = "Cup Competitions";
                    break;
                default:
                    LFCvalue = "'L','F','C'";
                    heading = "All Competitions";
                    break;
            }
            switch (Request.Query["SN1"])
            {
                case null:
                    SN1 = "1";
                    break;
                default:
                    SN1 = Request.Query["SN1"];
                    break;
            }
            switch (Request.Query["SN2"])
            {
                case null:
                    SN2 = "120";
                    break;
                default:
                    SN2 = Request.Query["SN2"];
                    break;
            }
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "select top 50 rank() over (order by attendance desc) as rank, date, attendance, opposition, shortcomp, subcomp, goalsfor, goalsagainst from v_match_all join season on date between date_start and date_end where season_no between " + SN1 + " and " + SN2 + " and LFC in (" + LFCvalue + ") and homeaway = 'H' and attendance is not null \r\n";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            table GHAHP = new table();
                            GHAHP.rank = (int)reader.GetInt64(0);
                            GHAHP.date = reader.GetDateTime(1);
                            GHAHP.attendance = reader.GetInt32(2);
                            GHAHP.opposition = reader.GetString(3);
                            GHAHP.shortcomp = reader.GetString(4);
                            try { GHAHP.subcomp = reader.GetString(5); } catch { GHAHP.subcomp = ""; }
                            GHAHP.goalsfor = (int)reader.GetInt16(6);
                            GHAHP.goalsagainst = (int)reader.GetInt16(7);
                            HAHP.Add(GHAHP);
                        }
                    }
                }
                sql = "select top 50 rank() over (order by attendance desc) as rank, date, attendance, opposition, shortcomp, subcomp, goalsfor, goalsagainst from v_match_all join season on date between date_start and date_end where season_no between " + SN1 + " and " + SN2 + " and LFC in (" + LFCvalue + ") and homeaway <> 'H' and attendance is not null \r\n";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            table GHATR = new table();
                            GHATR.rank = (int)reader.GetInt64(0);
                            GHATR.date = reader.GetDateTime(1);
                            GHATR.attendance = reader.GetInt32(2);
                            GHATR.opposition = reader.GetString(3);
                            GHATR.shortcomp = reader.GetString(4);
                            try { GHATR.subcomp = reader.GetString(5); } catch { GHATR.subcomp = ""; }
                            GHATR.goalsfor = (int)reader.GetInt16(6);
                            GHATR.goalsagainst = (int)reader.GetInt16(7);
                            HATR.Add(GHATR);
                        }
                    }
                }
                sql = "select top 50 rank() over (order by attendance) as rank, date, attendance, opposition, shortcomp, subcomp, goalsfor, goalsagainst from v_match_all join season on date between date_start and date_end where season_no between " + SN1 + " and " + SN2 + " and LFC in (" + LFCvalue + ") and homeaway = 'H' and attendance is not null";

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            table GLAHP = new table();
                            GLAHP.rank = (int)reader.GetInt64(0);
                            GLAHP.date = reader.GetDateTime(1);
                            GLAHP.attendance = reader.GetInt32(2);
                            GLAHP.opposition = reader.GetString(3);
                            GLAHP.shortcomp = reader.GetString(4);
                            try { GLAHP.subcomp = reader.GetString(5); } catch { GLAHP.subcomp = ""; }
                            GLAHP.goalsfor = (int)reader.GetInt16(6);
                            GLAHP.goalsagainst = (int)reader.GetInt16(7);
                            LAHP.Add(GLAHP);
                        }
                    }
                }
            }
        }
        public string Conclusion(int goalsfor, int goalsagainst)
        {
            int R =  goalsfor - goalsagainst;
            switch (R)
            {
                case 0:
                    return "D";
                default:
                    if (R > 0)
                        return "W";
                    else
                        return "L";                  
            }          
        }

        public class table
        {
            public int rank;
            public DateTime date;
            public int attendance;
            public string opposition;
            public string shortcomp;
            public string subcomp;
            public int goalsfor;
            public int goalsagainst;
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
