using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;
using System;
using System.Text.RegularExpressions;
using static GoSWeb.Pages.IndexModel;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Security.Permissions;
using System.Linq;

namespace GoSWeb.Pages //Jonathan's Code
{
    public class progresstablesModel : PageModel
    {
        public SqlLists sqlLists = new SqlLists();

        public class SqlLists
        {
            public int target_season = 118;
            public string target_date = DateTime.Now.ToString("yyyyMM-DD");
            public List<maxdate> Maxdate = new List<maxdate>();
            public List<SeasonData> seasondata = new List<SeasonData>();
            public List<seasonsByYear> SeasonsByYear = new List<seasonsByYear>();
            public List<MatchData> matchData = new List<MatchData>();
            public int CurrentSeason;


            
        }

            public void OnGet()
            {
            sqlLists.CurrentSeason = sqlLists.SeasonsByYear.OrderByDescending(season => season.season_no)
                                       .Select(season => season.season_no)
                                       .LastOrDefault();
            string connectionString = "Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G"; ///connects to database
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string sql = "select max(date) as maxdate from match a join competition b on a.compcode = b.compcode where lfc = 'F'"; ///query to run on database 


                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader()) ///returns data from query to the maxdate list to be used on the site
                        {
                            while (reader.Read())
                            {
                                maxdate d = new maxdate();
                                d.date = reader.GetDateTime(0);


                                sqlLists.Maxdate.Add(d);

                            }
                        }
                    }
                }

                using (SqlConnection connection2 = new SqlConnection(connectionString))
                {
                    connection2.Open();//possible queries to run on database depending on conditions met

                    string sql = "";
                    if (sqlLists.target_season > 0)
                    {
                        sql = "select season_no, years, division, date_start, date_end, endpos, tier, teams_in_div, pos_promote, pos_promote_playoff, pos_relegate_playoff, pos_relegate from season where season_no =" + sqlLists.target_season;
                    }
                    else if (sqlLists.target_date != null)
                    {
                        sql = "select season_no, years, division, date_start, date_end, endpos, tier, teams_in_div, pos_promote, pos_promote_playoff, pos_relegate_playoff, pos_relegate from season where date_start <= '" + sqlLists.target_date + " ' and date_end >= '" + sqlLists.target_date + "' ";
                    }
                    else
                    {
                        sql = "select season_no, years, division, date_start, date_end, endpos, tier, teams_in_div, pos_promote, pos_promote_playoff, pos_relegate_playoff, pos_relegate from season where season_no = (select max(season_no) from season) ";
                    }

                    using (SqlCommand command = new SqlCommand(sql, connection2))
                    {
                        using (SqlDataReader reader = command.ExecuteReader()) //returns data from query to the season data list to be used on the site
                        {
                            while (reader.Read())
                            {
                                SeasonData sd = new SeasonData();
                                sd.season_no = reader.GetByte(0);
                                sd.years = reader.GetString(1);
                                sd.division = reader.GetString(2);
                                sd.date_start = reader.GetDateTime(3);
                                sd.date_end = reader.GetDateTime(4);
                                sd.endpos = reader.GetByte(5);
                                sd.tier = reader.GetByte(6);
                                sd.teams_in_div = reader.GetByte(7);
                                try { sd.pos_promote = reader.GetByte(8); } //Catching null values
                                catch { }
                                try { sd.pos_promote_playoff = reader.GetByte(9); }
                                catch { }
                                try { sd.pos_relegate_playoff = reader.GetByte(10); }
                                catch { }
                                try { sd.pos_relegate = reader.GetByte(11); }
                                catch { }


                                sqlLists.seasondata.Add(sd);

                            }
                        }
                    }
                }
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string sql = "select distinct season_no, years from v_match_season where year(date_start) >= 1920 and year(date_start) <> 1945 and totpoints is not null order by years "; ///query to run on database 


                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader()) ///returns data from query to the maxdate list to be used on the site
                        {
                            while (reader.Read())
                            {
                                seasonsByYear sby = new seasonsByYear();
                                sby.season_no = reader.GetByte(0);
                                sby.years = reader.GetString(1);

                                sqlLists.SeasonsByYear.Add(sby);

                            }
                        }
                    }
                }
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string sql = "";
                    if (sqlLists.target_season == sqlLists.CurrentSeason)
                    {
                        sql = "select a.date, a.opposition, name_abbrev, lfc, homeaway, 'Home Park', NULL, NULL, NULL, NULL, competition, NULL from season_this a join competition b on a.compcode = b.compcode join opposition c on a.opposition = c.name_then where not exists (select * from match c where c.date = a.date) and homeaway = 'H' and lfc = 'F' and right(rtrim(shortcomp),2) <> 'PO' and date between '2022-07-30' and '2023-05-28' union all select a.date, a.opposition, name_abbrev, lfc, homeaway, ground_name, NULL, NULL, NULL, NULL, competition, NULL from season_this a join competition b on a.compcode = b.compcode join opposition c on a.opposition = c.name_then join venue d on a.opposition = d.club_name_then and a.date between d.first_game and d.last_game where not exists (select * from match c where c.date = a.date) and homeaway in ('A', 'N') and lfc = 'F' and right(rtrim(shortcomp),2) <> 'PO' and date between '2022-07-30' and '2023-05-28'"; ///query to run on database 

                    }
                    else
                    {
                        sql = "select date, opposition, NULL as name_abbrev, lfc, homeaway, 'Home Park' as ground_name, goalsfor, goalsagainst, totpoints, position, competition, attendance from v_match_season a where season_no =  112 and homeaway = 'H' and lfc = 'F' and right(rtrim(rtrim(shortcomp)),2) <> 'PO' union all select date, opposition, NULL, lfc, homeaway, ground_name, goalsfor, goalsagainst, totpoints, position, competition, attendance from v_match_season a join venue b on a.opposition = b.club_name_then and a.date between b.first_game and b.last_game where season_no =  112 and homeaway in ('A', 'N') and lfc = 'F' and right(rtrim(shortcomp),2) <> 'PO' union all select a.date, a.opposition, name_abbrev, lfc, homeaway, 'Home Park', NULL, NULL, NULL, NULL, competition, NULL from season_this a join competition b on a.compcode = b.compcode join opposition c on a.opposition = c.name_then where not exists (select * from match c where c.date = a.date) and homeaway = 'H' and lfc = 'F' and right(rtrim(shortcomp),2) <> 'PO' and date between '2022-07-30' and '2023-05-28'union all select a.date, a.opposition, name_abbrev, lfc, homeaway, ground_name, NULL, NULL, NULL, NULL, competition, NULL from season_this a join competition b on a.compcode = b.compcode  join opposition c on a.opposition = c.name_then join venue d on a.opposition = d.club_name_then and a.date between d.first_game and d.last_game where not exists (select * from match c where c.date = a.date) and homeaway in ('A', 'N') and lfc = 'F' and right(rtrim(shortcomp),2) <> 'PO' and date between '2022-07-30' and '2023-05-28'order by lfc desc, date";
                    }

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader()) ///returns data from query to the maxdate list to be used on the site
                        {
                            if (sqlLists.target_season == sqlLists.CurrentSeason)
                            {
                                while (reader.Read())
                                {
                                    MatchData md = new MatchData();
                                    md.date = reader.GetDateTime(0);
                                    md.opposition = reader.GetString(1);
                                    md.name_abbrev = reader.GetString(2);
                                    md.lfc = reader.GetString(3);
                                    md.homeaway = reader.GetString(4);
                                    md.ground_name = reader.GetString(5);
                                    md.competition = reader.GetString(10);

                                    sqlLists.matchData.Add(md);

                                }
                            }
                            else
                            {
                                while (reader.Read())
                                {
                                    MatchData md = new MatchData();
                                    md.date = reader.GetDateTime(0);
                                    md.opposition = reader.GetString(1);
                                    try { md.name_abbrev = reader.GetString(2); }
                                    catch { }
                                    md.lfc = reader.GetString(3);
                                    md.homeaway = reader.GetString(4);
                                    md.ground_name = reader.GetString(5);
                                    try { md.goalsfor = reader.GetInt16(6); }
                                    catch { }
                                    try { md.goalsagainst = reader.GetInt16(7); }
                                    catch { }
                                    try { md.totpoints = reader.GetByte(8); }
                                    catch { }
                                    try { md.position = reader.GetByte(9); }
                                    catch { }
                                    md.competition = reader.GetString(10);
                                    try { md.attendance = reader.GetInt32(11); }
                                    catch { }
                                

                                    sqlLists.matchData.Add(md);

                                }
                            }
                            }
                           
                        }
                    }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string sql = "SELECT MAX(season_no) FROM v_match_season"; ///query to run on database 


                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader()) ///returns data from query to the maxdate list to be used on the site
                    {
                        while (reader.Read())
                        {
                            sqlLists.CurrentSeason = reader.GetByte(0);

                        }
                    }
                }
            }
        }
            }

            public class maxdate
            {
                public DateTime date;
            }
            public class SeasonData
            {
                public int season_no;
                public string years;
                public string division;
                public DateTime date_start;
                public DateTime date_end;
                public int endpos;
                public int tier;
                public int teams_in_div;
                public int pos_promote;
                public int pos_promote_playoff;
                public int pos_relegate_playoff;
                public int pos_relegate;
            }
            public class seasonsByYear
            {
                public int season_no;
                public string years;
            }
            public class MatchData
            {
                public DateTime date;
                public string opposition;
                public string name_abbrev;
                public string lfc;
                public string homeaway;
                public string ground_name;
                public int goalsfor;
                public int goalsagainst;
                public int totpoints;
                public int position;
                public string competition;
                public int attendance;

        
    }
        }
    

