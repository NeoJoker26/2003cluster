using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
namespace _2003v5.Pages
{
    public class SuccessRankingsByOppositionModel : PageModel
    {
        public string competition;
        public string heading;
        public string season1;
        public string season2;
        public List<Success_Ranking> SRHtable = new List<Success_Ranking>(); // Success Ranking at Home
        public List<Success_Ranking> SRAtable = new List<Success_Ranking>();
        public List<Success_Ranking> SRHAtable = new List<Success_Ranking>();
        public void OnGet()
        {
            competition = Request.Query["competition"];
            season1 = Request.Query["season1"];
            season2 = Request.Query["season2"];

            string tableview;
            switch (competition)
            {
                case "FLG":
                    tableview = "v_match_FL";
                    heading = "Football League";
                    break;
                case "CUP":
                    tableview = "v_match_cups";
                    heading = "Cup Competitions";
                    break;
                default:
                    tableview = "v_match_all";
                    heading = "All Competitions";
                    break;
            }
            if (season1 == "" || season1 == null)
                season1 = "1";
            else
                heading = heading + " from " + IndexYear(Int32.Parse(season1));
            if (season2 == "" || season2 == null)
                season2 = (DateTime.Now.Year - 1903 - 1).ToString();//this will select the lastest year (2022-2023 wrote this code in 09/03/2023)							 
            else
                heading = heading + " to " + IndexYear(Int32.Parse(season2));

            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "WITH CTE1 AS ( select homeaway, name_now, cast(sum(points) as dec(7,2))/SUM(p) as pointspergame, sum(p) as p, sum(w) as w, sum(d) as d, sum(l) as l from ( select homeaway, name_now, 1 as p, case when goalsfor > goalsagainst then 1 else 0 end as w, case when goalsfor = goalsagainst then 1 else 0 end as d, case when goalsfor < goalsagainst then 1 else 0 end as l, case when goalsfor > goalsagainst then 3 when goalsfor = goalsagainst then 1 when goalsfor < goalsagainst then 0 end as points from " + tableview + " join opposition on opposition = name_then join season on date between date_start and date_end where homeaway in ('H') and season_no between " + season1 + " and " + season2 + ") as sub group by homeaway, name_now having count(*) > 3 ), CTE2 as ( select rank() over(partition by homeaway order by homeaway, pointspergame desc) as rank, homeaway, pointspergame, name_now, p, w, d, l from CTE1 ) select rank, homeaway, pointspergame, name_now, p, w, d, l from CTE2 order by homeaway desc, rank";
                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Success_Ranking GTable = new Success_Ranking();
                            GTable.rank = (int)reader.GetInt64(0);
                            GTable.homeaway = reader.GetString(1);
                            GTable.pointspergame = reader.GetDecimal(2);
                            GTable.name_now = reader.GetString(3);
                            GTable.p = reader.GetInt32(4);
                            GTable.w = reader.GetInt32(5);
                            GTable.d = reader.GetInt32(6);
                            GTable.l = reader.GetInt32(7);
                            SRHtable.Add(GTable);
                        }
                    }
                }
                sql = "WITH CTE1 AS ( select homeaway, name_now, cast(sum(points) as dec(7,2))/SUM(p) as pointspergame, sum(p) as p, sum(w) as w, sum(d) as d, sum(l) as l from ( select homeaway, name_now, 1 as p, case when goalsfor > goalsagainst then 1 else 0 end as w, case when goalsfor = goalsagainst then 1 else 0 end as d, case when goalsfor < goalsagainst then 1 else 0 end as l, case when goalsfor > goalsagainst then 3 when goalsfor = goalsagainst then 1 when goalsfor < goalsagainst then 0 end as points from " + tableview + " join opposition on opposition = name_then join season on date between date_start and date_end where homeaway in ('A') and season_no between " + season1 + " and " + season2 + ") as sub group by homeaway, name_now having count(*) > 3 ), CTE2 as ( select rank() over(partition by homeaway order by homeaway, pointspergame desc) as rank, homeaway, pointspergame, name_now, p, w, d, l from CTE1 ) select rank, homeaway, pointspergame, name_now, p, w, d, l from CTE2 order by homeaway desc, rank";
                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Success_Ranking GTable = new Success_Ranking();
                            GTable.rank = (int)reader.GetInt64(0);
                            GTable.homeaway = reader.GetString(1);
                            GTable.pointspergame = reader.GetDecimal(2);
                            GTable.name_now = reader.GetString(3);
                            GTable.p = reader.GetInt32(4);
                            GTable.w = reader.GetInt32(5);
                            GTable.d = reader.GetInt32(6);
                            GTable.l = reader.GetInt32(7);
                            SRAtable.Add(GTable);
                        }
                    }
                }
                sql = "WITH CTE1 AS (select name_now, cast(sum(points) as dec(7,2))/SUM(p) as pointspergame, sum(p) as p, sum(w) as w, sum(d) as d, sum(l) as l from (  select name_now, 1 as p,  case when goalsfor > goalsagainst then 1 else 0 end as w,  case when goalsfor = goalsagainst then 1 else 0 end as d,  case when goalsfor < goalsagainst then 1 else 0 end as l,  case when goalsfor > goalsagainst then 3 \r\n when goalsfor = goalsagainst then 1 \r\n when goalsfor < goalsagainst then 0 end as points  from " + tableview + " join opposition on opposition = name_then join season on date between date_start and date_end   and season_no between " + season1 + " and " + season2 + ") as sub group by name_now having count(*) > 6), CTE2 as ( select rank() over(order by pointspergame desc) as rank, pointspergame, name_now, p, w, d, l from CTE1 ) select rank, pointspergame, name_now, p, w, d, l from CTE2 order by rank, pointspergame DESC";
                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            Success_Ranking GTable = new Success_Ranking();
                            GTable.rank = (int)reader.GetInt64(0);
                            GTable.pointspergame = reader.GetDecimal(1);
                            GTable.name_now = reader.GetString(2);
                            GTable.p = reader.GetInt32(3);
                            GTable.w = reader.GetInt32(4);
                            GTable.d = reader.GetInt32(5);
                            GTable.l = reader.GetInt32(6);
                            SRHAtable.Add(GTable);
                        }
                    }
                }

            }
        }
        public string IndexYear(int Index)
        {
            Index--;
            return (1903 + Index).ToString() + "-" + (1904 + Index).ToString();
        }
        public class Success_Ranking
        {
            public int rank;
            public string homeaway;
            public decimal pointspergame;
            public string name_now;
            public int p;
            public int w;
            public int d;
            public int l;
        }
    }
}
