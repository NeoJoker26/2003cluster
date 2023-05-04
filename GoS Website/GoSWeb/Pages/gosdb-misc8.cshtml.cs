using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static GoSWeb.Pages.gosdb_misc5Model;
using System.Data.SqlClient;
using System.Collections.Generic;
using System;

namespace GoSWeb.Pages
{
    public class gosdb_misc8Model : PageModel
    {
        public string yearcount = "1";
        public string homeaway;
        public string bycolumn;
        public string bestworst;
        public string heading1;
        public string heading2;
        public List<LR> LRTable = new List<LR>();
        public void OnGet()
        {
            yearcount = Request.Query["yearcount"];
            homeaway = Request.Query["homeaway"];
            bycolumn = Request.Query["bycolumn"];
            bestworst = Request.Query["bestworst"];
            //
            string homeawayclause1 = ""; //delete this later
            string homeawayclause2 = "";
            heading1 = "";
            heading2 = "";
            string orderby = "";
            if (yearcount == null)
                yearcount = "1";
            switch (homeaway)
            {
                case "ho":
                    homeawayclause1 = "and homeaway = 'H' ";
                    homeawayclause2 = "where homeaway = 'H' ";
                    heading2 = "for Home League Matches only";
                    break;
                case "ao":
                    homeawayclause1 = "and homeaway = 'A' ";
                    homeawayclause2 = "where homeaway = 'A' ";
                    heading2 = "for Away League Matches only";
                    break;
                default:
                    homeawayclause1 = "";
                    homeawayclause2 = "";
                    heading2 = "for all League Matches";
                    break;
            }
            switch (bycolumn)
            {
                case "diff":
                    orderby = "goaldiff";
                    break;
                case "wins":
                    orderby = "wins";
                    break;
                case "defeats":
                    orderby = "defeats";
                    break;
                default:
                    orderby = "modernpoints";
                    break;
            }
            switch (bestworst)
            {
                case "worst":
                    orderby = orderby + " asc";
                    break;
                default:
                    orderby = orderby + " desc";
                    break;
            }
            if (yearcount == "1")
                heading1 = "Figures accumulated over a Single Calendar Year";
            else
                heading1 = "Figures accumulated over " + yearcount + " Calendar Years";

            int yearcountm1 = Int32.Parse(yearcount) - 1;
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "WITH CTE1 as \r\n(select distinct cast(year(date) as varchar) + '-' + cast(year(date) + "+yearcountm1+" as varchar) as years, year(date) as startyear, year(date) + "+ yearcountm1 + " as endyear\r\nfrom [v_match_FL-39] a \r\nwhere year(date) <> 1920 \r\n  and year(date) <> 1946 \r\n    and year(date) + "+ yearcountm1 + " <  GETDATE()\r\n      and year(date) + "+ yearcountm1 + " not between 1939 and 1946 \r\n), \r\nCTE2 as \r\n(select years, 1 as played, goalsfor, goalsagainst, \r\ncase when goalsfor > goalsagainst then 1 else 0 end as wins, \r\ncase when goalsfor = goalsagainst then 1 else 0 end as draws, \r\ncase when goalsfor < goalsagainst then 1 else 0 end as defeats, \r\ncase when goalsfor > goalsagainst then 3 \r\n\t when goalsfor = goalsagainst then 1 \r\n     else 0 \r\n\t end as modernpoints \r\n     from [v_match_FL-39] join CTE1 on year(date) between startyear and endyear \r\n"+ homeawayclause2 +"\r\n),\r\nCTE3 as\r\n(select years, sum(played) as played, sum(goalsfor) as goalsfor, sum(goalsagainst) as goalsagainst, \r\n\t\t\t  sum(goalsfor) - sum(goalsagainst) as goaldiff, \r\n\t\t\t  sum(wins) as wins, sum(draws) as draws, sum(defeats) as defeats, \r\n\t\t\t  sum(modernpoints) as modernpoints\r\nfrom CTE2 \r\n group by years \r\n )  \r\nselect rank() over (order by " + orderby + ") as rank, years, played, goalsfor, goalsagainst, goaldiff, wins, draws, defeats, modernpoints from CTE3 order by rank, years";
                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            LR GLRTable = new LR();
                            GLRTable.rank = (int)reader.GetInt64(0);                         
                            string[] yearsarray = reader.GetString(1).Split('-');
                            GLRTable.years = yearsarray[0];
                            GLRTable.played = reader.GetInt32(2);
                            GLRTable.goalsfor = reader.GetInt32(3);
                            GLRTable.goalsagainst = reader.GetInt32(4);
                            GLRTable.goaldiff = reader.GetInt32(5);
                            GLRTable.wins = reader.GetInt32(6);
                            GLRTable.draws = reader.GetInt32(7);
                            GLRTable.defeats = reader.GetInt32(8);
                            GLRTable.mordenpoints = reader.GetInt32(9);
                            decimal n1 = reader.GetInt32(9);
                            decimal n2 = reader.GetInt32(2);
                            GLRTable.MPG = n1 / n2;
                            LRTable.Add(GLRTable);
                        }
                    }
                }

            }
        }
        public class LR //League Results
        {
            public int rank;
            public string years;
            public int played;
            public int goalsfor;
            public int goalsagainst;
            public int goaldiff;
            public int wins;
            public int draws;
            public int defeats;
            public int mordenpoints;
            public decimal MPG;
        } 
    }
}
