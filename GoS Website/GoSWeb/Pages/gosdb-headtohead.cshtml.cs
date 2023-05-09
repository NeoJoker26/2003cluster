using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Routing;
using System;
using static GoSWeb.Pages.gosdb_misc2Model;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace GoSWeb.Pages
{
    public class gosdb_headtoheadModel : PageModel
    {
        public string heading1;
        public string headtext;
        public string restrictions;
        public string ordered;
        public int N;
        public List<head_to_head> hthtable = new List<head_to_head>();
        public void OnGet()
        {
            N = 0;
            string competition = Request.Query["competition"];
            string homeaway = Request.Query["homeaway"];
            string season_no1 = Request.Query["season1"];
            string season_no2 = Request.Query["season2"];
            string orderby = Request.Query["order"];

            string tableview = "";

            switch (competition)
            {
                case "FLG":
                    tableview = "v_match_FL";
                    heading1 = "Football League";
                    headtext = "all matches in tier 2[Div 2 to 1991, Div 1 to 2003 and the Championship]; tier 3[Div 3 South to 1958, Div 3 to 1991, Div 2 to 2003]; and tier 4[Div 3, 1992 - 2003], all from 1920 to 1939 and 1946 to the present day.";
                    restrictions = "Y";
                    break;
                case "LGS":
                    tableview = "v_match_all_league";
                    heading1 = "All Leagues";
                    headtext = "all matches in all league competitions, including the Southern league [1903-20]; the Western League, a mid-week league including south-east clubs [1903-08]; the Football League [1920-39, 1946-present]; the South West Regional League [1939-40]; and the Football League South [1945-46].";
                    restrictions = "Y";
                    break;
                case "FAC":
                    tableview = "v_match_FA";
                    heading1 = "FA Cup";
                    headtext = "all matches in the FA Cup [1903-present].";
                    restrictions = "Y";
                    break;
                case "CUP":
                    tableview = "v_match_cups";
                    heading1 = "All Cups";
                    headtext = "all matches in all knock-out competitions, including the FA Cup [1903-present]; the Football League War Cup [1940]; the Football League Cup (various sponsors) [1960-present]; the Full Members Cup [1986] (a competition for tiers 1 and 2, also known as the Simod Cup [1987-88] and Zenith Data Systems Cup [1989-91]); the Football League Trophy (a generic name for a competition for tiers 3 and 4, including the Associate Members Cup [1984], the Freight Rovers Trophy [1985-86], the Autoglass Trophy[1993], the Auto Windshields Shield [1994-2000], the LDV Vans Trophy [2000-03] and the Johnstone's Paint Trophy [from 2010]); and official pre-season competitions (the Watney Cup [1973], the Anglo Scottish Cup [1977-79] and the Football League Group Cup [1981]).";
                    restrictions = "Y";
                    break;
                default:
                    tableview = "match";
                    heading1 = "All Competitions";
                    headtext = "all competitive first-team games since the club turned professional, including the Southern League [1903-1920]; the Western League, a mid-week league including south-east clubs [1903-08]; the Football League [1920-39, 1946-present]; the South West Regional League [1939-40]; the Football League South [1945-46]; and all Cup competitions (see the Cup option for details).";
                    break;
            }
            string homeaway_text = "";
            switch (homeaway)
            {
                case "ho":
                    homeaway_text = " and homeaway = 'H'";
                    heading1 += ", home only";
                    restrictions = "Y";
                    break;
                case "ao":
                    homeaway_text = " and homeaway = 'A'";
                    heading1 += ", away only";
                    restrictions = "Y";
                    break;
                default:
                    homeaway_text = "";
                    break;
            }
            if (season_no1 == "" || season_no1 == null)
            {
                season_no1 = "1";
            }
            if (season_no2 == "" || season_no2 == null)
            {
                season_no2 = "112";
            }
            try { 
            if (Int32.Parse(season_no1) < 1)
            {
                restrictions = "Y";
            }
            }catch(Exception e){ 
                restrictions = "Y"; 
            }
            string orderby_text = "";
            switch (orderby)
            {
                case "P":
                    orderby_text = " P DESC, name_now ";
                    ordered = "Y";
                    break;
                case "W":
                    orderby_text = " W DESC, name_now ";
                    heading1 += ", ordered by wins";
                    ordered = "Y";
                    break;
                case "D":
                    orderby_text = " D DESC, name_now ";
                    heading1 += ", ordered by draws";
                    ordered = "Y";
                    break;
                case "L":
                    orderby_text = " L DESC, name_now ";
                    heading1 += ", ordered by defeats";
                    ordered = "Y";
                    break;
                case "F":
                    orderby_text = " F DESC, name_now ";
                    heading1 += ", ordered by goals-for";
                    ordered = "Y";
                    break;
                case "A":
                    orderby_text = " A DESC, name_now ";
                    heading1 += ", ordered by goals-against";
                    ordered = "Y";
                    break;
                case "WP":
                    orderby_text = " WP DESC, name_now ";
                    heading1 += ", ordered by wins per game";
                    ordered = "Y";
                    break;
                case "DP":
                    orderby_text = " DP DESC, name_now ";
                    heading1 += ", ordered by draws per game";
                    ordered = "Y";
                    break;
                case "LP":
                    orderby_text = " LP DESC, name_now ";
                    heading1 += ", ordered by defeats per game";
                    ordered = "Y";
                    break;
                case "FP":
                    orderby_text = " FP DESC, name_now ";
                    heading1 += ", ordered by goals-for per game";
                    ordered = "Y";
                    break;
                case "AP":
                    orderby_text = " AP DESC, name_now ";
                    heading1 += ", ordered by goals-against per game";
                    ordered = "Y";
                    break;
                default:
                    orderby_text = " name_now ";
                    ordered = "";
                    break;
            }
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                //SQL command
                string sql = "select name_now, sum(p) as P, sum(w) as W, sum(d) as D, sum(l) as L, sum(f) as F, sum(a) as A, 1000*sum(w)/sum(p) as WP, 1000*sum(d)/sum(p) as DP, 1000*sum(l)/sum(p) as LP, 1000*sum(f)/sum(p) as FP, 1000*sum(a)/sum(p) as AP from ( select name_now, 1 as p, case when goalsfor > goalsagainst then 1 else 0 end as w, case when goalsfor = goalsagainst then 1 else 0 end as d, case when goalsfor < goalsagainst then 1 else 0 end as l, goalsfor as f, goalsagainst as a  from " + tableview + " join opposition on opposition = name_then join season on date between date_start and date_end where season_no between "+ season_no1 +" and "+season_no2 + homeaway_text +") as subsel group by name_now with rollup order by " + orderby_text;

                //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        //get which row and store it in a list
                        while (reader.Read())
                        {
                            head_to_head table = new head_to_head();
                            try { table.name_now = reader.GetString(0); }catch(Exception e) { table.name_now = ""; }
                            table.P = reader.GetInt32(1);
                            table.W = reader.GetInt32(2);
                            table.D = reader.GetInt32(3);
                            table.L = reader.GetInt32(4);
                            table.F = reader.GetInt32(5);
                            table.A = reader.GetInt32(6);
                            table.WP = reader.GetInt32(7);
                            table.DP = reader.GetInt32(8);
                            table.LP = reader.GetInt32(9);
                            table.FP = reader.GetInt32(10);
                            table.AP = reader.GetInt32(11);
                            hthtable.Add(table);
                        }
                    }
                }

            }
        }
        public class head_to_head
        {
            public string name_now;
            public int P;
            public int W;
            public int D;
            public int L;
            public int F;
            public int A;
            public int WP;
            public int DP;
            public int LP;
            public int FP;
            public int AP;
        }
    }
}
