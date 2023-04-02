using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;
using System.Net;
using Microsoft.AspNetCore;
using System.Web;


namespace GoSWeb.Pages
{
    //this is the back-end were the magic happens
    public class gosdb_mic6Model : PageModel
    {      
        public string comp;
        public string points;
        public string length;
        public string venue;
        public List<Matches> MatchesTable = new List<Matches>();
        public void OnGet()
        {
            //get URL params 
            comp = Request.Query["comp"]; // all or league
            points = Request.Query["points"]; // mp or gd  (modernpoint or goal diference)
            length = Request.Query["length"]; // int is a number 2-30
            venue = Request.Query["venue"]; // H | A | HA    
            //Variable that will stored SQL commands
            string tablename = "";            
            string homeawayclause2;
            string orderby = "";
            int num = 28;
            //were we going to sort the URL params data to SQL
                 
            switch (comp)
            {
                 default:
                    comp = "league";
                    tablename = "[v_match_FL-39]";
                    break;
                case "league":
                    comp = "league";
                    tablename = "[v_match_FL-39]";
                    break;
                case "all":
                    comp = "all";
                    tablename = "match";
                    break;
           
            }
           //lenght                 
            switch (length)
            {
                case null:
                    if (comp == "league")                        
                        num = 28;
                        length = num.ToString();
                    if (comp == "all")
                        num = 30;                       
                    break;
                default:
                    num = Int32.Parse(length);
                    break;
            }          
            switch (venue)
            {
                case "H":                   
                    homeawayclause2 = "where homeaway = 'H' ";
                    break;
                case "A":                    
                    homeawayclause2 = "where homeaway = 'A' ";
                    break;
                case "HA":                    
                    homeawayclause2 = "";
                    break;
                default:                   
                    homeawayclause2 = "";
                    break;
            }
            switch (points)
            {
                case "gd":
                    orderby = "goaldiff DESC";
                    break;
                case "mp":
                    orderby = "modernpoints DESC";
                    break;
                default:
                    orderby = "modernpoints DESC";
                    break;
            }
            
            //get data from sqlserver
            
                string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G"; 
                using (SqlConnection cnn = new SqlConnection(connetionString))
                {
                    cnn.Open();
                    //SQL command
                    string sql = "WITH CTE1 as (select ROW_NUMBER() over (PARTITION by years order by date) as game, date, years, goalsfor, goalsagainst, case when goalsfor > goalsagainst then 1 else 0 end as wins, case when goalsfor = goalsagainst then 1 else 0 end as draws, case when goalsfor < goalsagainst then 1 else 0 end as defeats,case when goalsfor > goalsagainst then 3 when goalsfor = goalsagainst then 1 else 0 end as modernpoints, endpos, promrel from "+ tablename +" a JOIN season e on a.date >= e.date_start and a.date <= e.date_end "+ homeawayclause2 + "),CTE2 as (select years, sum(goalsfor) as goalsfor, sum(goalsagainst) as goalsagainst, sum(goalsfor) - sum(goalsagainst) as goaldiff, sum(wins) as wins,sum(draws) as draws,  sum(defeats) as defeats, sum(modernpoints) as modernpoints, endpos, promrel from CTE1 where game <= "+ num +" group by years, endpos, promrel)select rank() over (order by "+ orderby + ") as rank, years, goalsfor, goalsagainst, goaldiff, wins, draws, defeats, modernpoints, endpos, promrel from CTE2 order by rank, years"; //long string goes... to infinity and beyond  
                    // string must be seperated form the command line under me for better maintenaince 
                    //also chance it to the login that is not the admin login
                using (SqlCommand command = new SqlCommand(sql, cnn))
                    { 
                        using (SqlDataReader reader = command.ExecuteReader())
                        {            
                        //get which row and store it in a list
                            while (reader.Read())
                            {
                                Matches table = new Matches();
                                table.rank = reader.GetInt64(0);                             
                                table.years = reader.GetString(1);
                                table.goalsfor = reader.GetInt32(2);
                                table.goalsagainst = reader.GetInt32(3);
                                table.goaldiff = reader.GetInt32(4);
                                table.wins = reader.GetInt32(5);
                                table.draws = reader.GetInt32(6);
                                table.defeats = reader.GetInt32(7);
                                table.modernpoints = reader.GetInt32(8);
                            try {
                                table.endpos = reader.GetByte(9);
                            }
                            catch (Exception e)
                            {

                            }
                            try
                            {
                                table.promrel = reader.GetString(10);
                            }
                            catch (Exception e)
                            {

                            }                          
                            MatchesTable.Add(table);
                            } 
                        }                                       
                    }
               }
            
            
        }
        //the list colums the same colum type has in the SQL
        public class Matches
        {
            public Int64 rank;
            public string years;
            public int goalsfor;
            public int goalsagainst;
            public int goaldiff;
            public int wins;
            public int draws;
            public int defeats;
            public int modernpoints;
            public byte endpos;
            public string promrel;
        }
    }
}
