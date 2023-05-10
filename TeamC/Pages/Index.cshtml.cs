using System.Data.SqlClient;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;


namespace _2003v5.Pages;

public class IndexModel : PageModel
{
    public List<onThisDay> onThisDayTBL = new List<onThisDay>();
    public List<bornThisDay> bornThisDayTBL = new List<bornThisDay>();
    public List<Month> Monthli = new List<Month>();
    public void OnGet()
    {
        String monthChar = DateTime.Now.ToString("MMM");



        string connectionString = "Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G"; ///connects to database
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();
            string sql = "select year, fact from onthisday where month = '" + monthChar + "'and day = '" + DateTime.Now.Day + "'and seqno < 99 order by seqno"; ///On this day query to run on database 


            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                using (SqlDataReader reader = command.ExecuteReader()) ///returns on this day data from query to the onThisDayTBL list to be used on the site
                {
                    while (reader.Read())
                    {
                        onThisDay otd = new onThisDay();
                        otd.year = reader.GetInt16(0);
                        otd.fact = reader.GetString(1);

                        onThisDayTBL.Add(otd);

                    }
                }
            }
        }
        using (SqlConnection connection2 = new SqlConnection(connectionString))
        {
            connection2.Open();
            string sql = "select a.player_id, a.forename, a.surname, year(a.dob) as year, a.first_game_year, max(b.last_game_year) as last_game_year, left(a.penpic,160) as penpic, a.prime_photo from player a left outer join player b on a.player_id = b.player_id_spell1 where month(a.dob) = '" + DateTime.Now.Month + "' and day(a.dob) = '" + DateTime.Now.Day + "' and a.spell = 1  group by a.player_id, a.forename, a.surname, a.dob, a.first_game_year, a.penpic, a.prime_photo order by a.dob ";


            using (SqlCommand command = new SqlCommand(sql, connection2))
            {
                using (SqlDataReader reader = command.ExecuteReader()) ///returns born this day data from query to the bornThisDayTBL list to be used on the site
                {
                    while (reader.Read())
                    {
                        bornThisDay btd = new bornThisDay();
                        btd.forename = reader.GetString(1);
                        btd.surname = reader.GetString(2);
                        btd.year = reader.GetInt32(3);
                        btd.first_game_year = reader.GetInt16(4);
                        btd.last_game_year = reader.GetInt16(5);
                        btd.penpic = reader.GetString(6);
                        try
                        {
                            btd.prime_photo = reader.GetInt16(7);
                        }
                        catch { }


                        bornThisDayTBL.Add(btd);

                    }
                }
            }
        }


    }


    public class onThisDay
    {
        public int id;
        public char month;
        public int day;
        public int year;
        public int seqno;
        public string fact;

    }
    public class bornThisDay
    {
        public int player_id;
        public string forename;
        public string surname;
        public int year;
        public int first_game_year;
        public int last_game_year;
        public string penpic;
        public int prime_photo;

    }
    public class Month
    {

        public string month;
        public int day;

    }

}