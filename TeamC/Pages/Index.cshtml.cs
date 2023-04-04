using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;
namespace _2003v5.Pages;

public class IndexModel : PageModel
{
    public List<onThisDay> onThisDayTBL = new List<onThisDay>();
    public List<Month> Monthli = new List<Month>();
    public void OnGet()
    {
        String monthChar = DateTime.Now.ToString("MMM");
        Console.WriteLine(monthChar);


        string connectionString = "Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();
            string sql = "select year, fact from onthisday where month = '" + monthChar + "'and day = '" + DateTime.Now.Day + "'and seqno < 99 order by seqno";


            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                using (SqlDataReader reader = command.ExecuteReader())
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
    public class Month
    {
        public string month;
        public int day;
    }
}