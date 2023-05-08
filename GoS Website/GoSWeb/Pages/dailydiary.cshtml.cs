using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using static GoSWeb.Pages.IndexModel;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace GoSWeb.Pages
{
    public class dailydiaryModel : PageModel
    {
        public List<DailyDiary> dailydiary = new List<DailyDiary>();
        public void OnGet()
        {

            string archive;
            string selectcriteria = ""; 
            DateTime TodayDate = DateTime.Now;
            if (!string.IsNullOrEmpty(Request.Query["todaydate"]))
            {
                TodayDate = Convert.ToDateTime(Request.Query["todaydate"]);
            }
            archive = Request.Query["archive"];
            if (string.IsNullOrEmpty(archive))
            {
                selectcriteria = "date <= '" + TodayDate.ToString("dd/MM/yyyy") + "' and date >= '1 ";
                if (TodayDate.Month > 1)
                {
                    selectcriteria += DateTime.Now.AddMonths(-1).ToString("MMMM yyyy") + "' "; //For the current diary, "month(TodayDate)-1" forces the start to be the previous month
                }
                else
                {
                    selectcriteria += "Dec" + " " + (TodayDate.Year - 1) + "' "; //For the current diary in January, start from December in the previous year
                }

            }
            else
            {
                selectcriteria = "left(datename(month,date),3) = '" + archive.Substring(0, 3) + "' and year(date) = 20" + archive.Substring(3);
            }
            string connectionString = "Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G"; ///connects to database
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string sql = "set dateformat dmy; select date, entry_no, entry_para_no, entry_para from daily_diary where " + selectcriteria + " order by date desc, entry_no, entry_para_no";
                


                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader()) ///returns on this day data from query to the onThisDayTBL list to be used on the site
                    {
                        while (reader.Read())
                        {
                            DailyDiary dd = new DailyDiary();
                            dd.date = reader.GetDateTime(0);
                            dd.entry_no = reader.GetInt16(1);
                            dd.entry_para_no = reader.GetInt16(2);
                            dd.entry_para = reader.GetString(3);

                            dailydiary.Add(dd);

                        }
                    }
                }
            }
            
            

        }
        public class DailyDiary
        {
            public DateTime date;
            public int entry_no;
            public int entry_para_no;
            public string entry_para;
            

        }
    }
}
