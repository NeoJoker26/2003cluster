using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static GoSWeb.Pages.gosdb_misc2Model;
using System.Data.SqlClient;
using System;
using Microsoft.AspNetCore.Http;

namespace GoSWeb.Pages._19981226
{
    public class add_youtube_linkModel : PageModel
    {
        public string Username;
		public string match_date;
		public int type;
        public void OnGet()
        {
            Username = Request.Query["GoS_administrator"];
			if(Username == "" || Username == null)
			{
				Response.Redirect("AdminLogin");
			}
			string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";

			using (SqlConnection cnn = new SqlConnection(connetionString))
			{
				cnn.Open();
				//SQL command
				string sql = "select max(date) as maxdate \r\nfrom match ";
				DateTime datetime;
				//also chance it to the login that is not the admin login
				using (SqlCommand command = new SqlCommand(sql, cnn))
				{
					using (SqlDataReader reader = command.ExecuteReader())
					{
						
						while (reader.Read())
						{
							datetime = reader.GetDateTime(0);
							match_date = datetime.ToString("dd/MM/yyyy");
						}
					}
				
				}
		}	}
		public int Add_youtube_Link(int T,string Match_date,string youtubelink)
		{
			if(youtubelink == "1")
			{
				return 1;
			}
			else
			{
				string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
				string type = "";
				switch (T)
				{
					case 1:
						type = "action";
						break;
					case 2:
						type = "moments";
						break;
					case 3:
						type = "found";
						break;
				}
				using (SqlConnection cnn = new SqlConnection(connetionString))
				{
					cnn.Open();
					//SQL command
					string sql = "set dateformat ymd; insert into event_control (event_date, event_published, event_type, material_type, material_seq, publish_timestamp, updateno, material_details1, material_details2) values ('" + Match_date + "','Y','M','Y',1,'"+ DateTime.Now + "',99,'"+ youtubelink + "','"+type+"')";

					//also chance it to the login that is not the admin login
					using (SqlCommand command = new SqlCommand(sql, cnn))
					{
						using (SqlDataReader reader = command.ExecuteReader())
						{

							while (reader.Read())
							{
								//nothing here
							}
						}

					}
				}
			}
			return 0;
		}

    }
}
