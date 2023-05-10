using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using System.Data.SqlClient;
using System;
using System.Text;
using System.Data;
using System.Windows.Input;

namespace _2003v5.Pages;

public class ManagersModel : PageModel
{
    public List<table> MANAG = new List<table>();
    public List<table> MANAGNO = new List<table>();
    public List<table> MANAGSPE = new List<table>();
    public List<table> combinedList = new List<table>();

    public void OnGet()
    {

        string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";

        using (SqlConnection cnn = new SqlConnection(connetionString))
        {
            cnn.Open();
            {
            }
            string sql = "SELECT manager.manager_id, manager.surname, manager.forename, DATEDIFF(day, manager_spell.from_date, manager_spell.to_date), FORMAT(manager_spell.from_date, 'yyyy-MM-dd'), FORMAT(manager_spell.to_date, 'yyyy-MM-dd'), manager_spell.manager_id1, manager_spell.footnote_no, manager_spell.mynotes FROM manager INNER JOIN manager_spell ON manager.manager_id = manager_spell.manager_id1 ORDER BY manager_spell.from_date ";

            using (SqlCommand command = new SqlCommand(sql, cnn))
            {
                using (SqlDataReader reader = command.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        table display = new table();
                        display.manager_id = reader.GetInt16(0);
                        display.surname = reader.GetString(1);
                        display.forname = reader.GetString(2);

                        if (reader.IsDBNull(3))
                        {
                            display.days= 0;
                        }
                        else
                        {
                            display.days = reader.GetInt32(3);
                        }
                        
                        display.from_date = reader.GetString(4);
                        if (reader.IsDBNull(5))
                        {
                            display.to_date = "N/A";
                        }
                        else
                        {
                            display.to_date = reader.GetString(5);
                        }
                        combinedList.Add(display);

                    }
                }
            }

            //combinedList = MANAG.Union(MANAGNO).Union(MANAGSPE).Where(x => x.manager_id > 0).TakeWhile(x => x.manager_id > 0).ToList();
                
        }
    }


    public class table
    {
        public int manager_id { get; set; }
        public string surname { get; set; }
        public string forname { get; set; }
        public string from_date { get; set; }
        public string to_date { get; set; }
        public int days { get; set; }
        public string footnote { get; set; }
        public string footnote_no { get; set; }
        public string mynotes { get; set; }
        public string penpic { get; set; }
    }

}

