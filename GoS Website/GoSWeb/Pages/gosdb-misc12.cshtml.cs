using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;

namespace GoSWeb.Pages
{
    public class gosdb_misc12Model : PageModel
    {
        public void OnGet()
        {

        }
    }
    public class Appearances
    {
        public int rank;
        public int player_id;
        public string surname;
        public string forename;
        public int consec_count;
        public DateTime start_date;
        public DateTime end_date;
    }

}
