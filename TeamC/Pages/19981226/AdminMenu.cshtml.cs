using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace GoSWeb.Pages._19981226
{
    public class AdminMenuModel : PageModel
    {
        public string Name;
        public string Password;
        public string fullname;
        public void OnGet()
        {
            Name = Request.Query["GoS_administrator"];
            Password = Request.Query["GoS_administrator_PassWord"];
            fullname = "";

            if(Name == null || Name == "" || Name.Length == 0)
                Response.Redirect("AdminLogin");
           
            for (int i = 0; i < Name.Length; i++)
            {
                if (i > 0 && char.IsUpper(Name[i]))
                {
                    fullname += " ";
                }
                fullname += Name[i];
            }
            
            
            
        }
        public int add_youtube_link()
        {
            Response.Redirect("add_youtube_link?GoS_administrator=" + Name + "&GoS_administrator_PassWord=" + Password + "&&Check=add_youtube_link");
            return 0;
        }
    }
}
