using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace GoSWeb.Pages._19981226
{
    public class CheckModel : PageModel
    {
        public void OnGet()
        {
            string Username = Request.Query["GoS_administrator"];
            string Password = Request.Query["GoS_administrator_PassWord"];
            string Page = Request.Query["page"];

            if (Checkin(Username,Password) == 0)
            {
                Page = WebPage(Page);

                if (Page == "1")
                    Response.Redirect("AdminLogin");

                Response.Redirect(Page + "?GoS_administrator=" + Username + "&GoS_administrator_PassWord=" + Password);
            }
            else
               Response.Redirect("AdminLogin");


        }
        public int Checkin(string Username, string Password)
        {
            
  
            switch (Username)
            {
                case "MiguelRelvas":
                    if (Password == "GoSadminpassword123")
                        return 0;                                       
                    break;
                case "MattEllacott":
                    break;
                case "KeithGreening":
                    break;
                case "SteveDean":
                    break;
                case "AlecHepburn":
                    break;
                default:
                    return 1;                 
            }
            return 1;
        }
        public string WebPage(string id)
        {
            switch (id)
            {
                case "12934756":
                    return "AdminMenu";
                default:                
                    return "1";                
            }
        }
    }
    
}
