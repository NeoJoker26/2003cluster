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
            string Page = Request.Query["Check"];

            if (Checkin(Username,Password) == 0)
            {
                Response.Redirect("AdminMenu?GoS_administrator=" + Username + "&GoS_administrator_PassWord=" + Password);
            }
            else
               Response.Redirect("AdminLogin");


        }
        public int Checkin(string Username, string Password)
        {
            switch (Username)
            {
                case "MiguelRelvas": //delete this
                    if (Password == "123") //GoSadminpassword123
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
        
    }
    
}
