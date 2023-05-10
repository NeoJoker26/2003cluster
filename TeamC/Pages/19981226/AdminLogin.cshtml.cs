using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace GoSWeb.Pages._19981226
{
    public class testModel : PageModel
    {
        public void OnGet()
        {

        }       
        public int Redirect(string Username,string Password)
        {
            Response.Redirect("Check?GoS_administrator=" + Username + "&GoS_administrator_PassWord=" + Password);
            return 0;
        }
    }
    
}
