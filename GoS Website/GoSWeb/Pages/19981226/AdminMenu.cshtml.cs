using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace GoSWeb.Pages._19981226
{
    public class AdminMenuModel : PageModel
    {
        public string Name ;
        public string fullname;
        public void OnGet()
        {
            Name = Request.Query["GoS_administrator"];
            fullname = "";
            for (int i = 0; i < Name.Length; i++)
            {
                if (i > 0 && char.IsUpper(Name[i]))
                {
                    fullname += " ";
                }
                fullname += Name[i];
            }
        }
        
    }
}
