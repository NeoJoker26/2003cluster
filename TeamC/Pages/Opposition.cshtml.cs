using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace _2003v5.Pages;

public class OppositionModel : PageModel
{
    public void OnGet()
    {
    }
    public class OppositionData
    {
        public string competition { get; set; }
        public string game { get; set; }
        public string yearFrom { get; set; }
        public string yearTo { get; set; }
        public string order { get; set; }
    }
   
}