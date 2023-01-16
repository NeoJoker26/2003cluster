using Microsoft.AspNetCore.Mvc;

namespace GoS.Controllers
{
    public class HistoryController : Controller
    {
        public IActionResult History()
        {
            return View("History");
        }
        public IActionResult ShowChapters()
        {
            return View("ShowChapters");
        }
        public IActionResult HowItAllBegan()
        {
            return View("HowItAllBegan");
        }
        public IActionResult ArgyleFCFacts()
        {
            return View("ArgyleFCFacts");
        }

        public IActionResult HistoryMenu()
        {
            return View("HistoryMenu");
        }
    }
}
