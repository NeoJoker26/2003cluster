using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using GoSWeb.Model;

namespace GoSWeb.Data
{
    public class ApplicationDbContext: DbContext
    {
        public DbSet<Category> Category { get; set;}
    }
}
