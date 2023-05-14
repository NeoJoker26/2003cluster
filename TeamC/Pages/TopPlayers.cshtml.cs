using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Routing;
using System;
using System.Data.SqlClient;
using System.Collections.Generic;
namespace _2003v5.Pages;

public class TopPlayersModel : PageModel
{
    public List<RTC> RTCtable= new List<RTC>();
    public List<RA> RAtable = new List<RA>(); 
    public List<RGS> RGStable = new List<RGS>();
    public List<RGG> RGGtable = new List<RGG>();
    public void OnGet()
    {
        string season_no1 = Request.Query["season1"];
        string season_no2 = Request.Query["season2"];
        string scope = Request.Query["scope"];
        string rank = Request.Query["rank"];

        if (scope == "" || scope == null)       
            scope = "1,2,3,4,5,6,7";
        if (season_no1 == "" || season_no1 == null)
            season_no1 = "1";
        if (season_no2 == "" || season_no2 == null)
            season_no2 = "112";

        string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
        using (SqlConnection cnn = new SqlConnection(connetionString))
        {
            cnn.Open();
            //SQL command
            string sql = "with detailCTE as \r\n( \t\r\n    select spell, a.player_id_spell1, surname, forename, initials, datediff(day,d.date,e.date)+1 as duration \r\n    from player a \r\n    join match_player b on a.player_id = b.player_id \r\n    join match_player c on a.player_id = c.player_id \r\n    join v_match_all d on b.date = d.date join season f on d.date between f.date_start and f.date_end \r\n    join v_match_all e on c.date = e.date join season g on e.date between g.date_start and g.date_end \r\n    and d.date = ( \r\n        select min(b1.date) \r\n        from player a1 \r\n        join match_player b1 on a1.player_id = b1.player_id \r\n        join v_match_all d1 on b1.date = d1.date join season f1 on d1.date between f1.date_start and f1.date_end \r\n        where a1.player_id = a.player_id \r\n        and f1.season_no between "+ season_no1 +" and "+ season_no2 +"\r\n        )\r\n    and e.date = ( \r\n        select max(c2.date) \r\n        from player a2 \r\n        join match_player c2 on a2.player_id = c2.player_id \r\n        join v_match_all e2 on c2.date = e2.date join season g2 on e2.date between g2.date_start and g2.date_end \r\n        where a2.player_id = a.player_id \r\n        and g2.season_no between "+ season_no1 +" and "+ season_no2 +"\r\n    ) \r\n), \r\nsumCTE as \r\n( \t\r\n    select top 100 player_id_spell1, surname, forename, initials, sum(duration) as totduration \r\n    from detailCTE \r\n    group by player_id_spell1, surname, forename, initials \r\n    order by totduration desc, surname \r\n), \r\nspellCTE as \r\n( \r\n     select player_id_spell1, max(spell) as maxspell \r\n     from player \r\n     group by player_id_spell1 \r\n) \r\nselect rank() over (order by totduration desc) as rank, a.player_id_spell1, surname, forename, initials, maxspell, \r\nfloor(totduration/365.25) as years, cast(round(totduration - (365.25*floor(totduration/365.25)),0) as integer) as days \r\nfrom sumCTE a join spellCTE b on a.player_id_spell1 = b.player_id_spell1;   ";

            //also chance it to the login that is not the admin login
            using (SqlCommand command = new SqlCommand(sql, cnn))
            {
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    //get which row and store it in a list
                    while (reader.Read())
                    {
                        RTC table = new RTC();
                        table.rank = (int)reader.GetInt64(0);
                        table.player_id_spell1 = (int)reader.GetInt16(1);
                        table.surname = reader.GetString(2);
                        try { table.forename = reader.GetString(3); } catch { table.forename = "N"; }
                        try { table.initials = reader.GetString(4); } catch { table.initials = "N"; }
                        table.maxspell = reader.GetByte(5);
                        table.years= reader.GetDecimal(6);
                        table.days= reader.GetInt32(7);
                        RTCtable.Add(table);
                    }
                }
            }
            sql = "with detailCTE as \r\n( \t\r\n    select d.player_id_spell1, surname, forename, initials, 1 as starts, 0 as subs \r\n    from v_match_all a join season on date between date_start and date_end \r\n    join match_player b on a.date = b.date \r\n    join player d on b.player_id = d.player_id \r\n    where season_no between "+season_no1+" and "+season_no2+"\r\n     and d.player_id <> 9000 and startpos > 0 \r\n      and a.compcat in ("+ scope +") \r\n      union all \r\n      select d.player_id_spell1, surname, forename, initials, 0 as starts, 1 as subs \r\n      from v_match_all a join season on date between date_start and date_end \r\n      join match_player b on a.date = b.date \r\n      join player d on b.player_id = d.player_id \r\n      where season_no between "+season_no1+" and "+season_no2+"\r\n      and d.player_id <> 9000 and startpos = 0 \r\n      and a.compcat in ("+scope+") \r\n      ), \r\n        sumCTE as \r\n        ( \t\r\n            select top 100 player_id_spell1, surname, forename, initials, sum(starts) as totstarts, sum(subs) as totsubs, sum(starts+subs) as tot \r\n            from detailCTE \r\n            group by player_id_spell1, surname, forename, initials \r\n            order by tot desc, surname \r\n            ) \r\n            select rank() over (order by tot desc) as rank, player_id_spell1, surname, forename, initials, totstarts, totsubs, tot \r\n            from sumCTE ";

            //also chance it to the login that is not the admin login
            using (SqlCommand command = new SqlCommand(sql, cnn))
            {
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    //get which row and store it in a list
                    while (reader.Read())
                    {
                        RA table = new RA();
                        table.rank = (int)reader.GetInt64(0);
                        table.player_id_spell1 = (int)reader.GetInt16(1);
                        table.surname = reader.GetString(2);
                        try { table.forename = reader.GetString(3); } catch { table.forename = "N"; }
                        try { table.initials = reader.GetString(4); } catch { table.initials = "N"; }
                        table.totstarts = reader.GetInt32(5);
                        table.totsubs = reader.GetInt32(6);
                        table.tot = reader.GetInt32(7);
                        RAtable.Add(table);
                    }
                }
            }

            sql = "with CTE as \r\n( \t\r\n    select top 100 player_id_spell1, surname, forename, initials, count(c.player_id) as goals, round(count(c.player_id)/cast(count(distinct b.date) as dec(7,3)),2) as pergame \r\n    from v_match_all a join season on date between date_start and date_end \r\n    join match_player b on a.date = b.date \r\n    left outer join match_goal c on b.player_id = c.player_id and b.date = c.date \r\n    join player d on b.player_id = d.player_id \r\n    where season_no between "+season_no1+" and "+season_no2+"\r\n     and d.player_id <> 9000 \r\n      and a.compcat in ("+scope+") \r\n      group by player_id_spell1, surname, forename, initials \r\n      order by goals desc, surname \r\n) \r\nselect rank() over (order by goals desc) as rank, player_id_spell1, surname, forename, initials, pergame, goals \r\nfrom CTE ";

            //also chance it to the login that is not the admin login
            using (SqlCommand command = new SqlCommand(sql, cnn))
            {
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    //get which row and store it in a list
                    while (reader.Read())
                    {
                        RGS table = new RGS();
                        table.rank = (int)reader.GetInt64(0);
                        table.player_id_spell1 = (int)reader.GetInt16(1);
                        table.surname = reader.GetString(2);
                        try { table.forename = reader.GetString(3); } catch { table.forename = "N"; }
                        try { table.initials = reader.GetString(4); } catch { table.initials = "N"; }
                        table.pergame = reader.GetDecimal(5);
                        table.goals = reader.GetInt32(6);                       
                        RGStable.Add(table);
                    }
                }
            }
            sql = "with CTE as \r\n( \t\r\n    select top 100 player_id_spell1, surname, forename, initials, \r\n    count(c.player_id) as goals, round(count(c.player_id)/cast(count(distinct b.date) as dec(7,3)),3) as pergame \r\n    from v_match_all a join season on date between date_start and date_end \r\n    join match_player b on a.date = b.date \r\n    left outer join match_goal c on b.player_id = c.player_id and b.date = c.date \r\n    join player d on b.player_id = d.player_id \r\n    where season_no between "+season_no1+" and "+season_no2+"\r\n     and d.player_id <> 9000 \r\n     and a.compcat in ("+scope+")\r\n     group by player_id_spell1, surname, forename, initials \r\n     order by pergame desc, surname \r\n     ) \r\n     select rank() over (order by pergame desc) as rank, player_id_spell1, surname, forename, initials, pergame, goals \r\n     from CTE ";

            //also chance it to the login that is not the admin login
            using (SqlCommand command = new SqlCommand(sql, cnn))
            {
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    //get which row and store it in a list
                    while (reader.Read())
                    {
                        RGG table = new RGG();
                        table.rank = (int)reader.GetInt64(0);
                        table.player_id_spell1 = (int)reader.GetInt16(1);
                        table.surname = reader.GetString(2);
                        try { table.forename = reader.GetString(3); } catch { table.forename = "N"; }
                        try { table.initials = reader.GetString(4); } catch { table.initials = "N"; }
                        table.pergame = reader.GetDecimal(5);
                        table.goals = reader.GetInt32(6);
                        RGGtable.Add(table);
                    }
                }
            }

        }

    }
}
public class RTC // Ranked by Time at Club
{
    public int rank;
    public int player_id_spell1;
    public string surname;
    public string forename;
    public string initials;
    public byte maxspell;
    public Decimal years;
    public int days;
}
public class RA // Ranked by Appearances
{
    public int rank;
    public int player_id_spell1;
    public string surname;
    public string forename;
    public string initials;
    public int totstarts;
    public int totsubs;
    public int tot;
}
public class RGS // Ranked by Goals Scored
{
    public int rank;
    public int player_id_spell1;
    public string surname;
    public string forename;
    public string initials;
    public decimal pergame;
    public int goals;

} 
public class RGG //Ranked by Goals per Game
{
    public int rank;
    public int player_id_spell1;
    public string surname;
    public string forename;
    public string initials;
    public decimal pergame;
    public int goals;
}