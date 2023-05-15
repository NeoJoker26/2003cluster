using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Routing;
using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using static _2003v5.Pages.OppositionModel;

namespace _2003v5.Pages;

public class SeasonsModel : PageModel
{
    public string heading1;
    public List<Seasson> Stable = new List<Seasson>();
    public string e1 = "SL+WL";
    public string e2 = "D2 +";
    public string e3 = "SWRL";
    public void OnGet()
    {
        string season_no1 = Request.Query["season1"];
        string season_no2 = Request.Query["season2"];
        string orderby = Request.Query["orderby"];


        if (season_no1 == "" || season_no1 == null)
            season_no1 = "1";
        if (season_no2 == "" || season_no2 == null)
            season_no2 = "112";
        string orderby_text = "";
        string ordered = "";
        switch (orderby)
        {
            case "S":
                orderby_text = " years ";
                ordered = "";
                break;
            case "S-":
                heading1 += ", ordered by seasons, descending";
                orderby_text = " years DESC ";
                ordered = "Y";
                break;
            case "P":
                orderby_text = " P DESC, years ";
                heading1 += ", ordered by games played";
                ordered = "Y";
                break;
            case "W":
                orderby_text = " W DESC, years ";
                heading1 += ", ordered by wins";
                ordered = "Y";
                break;
            case "D":
                orderby_text = " D DESC, years ";
                heading1 += ", ordered by draws";
                ordered = "Y";
                break;
            case "L":
                orderby_text = " L DESC, years ";
                heading1 += ", ordered by defeats";
                ordered = "Y";
                break;
            case "F":
                orderby_text = " F DESC, years ";
                heading1 += ", ordered by goals-for";
                ordered = "Y";
                break;
            case "A":
                orderby_text = " A DESC, years ";
                heading1 += ", ordered by goals-against";
                ordered = "Y";
                break;
            case "PO":
                orderby_text = " PO DESC, years ";
                heading1 += ", ordered by points (3 for a win for all seasons)";
                ordered = "Y";
                break;
            case "HP":
                orderby_text = " HP DESC, years ";
                heading1 += ", ordered by home games played";
                ordered = "Y";
                break;
            case "HW":
                orderby_text = " HW DESC, years ";
                heading1 += ", ordered by home wins";
                ordered = "Y";
                break;
            case "HD":
                orderby_text = " HD DESC, years ";
                heading1 += ", ordered by home draws";
                ordered = "Y";
                break;
            case "HL":
                orderby_text = " HL DESC, years ";
                heading1 += ", ordered by home defeats";
                ordered = "Y";
                break;
            case "HF":
                orderby_text = " HF DESC, years ";
                heading1 += ", ordered by home goals-for";
                ordered = "Y";
                break;
            case "HA":
                orderby_text = " HA DESC, years ";
                heading1 += ", ordered by home goals-against";
                ordered = "Y";
                break;
            case "HPO":
                orderby_text = " HPO DESC, years ";
                heading1 += ", ordered by home points (3 for a win for all seasons)";
                ordered = "Y";
                break;
            case "AP":
                orderby_text = " AP DESC, years ";
                heading1 += ", ordered by away games played";
                ordered = "Y";
                break;
            case "AW":
                orderby_text = " AW DESC, years ";
                heading1 += ", ordered by away wins";
                ordered = "Y";
                break;
            case "AD":
                orderby_text = " AD DESC, years ";
                heading1 += ", ordered by away draws";
                ordered = "Y";
                break;
            case "AL":
                orderby_text = " AL DESC, years ";
                heading1 += ", ordered by away defeats";
                ordered = "Y";
                break;
            case "AF":
                orderby_text = " AF DESC, years ";
                heading1 += ", ordered by away goals-for";
                ordered = "Y";
                break;
            case "AA":
                orderby_text = " AA DESC, years ";
                heading1 += ", ordered by away goals-against";
                ordered = "Y";
                break;
            case "APO":
                orderby_text = " APO DESC, years ";
                heading1 += ", ordered by away points (3 for a win for all seasons)";
                ordered = "Y";
                break;
            case "AT":
                orderby_text = " AT desc, years ";
                heading1 += ", ordered by attendance";
                ordered = "Y";
                break;
            case "HAT":
                orderby_text = " HAT desc, years ";
                heading1 += ", ordered by home attendance";
                ordered = "Y";
                break;
            case "AAT":
                orderby_text = " AAT desc, years ";
                heading1 += ", ordered by away attendance";
                ordered = "Y";
                break;
            case "CP":
                orderby_text = " CP DESC, years ";
                heading1 += ", ordered by cup games played";
                ordered = "Y";
                break;
            case "CW":
                orderby_text = " CW DESC, years ";
                heading1 += ", ordered by cup wins";
                ordered = "Y";
                break;
            case "CD":
                orderby_text = " CD DESC, years ";
                heading1 += ", ordered by cup draws";
                ordered = "Y";
                break;
            case "CL":
                orderby_text = " CL DESC, years ";
                heading1 += ", ordered by cup defeats";
                ordered = "Y";
                break;
            case "CF":
                orderby_text = " CF DESC, years ";
                heading1 += ", ordered by cup goals-for";
                ordered = "Y";
                break;
            case "CA":
                orderby_text = " CA DESC, years ";
                heading1 += ", ordered by cup goals-against";
                ordered = "Y";
                break;
            case "U":
                orderby_text = " player_count DESC, years ";
                heading1 += ", ordered by players used";
                ordered = "Y";
                break;
            case "E":
                orderby_text = " flendpos, years ";
                heading1 += ", ordered by final position in the Football League";
                ordered = "Y";
                break;
            case "CS":
                orderby_text = " clean_sheets DESC, years ";
                heading1 += ", ordered by clean sheets";
                ordered = "Y";
                break;
            default:
                orderby_text = " years ";
                ordered = "";
                break;
        }


        string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
        using (SqlConnection cnn = new SqlConnection(connetionString))
        {
            cnn.Open();
            //SQL command
            string sql = "WITH seasonCTE1 AS ( \r\n  select season_no, years, division_short, tier, endpos, teams_above_div, promrel, sum(cs) as clean_sheets, \r\n  sum(p) as P, sum(w) as W, sum(d) as D, sum(l) as L, sum(f) as F, sum(a) as A, sum(po) as PO, avg(at) as AT, \r\n  sum(hp) as HP, sum(hw) as HW, sum(hd) as HD, sum(hl) as HL, sum(hf) as HF, sum(ha) as HA, sum(hpo) as HPO, avg(hat) as HAT,  \r\n  sum(ap) as AP, sum(aw) as AW, sum(ad) as AD, sum(al) as AL, sum(af) as AF, sum(aa) as AA, sum(apo) as APO, avg(aat) as AAT, \r\n  sum(cp) as CP, sum(cw) as CW, sum(cd) as CD, sum(cl) as CL, sum(cf) as CF, sum(ca) as CA \r\n  from ( \r\n    select season_no, years, division_short, tier, endpos, teams_above_div, promrel, \r\n    case when LFC <> 'C' then 1 else 0 end as p, \r\n    case when LFC <> 'C' and goalsfor > goalsagainst then 1 else 0 end as w, \r\n    case when LFC <> 'C' and goalsfor = goalsagainst then 1 else 0 end as d, \r\ncase when LFC <> 'C' and goalsfor < goalsagainst then 1 else 0 end as l, \r\ncase when LFC <> 'C' then goalsfor else 0 end as f,  \r\ncase when LFC <> 'C' then goalsagainst else 0 end as a,  \r\ncase when LFC <> 'C' and goalsagainst = 0 then 1 else 0 end as cs,  \r\ncase when LFC <> 'C' and goalsfor > goalsagainst then 3 when LFC <> 'C' and goalsfor = goalsagainst then 1 else 0 end as po,  \r\ncase when LFC <> 'C' then attendance else NULL end as at,  \r\ncase when LFC <> 'C' and homeaway = 'H' then 1 else 0 end as hp, \r\ncase when LFC <> 'C' and homeaway = 'H' and goalsfor > goalsagainst then 1 else 0 end as hw, \r\ncase when LFC <> 'C' and homeaway = 'H' and goalsfor = goalsagainst then 1 else 0 end as hd, \r\ncase when LFC <> 'C' and homeaway = 'H' and goalsfor < goalsagainst then 1 else 0 end as hl, \r\ncase when LFC <> 'C' and homeaway = 'H' then goalsfor else 0 end as hf, \r\ncase when LFC <> 'C' and homeaway = 'H' then goalsagainst else 0 end as ha, \r\ncase when LFC <> 'C' and homeaway = 'H' and goalsfor > goalsagainst then 3 when LFC <> 'C' and homeaway = 'H' and goalsfor = goalsagainst then 1 else 0 end as hpo,  \r\ncase when LFC <> 'C' and homeaway = 'H' then attendance else NULL end as hat, \r\ncase when LFC <> 'C' and homeaway = 'A' then 1 else 0 end as ap, \r\ncase when LFC <> 'C' and homeaway = 'A' and goalsfor > goalsagainst then 1 else 0 end as aw, \r\ncase when LFC <> 'C' and homeaway = 'A' and goalsfor = goalsagainst then 1 else 0 end as ad, \r\ncase when LFC <> 'C' and homeaway = 'A' and goalsfor < goalsagainst then 1 else 0 end as al, \r\ncase when LFC <> 'C' and homeaway = 'A' then goalsfor else 0 end as af, \r\ncase when LFC <> 'C' and homeaway = 'A' then goalsagainst else 0 end as aa, \r\ncase when LFC <> 'C' and homeaway = 'A' and goalsfor > goalsagainst then 3 when LFC <> 'C' and homeaway = 'A' and goalsfor = goalsagainst then 1 else 0 end as apo,  \r\ncase when LFC <> 'C' and homeaway = 'A' then attendance else NULL end as aat, \r\ncase when LFC = 'C' then 1 else 0 end as cp, \r\ncase when LFC = 'C' and goalsfor > goalsagainst then 1 else 0 end as cw, \r\ncase when LFC = 'C' and goalsfor = goalsagainst then 1 else 0 end as cd, \r\ncase when LFC = 'C' and goalsfor < goalsagainst then 1 else 0 end as cl, \r\ncase when LFC = 'C' then goalsfor else 0 end as cf, \r\ncase when LFC = 'C' then goalsagainst else 0 end as ca \r\nfrom v_match_all join season on date between date_start and date_end \r\nwhere season_no between " + season_no1 + " and " + season_no2 + " \r\n) as subsel \r\ngroup by season_no, years, division_short, tier, endpos, teams_above_div, promrel \r\n), \r\nseasonCTE2 AS ( \r\n  select season_no, count(distinct player_id) as player_count \r\n  from match_player join season on date between date_start and date_end \r\n  where season_no between " + season_no1 + " and " + season_no2 + " \r\n  group by season_no \r\n  ) \r\nselect player_count, years, division_short, tier, endpos, endpos + teams_above_div as flendpos, promrel, clean_sheets, \r\nP, W, D, L, F, A, PO, AT, HP, HW, HD, HL, HF, HA, HPO, HAT, AP, AW, AD, AL, AF, AA, APO, AAT, CP, CW, CD, CL, CF, CA \r\nfrom seasonCTE1 x join seasonCTE2 y on x.season_no = y.season_no \r\norder by " + orderby_text;

            //also chance it to the login that is not the admin login
            using (SqlCommand command = new SqlCommand(sql, cnn))
            {
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    //get which row and store it in a list
                    while (reader.Read())
                    {
                        Seasson table = new Seasson();
                        table.player_count = reader.GetInt32(0);
                        table.years = reader.GetString(1);
                        try { table.division_short = reader.GetString(2); } catch { }
                        try { table.tier = reader.GetInt32(3); } catch { }
                        try { table.endpos = reader.GetByte(4); } catch { }
                        try { table.flendpos = reader.GetInt32(5); } catch { }
                        try { table.promrel = reader.GetString(6); } catch { table.promrel = ""; }
                        table.clean_sheets = reader.GetInt32(7);
                        table.P = reader.GetInt32(8);
                        table.W = reader.GetInt32(9);
                        table.D = reader.GetInt32(10);
                        table.L = reader.GetInt32(11);
                        table.F = reader.GetInt32(12);
                        table.A = reader.GetInt32(13);
                        table.PO = reader.GetInt32(14);
                        try { table.AT = reader.GetInt32(15); } catch { }
                        table.HP = reader.GetInt32(16);
                        table.HW = reader.GetInt32(17);
                        table.HD = reader.GetInt32(18);
                        table.HL = reader.GetInt32(19);
                        table.HF = reader.GetInt32(20);
                        table.HA = reader.GetInt32(21);
                        table.HPO = reader.GetInt32(22);
                        try { table.HAT = reader.GetInt32(23); } catch { }
                        table.AP = reader.GetInt32(24);
                        table.AW = reader.GetInt32(25);
                        table.AD = reader.GetInt32(26);
                        table.AL = reader.GetInt32(27);
                        table.AF = reader.GetInt32(28);
                        table.AA = reader.GetInt32(29);
                        table.APO = reader.GetInt32(30);
                        try { table.AAT = reader.GetInt32(31); } catch { }
                        table.CP = reader.GetInt32(32);
                        table.CW = reader.GetInt32(33);
                        table.CD = reader.GetInt32(34);
                        table.CL = reader.GetInt32(35);
                        table.CF = reader.GetInt32(36);
                        table.CA = reader.GetInt32(37);
                        Stable.Add(table);
                    }
                }
            }

        }



    }

}
public class Seasson
{
    public int player_count;
    public string years;
    public string division_short;
    public int tier;
    public byte endpos;
    public int flendpos;
    public string promrel;
    public int clean_sheets;
    public int P;
    public int W;
    public int D;
    public int L;
    public int F;
    public int A;
    public int PO;
    public int AT;
    public int HP;
    public int HW;
    public int HD;
    public int HL;
    public int HF;
    public int HA;
    public int HPO;
    public int HAT;
    public int AP;
    public int AW;
    public int AD;
    public int AL;
    public int AF;
    public int AA;
    public int APO;
    public int AAT;
    public int CP;
    public int CW;
    public int CD;
    public int CL;
    public int CF;
    public int CA;
}