using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;


namespace _2003v5.Pages
{
    public class CompetitionTotalsModel : PageModel
    {
        public List<Matches> MatchesTable = new List<Matches>();

        public void OnGet()
        {
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                string sql = "WITH CTE1 AS ( select case LFC when 'F' then 'League' when 'L' then 'Non-league' else 'Cup' end as comptype1, case LFC when 'F' then 'League' + ' tier ' + cast(tier as varchar) when 'L' then 'Non-league' else 'Cup' end as comptype2, competition, 1 as p, case when goalsfor > goalsagainst then 1 else 0 end as w, case when goalsfor = goalsagainst then 1 else 0 end as d, case when goalsfor < goalsagainst then 1 else 0 end as l, goalsfor as f, goalsagainst as a, attendance at, case when homeaway = 'H' then 1 else 0 end as hp, case when homeaway = 'H' and goalsfor > goalsagainst then 1 else 0 end as hw, case when homeaway = 'H' and goalsfor = goalsagainst then 1 else 0 end as hd, case when homeaway = 'H' and goalsfor < goalsagainst then 1 else 0 end as hl, case when homeaway = 'H' then goalsfor else 0 end as hf, case when homeaway = 'H' then goalsagainst else 0 end as ha, case when homeaway = 'H' then attendance else NULL end as hat, case when homeaway <> 'H' then 1 else 0 end as ap, case when homeaway <> 'H' and goalsfor > goalsagainst then 1 else 0 end as aw, case when homeaway <> 'H' and goalsfor = goalsagainst then 1 else 0 end as ad, case when homeaway <> 'H' and goalsfor < goalsagainst then 1 else 0 end as al, case when homeaway <> 'H' then goalsfor else 0 end as af, case when homeaway <> 'H' then goalsagainst else 0 end as aa, case when homeaway <> 'H' then attendance else NULL end as aat from v_match_all join season on date between date_start and date_end ) select case when grouping(comptype1) = 1 then 'zzz' else comptype1 end as comptype1, case when grouping(comptype2) = 1 then 'zzzz' else comptype2 end as comptype2, case when grouping(competition) = 1 then 'zzzzz' else competition end as competition, sum(p) as P, sum(w) as W, sum(d) as D, sum(l) as L, sum(f) as F, sum(a) as A, avg(at) as AT, sum(hp) as HP, sum(hw) as HW, sum(hd) as HD, sum(hl) as HL, sum(hf) as HF, sum(ha) as HA, avg(hat) as HAT, sum(ap) as AP, sum(aw) as AW, sum(ad) as AD, sum(al) as AL, sum(af) as AF, sum(aa) as AA, avg(aat) as AAT from CTE1 group by comptype1, comptype2, competition with rollup order by comptype1, comptype2, competition ";
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Matches table = new Matches();
                            table.comptype1 = reader.GetString(0);
                            table.comptype2 = reader.GetString(1);
                            table.competition = reader.GetString(2);
                            table.P = reader.GetInt32(3);
                            table.W = reader.GetInt32(4);
                            table.D = reader.GetInt32(5);
                            table.L = reader.GetInt32(6);
                            table.F = reader.GetInt32(7);
                            table.A = reader.GetInt32(8);
                            table.AT = reader.GetInt32(9);
                            table.HP = reader.GetInt32(10);
                            table.HW = reader.GetInt32(11);
                            table.HD = reader.GetInt32(12);
                            table.HL = reader.GetInt32(13);
                            table.HF = reader.GetInt32(14);
                            table.HA = reader.GetInt32(15);
                            try { table.HAT = reader.GetInt32(16); } catch (Exception e) { table.HAT = 0; }
                            table.AP = reader.GetInt32(17);
                            table.AW = reader.GetInt32(18);
                            table.AD = reader.GetInt32(19);
                            table.AL = reader.GetInt32(20);
                            table.AF = reader.GetInt32(21);
                            try { table.AA = reader.GetInt32(22); } catch (Exception e) { table.AA = 0; }
                            try { table.AAT = reader.GetInt32(23); } catch (Exception e) { table.AAT = 0; }
                            MatchesTable.Add(table);
                        }
                    }
                }
            }
        }
        public class Matches
        {
            public string comptype1;
            public string comptype2;
            public string competition;
            public int P;
            public int W;
            public int D;
            public int L;
            public int F;
            public int A;
            public int AT;
            public int HP;
            public int HW;
            public int HD;
            public int HL;
            public int HF;
            public int HA;
            public int HAT;
            public int AP;
            public int AW;
            public int AD;
            public int AL;
            public int AF;
            public int AA;
            public int AAT;
        }
    }
}
    
        