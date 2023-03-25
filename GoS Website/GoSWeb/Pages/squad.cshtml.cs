using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;

namespace GoSWeb.Pages
{
    public class squadModel : PageModel
    {
        public List<Player> PlayerTable = new List<Player>();
        
        public void OnGet()
        {
            string connetionString = @"Data Source=sql11.hostinguk.net;Initial Catalog=greenson_greensonscreen;User ID=greenson_gosadmin;Password=d7^fIY5[K6km1G";
            using (SqlConnection cnn = new SqlConnection(connetionString))
            {
                cnn.Open();
                string sql = "with cte as ( select squad_no, b.dob, b.pob, b.signed_this_spell, b.came_from, b.loan, b.loaned_to, b.position, b.surname, b.forename, rtrim(b.forename) + ' ' + rtrim(b.surname)  as name, b.player_id_spell1, c.photo_exists, c.prime_photo, 0 as starts,0 as subs,0 as goals from player_squad a join player b on a.player_id = b.player_id join player c on b.player_id_spell1 = c.player_id where b.last_game_year = 9999 and season_no = (select max(season_no) from player_squad) union all select squad_no, b.dob, b.pob, b.signed_this_spell, b.came_from, b.loan, b.loaned_to, b.position, b.surname, b.forename, rtrim(b.forename) + ' ' + rtrim(b.surname)  as name, b.player_id_spell1, c.photo_exists, c.prime_photo, 1, 0, 0 from player_squad a join player b on a.player_id = b.player_id join player c on b.player_id_spell1 = c.player_id join match_player d on d.player_id in (select player_id from player e where e.player_id_spell1 = b.player_id_spell1) where b.last_game_year = 9999 and season_no = (select max(season_no) from player_squad) and startpos > 0 union all select squad_no, b.dob, b.pob, b.signed_this_spell, b.came_from, b.loan, b.loaned_to, b.position, b.surname, b.forename, rtrim(b.forename) + ' ' + rtrim(b.surname)  as name, b.player_id_spell1, c.photo_exists, c.prime_photo, 0, 1 ,0 from player_squad a join player b on a.player_id = b.player_id join player c on b.player_id_spell1 = c.player_id join match_player d on d.player_id in (select player_id from player e where e.player_id_spell1 = b.player_id_spell1) where b.last_game_year = 9999 and season_no = (select max(season_no) from player_squad) and startpos = 0 union all select squad_no, b.dob, b.pob, b.signed_this_spell, b.came_from, b.loan, b.loaned_to, b.position, b.surname, b.forename, rtrim(b.forename) + ' ' + rtrim(b.surname)  as name, b.player_id_spell1, c.photo_exists, c.prime_photo, 0, 0, 1 from player_squad a join player b on a.player_id = b.player_id join player c on b.player_id_spell1 = c.player_id join match_goal d on d.player_id in (select player_id from player e where e.player_id_spell1 = b.player_id_spell1) where b.last_game_year = 9999 and season_no = (select max(season_no) from player_squad) )select squad_no, dob, pob, signed_this_spell, came_from, loan, loaned_to, case left(position,3) when 'Goa' then '1' + position when 'Def' then '2' + position when 'Mid' then '3' + position when 'For' then '4' + position end as sortposition, surname, forename, name, player_id_spell1, photo_exists, prime_photo, sum(starts) + sum(subs) as appears, sum(starts) as starts, sum(subs) as subs, sum(goals) as goals from CTE group by squad_no, dob, pob, signed_this_spell, came_from, loan, loaned_to, position, surname, forename, name, player_id_spell1, photo_exists, prime_photo ";
                using (SqlCommand command = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Player table = new Player();
                            table.squad_no = reader.GetByte(0);
                            table.dob = reader.GetDateTime(1);
                            
                            try
                            {
                                table.pob = reader.GetString(2);
                            }
                            catch
                            {
                                table.pob = "N";
                            }
                            
                            try
                            {
                                table.signed_this_spell = reader.GetDateTime(3);
                            }
                            catch
                            {
                                table.signed_this_spell = new DateTime();
                            }
                            table.came_from = reader.GetString(4);
                            try
                            {
                               table.loan = reader.GetChar(5);
                            }
                            catch
                            {
                                table.loan = 'N';
                            }
                            try
                            {
                                table.loaned_to = reader.GetChar(6);
                            }
                            catch
                            {
                                table.loaned_to = 'N';
                            }                           
                            table.sortposition = reader.GetString(7);
                            table.cooper =  reader.GetString(8);
                            table.forename = reader.GetString(9);
                            table.name = reader.GetString(10);
                            table.player_id_spell1 = Convert.ToInt32(reader.GetInt16(11));
                            try
                            {
                                table.photo_exits = reader.GetChar(12);
                            }
                            catch
                            {
                                table.photo_exits = 'N';
                            }
                            try
                            {
                               table.prime_photo = reader.GetByte(13); 
                            }
                            catch
                            {
                                table.prime_photo = 0;
                            }
                            
                            table.appers = reader.GetInt32(14);
                            table.starts = reader.GetInt32(15);
                            table.subs = reader.GetInt32(16);
                            table.goal = reader.GetInt32(17);
                            PlayerTable.Add(table);
                        }
                    }
                }
            }
        }



                

        
        public class Player
        {
            public byte squad_no;
            public DateTime dob;
            public string pob;
            public DateTime signed_this_spell;
            public string came_from;
            public char loan;
            public char loaned_to;
            public string sortposition;
            public string cooper;
            public string forename;
            public string name;
            public int player_id_spell1;
            public char photo_exits;
            public byte prime_photo;
            public int appers;
            public int starts;
            public int subs;
            public int goal;
        }
    }
}
