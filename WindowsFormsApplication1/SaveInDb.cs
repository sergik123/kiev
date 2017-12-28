using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class SaveInDb:Form1
    {
        public SaveInDb(Dictionary<string,string> args,bool flag_edit)
        {
            var com= new SQLiteCommand();
            string path = Environment.CurrentDirectory + "/MyDataBase/info.bytes.db";
            SQLiteConnection con_db = new SQLiteConnection(string.Format("DATA Source={0};", path));
            SQLiteParameter param = new SQLiteParameter();
            con_db.Open();
            if (flag_edit)
            {
                com = new SQLiteCommand("UPDATE 'kartka_obliku' SET code_region=@param1,name_region=@param2,code_raion=@param3,name_raion=@param4,code_region_item=@param5,number_cartka=@param6,main_dop=@param7,date_viniknenya=@param8,code_adress=@param9," +
                   "name_adress=@param10,fire_code=@param11,fire_item=@param12,code_forma=@param13,name_forma=@param14,code_riziku=@param15,name_riziku=@param16,code_object=@param17,name_object=@param18,poverhovist=@param19,code_poverh=@param20,name_poverh=@param21,code_stoikist=@param22,name_stoikist=@param23,code_category=@param24,name_category=@param25,code_place=@param26,name_place=@param27,code_virib=@param28,item_virib=@param29,code_pricini=@param30,name_pricini=@param31," +
                   "viavleno=@param32,via_ditei=@param33,zag_vnaslidok=@param34,zag_ditei=@param35,zag_fire=@param36,zag_names=@param37,zag_vik=@param38,zag_stat_code=@param39,zag_stat_name=@param40,code_status=@param41,name_status=@param42,code_moment=@param43,moment=@param44,code_umovi=@param45,name_umovi=@param46," +
                   "travm=@param47,travm_ditei=@param48,travm_fire=@param49,pramiy=@param50,pobichniy=@param51,zn_bud=@param52,posh_bud=@param53,zn_tehnika=@param54,posh_tehnika=@param55,zn_zerno=@param56,zn_koreni=@param57,zn_valki=@param58,zn_korm=@param59,zn_torf=@param60,posh_torf=@param61,zag_tvarin=@param62,zag_ptici=@param63,dop_info=@param64," +
                   "vr_ludei=@param65,vr_ditei=@param66,vr_tvarin=@param67,vr_ptici=@param68,vr_bud=@param69,vr_tehnika=@param70,vr_zerno=@param71,vr_koreni=@param72,vr_valki=@param73,vr_korm=@param74,vr_torf=@param75,vr_dop=@param76,vr_mat=@param77," +
                   "data_pov=@param78,time_pov=@param79,time_pributa=@param80,code_fire_likvid=@param81,name_fire_likvid=@param82,data_lokal=@param83,time_lokal=@param84,data_likvid=@param85,time_likvid=@param86,code_umovi_posh=@param87,name_umovi_posh=@param88,code_umovi_uskl=@param89,name_umovi_uskl=@param90,code_spz=@param91,name_spz=@param92,code_system=@param93,name_system=@param94,code_dii=@param95,name_dii=@param96," +
                   "code_uch=@param97,name_uch=@param98,kilk_uch=@param99,code_fireauto=@param100,name_fireauto=@param101,kilk_auto=@param102,code_firestvol=@param103,name_firestvol=@param104,kilk_firestvol=@param105,code_rechovini=@param106,name_rechovini=@param107,code_pervini=@param108,name_pervini=@param109,code_djerela=@param110,name_djerela=@param111,vikr_gds=@param112,kilk_gds=@param113,time_gds=@param114," +
                   "data_perevirki=@param115,code_perevirka=@param116,name_perevirka=@param117,code_dialnist=@param118,name_dialnist=@param119,code_zahodi=@param120,name_zahodi=@param121,number_kku=@param121_1,data_zapovnenya=@param122,pid_osibi=@param123 WHERE number_cartka=@param6 AND main_dop=@param7", con_db);
            }
            else
            {
                com = new SQLiteCommand("INSERT INTO 'kartka_obliku' ('code_region','name_region','code_raion','name_raion','code_region_item','number_cartka','main_dop','date_viniknenya','code_adress'," +
                    "'name_adress','fire_code','fire_item','code_forma','name_forma','code_riziku','name_riziku','code_object','name_object','poverhovist','code_poverh','name_poverh','code_stoikist','name_stoikist','code_category','name_category','code_place','name_place','code_virib','item_virib','code_pricini','name_pricini'," +
                    "'viavleno','via_ditei','zag_vnaslidok','zag_ditei','zag_fire','zag_names','zag_vik','zag_stat_code','zag_stat_name','code_status','name_status','code_moment','moment','code_umovi','name_umovi'," +
                    "'travm','travm_ditei','travm_fire','pramiy','pobichniy','zn_bud','posh_bud','zn_tehnika','posh_tehnika','zn_zerno','zn_koreni','zn_valki','zn_korm','zn_torf','posh_torf','zag_tvarin','zag_ptici','dop_info'," +
                    "'vr_ludei','vr_ditei','vr_tvarin','vr_ptici','vr_bud','vr_tehnika','vr_zerno','vr_koreni','vr_valki','vr_korm','vr_torf','vr_dop','vr_mat'," +
                    "'data_pov','time_pov','time_pributa','code_fire_likvid','name_fire_likvid','data_lokal','time_lokal','data_likvid','time_likvid','code_umovi_posh','name_umovi_posh','code_umovi_uskl','name_umovi_uskl','code_spz','name_spz','code_system','name_system','code_dii','name_dii'," +
                    "'code_uch','name_uch','kilk_uch','code_fireauto','name_fireauto','kilk_auto','code_firestvol','name_firestvol','kilk_firestvol','code_rechovini','name_rechovini','code_pervini','name_pervini','code_djerela','name_djerela','vikr_gds','kilk_gds','time_gds'," +
                    "'data_perevirki','code_perevirka','name_perevirka','code_dialnist','name_dialnist','code_zahodi','name_zahodi','number_kku','data_zapovnenya','pid_osibi') VALUES (@param1,@param2,@param3,@param4,@param5,@param6,@param7,@param8,@param9,@param10,@param11,@param12,@param13,@param14,@param15,@param16,@param17,@param18,@param19,@param20,@param21,@param22,@param23,@param24,@param25,@param26,@param27,@param28,@param29,@param30,@param31,@param32,@param33,@param34,@param35,@param36,@param37,@param38,@param39,@param40,@param41,@param42,@param43,@param44,@param45,@param46,@param47,@param48,@param49,@param50,@param51,@param52,@param53,@param54,@param55,@param56,@param57,@param58,@param59,@param60,@param61,@param62,@param63,@param64,@param65,@param66,@param67,@param68,@param69," +
                    "@param70,@param71,@param72,@param73,@param74,@param75,@param76,@param77,@param78,@param79,@param80,@param81,@param82,@param83,@param84,@param85,@param86,@param87,@param88,@param89,@param90,@param91,@param92,@param93,@param94,@param95,@param96,@param97,@param98,@param99,@param100,@param101,@param102,@param103,@param104,@param105,@param106,@param107,@param108,@param109,@param110,@param111,@param112,@param113,@param114,@param115,@param116,@param117,@param118,@param119,@param120,@param121,@param121_1,@param122,@param123)", con_db);
            }
           
            //try
           // {
               

                //  Console.WriteLine(val.Key+":"+val.Value);
              /*  var com = new SQLiteCommand("INSERT INTO 'kartka_obliku' ('code_region','name_region','code_raion','name_raion','code_region_item','number_cartka','main_dop','date_viniknenya','code_adress'," +
                    "'name_adress','fire_code','fire_item','code_forma','name_forma','code_riziku','name_riziku','code_object','name_object','poverhovist','code_poverh','name_poverh','code_stoikist','name_stoikist','code_category','name_category','code_place','name_place','code_virib','item_virib','code_pricini','name_pricini'," +
                    "'viavleno','via_ditei','zag_vnaslidok','zag_ditei','zag_fire','zag_names','zag_vik','zag_stat_code','zag_stat_name','code_status','name_status','code_moment','moment','code_umovi','name_umovi'," +
                    "'travm','travm_ditei','travm_fire','pramiy','pobichniy','zn_bud','posh_bud','zn_tehnika','posh_tehnika','zn_zerno','zn_koreni','zn_valki','zn_korm','zn_torf','posh_torf','zag_tvarin','zag_ptici','dop_info'," +
                    "'vr_ludei','vr_ditei','vr_tvarin','vr_ptici','vr_bud','vr_tehnika','vr_zerno','vr_koreni','vr_valki','vr_korm','vr_torf','vr_dop','vr_mat'," +
                    "'data_pov','time_pov','time_pributa','code_fire_likvid','name_fire_likvid','data_lokal','time_lokal','data_likvid','time_likvid','code_umovi_posh','name_umovi_posh','code_umovi_uskl','name_umovi_uskl','code_spz','name_spz','code_system','name_system','code_dii','name_dii'," +
                    "'code_uch','name_uch','kilk_uch','code_fireauto','name_fireauto','kilk_auto','code_firestvol','name_firestvol','kilk_firestvol','code_rechovini','name_rechovini','code_pervini','name_pervini','code_djerela','name_djerela','vikr_gds','kilk_gds','time_gds'," +
                    "'data_perevirki','code_perevirka','name_perevirka','code_dialnist','name_dialnist','code_zahodi','name_zahodi','number_kku','data_zapovnenya','pid_osibi') VALUES (@param1,@param2,@param3,@param4,@param5,@param6,@param7,@param8,@param9,@param10,@param11,@param12,@param13,@param14,@param15,@param16,@param17,@param18,@param19,@param20,@param21,@param22,@param23,@param24,@param25,@param26,@param27,@param28,@param29,@param30,@param31,@param32,@param33,@param34,@param35,@param36,@param37,@param38,@param39,@param40,@param41,@param42,@param43,@param44,@param45,@param46,@param47,@param48,@param49,@param50,@param51,@param52,@param53,@param54,@param55,@param56,@param57,@param58,@param59,@param60,@param61,@param62,@param63,@param64,@param65,@param66,@param67,@param68,@param69," +
                    "@param70,@param71,@param72,@param73,@param74,@param75,@param76,@param77,@param78,@param79,@param80,@param81,@param82,@param83,@param84,@param85,@param86,@param87,@param88,@param89,@param90,@param91,@param92,@param93,@param94,@param95,@param96,@param97,@param98,@param99,@param100,@param101,@param102,@param103,@param104,@param105,@param106,@param107,@param108,@param109,@param110,@param111,@param112,@param113,@param114,@param115,@param116,@param117,@param118,@param119,@param120,@param121,@param121_1,@param122,@param123)", con_db);*/


                //   com.Parameters.AddWithValue("@param1", val.Key);
                
                com.Parameters.AddWithValue("@param1", args["code_region"]);
                com.Parameters.AddWithValue("@param2", args["name_region"]);
                com.Parameters.AddWithValue("@param3", args["code_raion"]);
                com.Parameters.AddWithValue("@param4", args["name_raion"]);
                com.Parameters.AddWithValue("@param5", args["code_region_item"]);
                com.Parameters.AddWithValue("@param6", args["number_cartka"]);
                com.Parameters.AddWithValue("@param7", args["main_dop"]);
                com.Parameters.AddWithValue("@param8", args["date_viniknenya"]);

              
                com.Parameters.AddWithValue("@param9", args["code_adress"]);
                com.Parameters.AddWithValue("@param10", args["name_adress"]);
                com.Parameters.AddWithValue("@param11", args["fire_code"]);
                com.Parameters.AddWithValue("@param12", args["fire_item"]);
                com.Parameters.AddWithValue("@param13", args["code_forma"]);
                com.Parameters.AddWithValue("@param14", args["name_forma"]);
                com.Parameters.AddWithValue("@param15", args["code_riziku"]);
                com.Parameters.AddWithValue("@param16", args["name_riziku"]);
                com.Parameters.AddWithValue("@param17", args["code_object"]);
                com.Parameters.AddWithValue("@param18", args["name_object"]);
                com.Parameters.AddWithValue("@param19", args["poverhovist"]);
                com.Parameters.AddWithValue("@param20", args["code_poverh"]);
                com.Parameters.AddWithValue("@param21", args["name_poverh"]);
                com.Parameters.AddWithValue("@param22", args["code_stoikist"]);
                com.Parameters.AddWithValue("@param23", args["name_stoikist"]);
                com.Parameters.AddWithValue("@param24", args["code_category"]);
                com.Parameters.AddWithValue("@param25", args["name_category"]);
                com.Parameters.AddWithValue("@param26", args["code_place"]);
                com.Parameters.AddWithValue("@param27", args["name_place"]);
                com.Parameters.AddWithValue("@param28", args["code_virib"]);
                com.Parameters.AddWithValue("@param29", args["item_virib"]);
                com.Parameters.AddWithValue("@param30", args["code_pricini"]);
                com.Parameters.AddWithValue("@param31", args["name_pricini"]);

                com.Parameters.AddWithValue("@param32", args["viavleno"]);
                com.Parameters.AddWithValue("@param33", args["via_ditei"]);
                com.Parameters.AddWithValue("@param34", args["zag_vnaslidok"]);
                com.Parameters.AddWithValue("@param35", args["zag_ditei"]);
                com.Parameters.AddWithValue("@param36", args["zag_fire"]);
                com.Parameters.AddWithValue("@param37", args["zag_names"]);
                com.Parameters.AddWithValue("@param38", args["zag_vik"]);
                com.Parameters.AddWithValue("@param39", args["zag_stat_code"]);
                com.Parameters.AddWithValue("@param40", args["zag_stat_name"]);
                com.Parameters.AddWithValue("@param41", args["code_status"]);
                com.Parameters.AddWithValue("@param42", args["name_status"]);
                com.Parameters.AddWithValue("@param43", args["code_moment"]);
                com.Parameters.AddWithValue("@param44", args["moment"]);
                com.Parameters.AddWithValue("@param45", args["code_umovi"]);
                com.Parameters.AddWithValue("@param46", args["name_umovi"]);
                com.Parameters.AddWithValue("@param47", args["travm"]);
                com.Parameters.AddWithValue("@param48", args["travm_ditei"]);
                com.Parameters.AddWithValue("@param49", args["travm_fire"]);
                com.Parameters.AddWithValue("@param50", args["pramiy"]);
                com.Parameters.AddWithValue("@param51", args["pobichniy"]);
                com.Parameters.AddWithValue("@param52", args["zn_bud"]);
                com.Parameters.AddWithValue("@param53", args["posh_bud"]);
                com.Parameters.AddWithValue("@param54", args["zn_tehnika"]);
                com.Parameters.AddWithValue("@param55", args["posh_tehnika"]);
                com.Parameters.AddWithValue("@param56", args["zn_zerno"]);
                com.Parameters.AddWithValue("@param57", args["zn_koreni"]);
                com.Parameters.AddWithValue("@param58", args["zn_valki"]);
                com.Parameters.AddWithValue("@param59", args["zn_korm"]);
                com.Parameters.AddWithValue("@param60", args["zn_torf"]);
                com.Parameters.AddWithValue("@param61", args["posh_torf"]);
                com.Parameters.AddWithValue("@param62", args["zag_tvarin"]);
                com.Parameters.AddWithValue("@param63", args["zag_ptici"]);
                com.Parameters.AddWithValue("@param64", args["dop_info"]);

                com.Parameters.AddWithValue("@param65", args["vr_ludei"]);
                com.Parameters.AddWithValue("@param66", args["vr_ditei"]);
                com.Parameters.AddWithValue("@param67", args["vr_tvarin"]);
                com.Parameters.AddWithValue("@param68", args["vr_ptici"]);
                com.Parameters.AddWithValue("@param69", args["vr_bud"]);
                com.Parameters.AddWithValue("@param70", args["vr_tehnika"]);
                com.Parameters.AddWithValue("@param71", args["vr_zerno"]);
                com.Parameters.AddWithValue("@param72", args["vr_koreni"]);
                com.Parameters.AddWithValue("@param73", args["vr_valki"]);
                com.Parameters.AddWithValue("@param74", args["vr_korm"]);
                com.Parameters.AddWithValue("@param75", args["vr_torf"]);
                com.Parameters.AddWithValue("@param76", args["vr_dop"]);
                com.Parameters.AddWithValue("@param77", args["vr_mat"]);

                com.Parameters.AddWithValue("@param78", args["data_pov"]);
                com.Parameters.AddWithValue("@param79", args["time_pov"]);
                com.Parameters.AddWithValue("@param80", args["time_pributa"]);
                com.Parameters.AddWithValue("@param81", args["code_fire_likvid"]);
                com.Parameters.AddWithValue("@param82", args["name_fire_likvid"]);
                com.Parameters.AddWithValue("@param83", args["data_lokal"]);
                com.Parameters.AddWithValue("@param84", args["time_lokal"]);
                com.Parameters.AddWithValue("@param85", args["data_likvid"]);
                com.Parameters.AddWithValue("@param86", args["time_likvid"]);
                com.Parameters.AddWithValue("@param87", args["code_umovi_posh"]);
                com.Parameters.AddWithValue("@param88", args["name_umovi_posh"]);
                com.Parameters.AddWithValue("@param89", args["code_umovi_uskl"]);
                com.Parameters.AddWithValue("@param90", args["name_umovi_uskl"]);
                com.Parameters.AddWithValue("@param91", args["code_spz"]);
                com.Parameters.AddWithValue("@param92", args["name_spz"]);
                com.Parameters.AddWithValue("@param93", args["code_system"]);
                com.Parameters.AddWithValue("@param94", args["name_system"]);
                com.Parameters.AddWithValue("@param95", args["code_dii"]);
                com.Parameters.AddWithValue("@param96", args["name_dii"]);

                com.Parameters.AddWithValue("@param97", args["code_uch"]);
                com.Parameters.AddWithValue("@param98", args["name_uch"]);
                com.Parameters.AddWithValue("@param99", args["kilk_uch"]);
                com.Parameters.AddWithValue("@param100", args["code_fireauto"]);
                com.Parameters.AddWithValue("@param101", args["name_fireauto"]);
                com.Parameters.AddWithValue("@param102", args["kilk_auto"]);
                com.Parameters.AddWithValue("@param103", args["code_firestvol"]);
                com.Parameters.AddWithValue("@param104", args["name_firestvol"]);
                com.Parameters.AddWithValue("@param105", args["kilk_firestvol"]);
                com.Parameters.AddWithValue("@param106", args["code_rechovini"]);
                com.Parameters.AddWithValue("@param107", args["name_rechovini"]);
                com.Parameters.AddWithValue("@param108", args["code_pervini"]);
                com.Parameters.AddWithValue("@param109", args["name_pervini"]);
                com.Parameters.AddWithValue("@param110", args["code_djerela"]);
                com.Parameters.AddWithValue("@param111", args["name_djerela"]);
                com.Parameters.AddWithValue("@param112", args["vikr_gds"]);
                com.Parameters.AddWithValue("@param113", args["kilk_gds"]);
                com.Parameters.AddWithValue("@param114", args["time_gds"]);

            com.Parameters.AddWithValue("@param115", args["data_perevirki"]);
            com.Parameters.AddWithValue("@param116", args["code_perevirka"]);
            com.Parameters.AddWithValue("@param117", args["name_perevirka"]);
            com.Parameters.AddWithValue("@param118", args["code_dialnist"]);
            com.Parameters.AddWithValue("@param119", args["name_dialnist"]);
            com.Parameters.AddWithValue("@param120", args["code_zahodi"]);
            com.Parameters.AddWithValue("@param121", args["name_zahodi"]);
            com.Parameters.AddWithValue("@param121_1", args["number_kku"]);
            com.Parameters.AddWithValue("@param122", args["data_zapovnenya"]);
            com.Parameters.AddWithValue("@param123", args["pid_osibi"]);
            com.ExecuteNonQuery();
        //    }
        //    catch (Exception ex)
         //   {

         //       Console.WriteLine(ex.Message);
          //  }
            

            
        }

    }
}
