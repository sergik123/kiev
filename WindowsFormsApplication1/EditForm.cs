using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SQLite;
using System.Data.Common;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class EditForm : Form
    {
     
        public void GetFieldfromDb(string zapros, Form1 form)
        {
            string path = Environment.CurrentDirectory + "/MyDataBase/info.bytes.db";
            SQLiteConnection con_db = new SQLiteConnection(string.Format("DATA Source={0};", path));
            SQLiteParameter param = new SQLiteParameter();
            con_db.Open();

            SQLiteCommand cmd_db = new SQLiteCommand(zapros, con_db);
           SQLiteDataReader rdr = cmd_db.ExecuteReader();
            string[] pib,vik,stat,status,moment,umova, posh, uskl,system, dii, code_uch, kilk_uch,fireauto, kilk_auto,code_stvol,kilk_stvol,code_rech,
                code_pervini, code_djerelo, code_zahodi;
            while (rdr.Read())
            {
                form.textBox1.Text = rdr[3].ToString();
                form.textBox70.Text = rdr[5].ToString();
                form.textBox2.Text = rdr[6].ToString();
                form.textBox3.Text = rdr[7].ToString();
                form.dateTimePicker1.Value =DateTime.Parse(rdr[8].ToString());

                /*2 раздел*/
                form.textBox71.Text = rdr[9].ToString();
                form.textBox4.Text = rdr[10].ToString();
                form.textBox72.Text = rdr[11].ToString();
                form.textBox73.Text = rdr[12].ToString();
                form.textBox74.Text = rdr[13].ToString();
                form.textBox75.Text = rdr[15].ToString();
                form.textBox76.Text = rdr[17].ToString();
                form.textBox67.Text = rdr[19].ToString();
                form.textBox68.Text = rdr[20].ToString();
                form.textBox77.Text = rdr[22].ToString();
                form.textBox78.Text = rdr[24].ToString();
                form.textBox79.Text = rdr[26].ToString();
                form.textBox156.Text = rdr[27].ToString();
                form.textBox80.Text = rdr[28].ToString();
                form.textBox157.Text = rdr[29].ToString();
                form.textBox81.Text = rdr[30].ToString();
                form.textBox158.Text = rdr[31].ToString();

                /*3 раздел*/
                form.textBox5.Text = rdr[32].ToString();
                form.textBox6.Text = rdr[33].ToString();
                form.textBox8.Text = rdr[34].ToString();
                form.textBox65.Text = rdr[35].ToString();
                form.textBox82.Text = rdr[36].ToString();
                
                pib = rdr[37].ToString().Split(',');
                if(pib[0]!=" ")
                form.textBox18.Text = pib[0];
                if (pib[1]!= " ")
                form.textBox20.Text = pib[1];
                if (pib[2]!= " ")
                form.textBox22.Text = pib[2];
                if (pib[3]!= " ")
                form.textBox24.Text = pib[3];
                if (pib[4]!= " ")
                form.textBox26.Text = pib[4];

                vik = rdr[38].ToString().Split(',');
                form.textBox17.Text = vik[0];
                form.textBox19.Text = vik[1];
                form.textBox21.Text = vik[2];
                form.textBox23.Text = vik[3];
                form.textBox25.Text = vik[4];

                stat = rdr[39].ToString().Split(',');
                form.textBox10.Text = stat[0];
                form.textBox83.Text = stat[1];
                form.textBox84.Text = stat[2];
                form.textBox85.Text = stat[3];
                form.textBox86.Text = stat[4];

                status = rdr[41].ToString().Split(',');
                form.textBox87.Text = status[0];
                form.textBox88.Text = status[1];
                form.textBox89.Text = status[2];
                form.textBox90.Text = status[3];
                form.textBox91.Text = status[4];

                moment = rdr[43].ToString().Split(',');
                form.textBox92.Text = moment[0];
                form.textBox93.Text = moment[1];
                form.textBox94.Text = moment[2];
                form.textBox95.Text = moment[3];
                form.textBox96.Text = moment[4];

                umova = rdr[45].ToString().Split(',');
                form.textBox97.Text = umova[0];
                form.textBox98.Text = umova[1];
                form.textBox99.Text = umova[2];
                form.textBox100.Text = umova[3];
                form.textBox101.Text = umova[4];

                form.textBox12.Text = rdr[47].ToString();
                form.textBox11.Text = rdr[48].ToString();
                form.textBox66.Text = rdr[49].ToString();
                form.textBox9.Text  = rdr[50].ToString();
                form.textBox15.Text = rdr[51].ToString();
                form.textBox7.Text  = rdr[52].ToString();
                form.textBox14.Text = rdr[53].ToString();
                form.textBox13.Text = rdr[54].ToString();
                form.textBox16.Text = rdr[55].ToString();
                form.textBox28.Text = rdr[56].ToString();
                form.textBox27.Text = rdr[57].ToString();
                form.textBox36.Text = rdr[58].ToString();
                form.textBox30.Text = rdr[59].ToString();
                form.textBox29.Text = rdr[60].ToString();
                form.textBox35.Text = rdr[61].ToString();
                form.textBox32.Text = rdr[62].ToString();
                form.textBox31.Text = rdr[63].ToString();
                form.textBox33.Text = rdr[64].ToString();

                /*4 раздел*/
                form.textBox37.Text = rdr[65].ToString();
                form.textBox34.Text = rdr[66].ToString();
                form.textBox39.Text = rdr[67].ToString();
                form.textBox38.Text = rdr[68].ToString();
                form.textBox41.Text = rdr[69].ToString();
                form.textBox40.Text = rdr[70].ToString();
                form.textBox44.Text = rdr[71].ToString();
                form.textBox46.Text = rdr[72].ToString();
                form.textBox45.Text = rdr[73].ToString();
                form.textBox43.Text = rdr[74].ToString();
                form.textBox42.Text = rdr[75].ToString();
                form.textBox47.Text = rdr[76].ToString();
                form.textBox48.Text = rdr[77].ToString();

                /*5 раздел*/
                bool res = DateTime.TryParse(rdr[78].ToString(), out DateTime date1);
                bool res1 = DateTime.TryParse(rdr[79].ToString(), out DateTime date2);
                bool res2 = DateTime.TryParse(rdr[80].ToString(), out DateTime date3);
                bool res3 = DateTime.TryParse(rdr[83].ToString(), out DateTime date4);
                bool res4 = DateTime.TryParse(rdr[84].ToString(), out DateTime date5);
                bool res5 = DateTime.TryParse(rdr[85].ToString(), out DateTime date6);
                bool res6 = DateTime.TryParse(rdr[86].ToString(), out DateTime date7);

                if(res)
                form.dateTimePicker2.Value = date1;
                if (res1)
                form.dateTimePicker3.Value = date2;
                if (res2)
                form.dateTimePicker4.Value = date3;
                form.textBox102.Text = rdr[81].ToString();
                if (res3)
                form.dateTimePicker6.Value = date4;
                if (res4)
                form.dateTimePicker5.Value = date5;
                if (res5)
                form.dateTimePicker8.Value = date6;
                if (res6)
                form.dateTimePicker7.Value = date7;

                posh = rdr[87].ToString().Split(',');
                form.textBox103.Text  = posh[0];
                form.textBox104.Text  = posh[1];
                form.textBox105.Text  = posh[2];
                form.textBox106.Text = posh[3];
                form.textBox107.Text = posh[4];

                uskl = rdr[89].ToString().Split(',');
                form.textBox112.Text  = uskl[0];
                form.textBox111.Text  = uskl[1];
                form.textBox110.Text  = uskl[2];
                form.textBox109.Text  = uskl[3];
                form.textBox108.Text  = uskl[4];

                form.textBox113.Text = rdr[91].ToString();

                system = rdr[93].ToString().Split(',');
                form.textBox118.Text = system[0];
                form.textBox117.Text = system[1];
                form.textBox116.Text = system[2];
                form.textBox115.Text = system[3];
                form.textBox114.Text = system[4];

                dii = rdr[95].ToString().Split(',');
                form.textBox123.Text = dii[0];
                form.textBox122.Text = dii[1];
                form.textBox121.Text = dii[2];
                form.textBox120.Text = dii[3];
                form.textBox119.Text = dii[4];

                /*6 раздел*/
                code_uch = rdr[97].ToString().Split(',');
                form.textBox128.Text = code_uch[0];
                form.textBox127.Text = code_uch[1];
                form.textBox126.Text = code_uch[2];
                form.textBox125.Text = code_uch[3];
                form.textBox124.Text = code_uch[4];

                kilk_uch = rdr[99].ToString().Split(',');
                form.textBox49.Text = kilk_uch[0];
                form.textBox50.Text = kilk_uch[1];
                form.textBox51.Text = kilk_uch[2];
                form.textBox52.Text = kilk_uch[3];
                form.textBox53.Text = kilk_uch[4];

                fireauto = rdr[100].ToString().Split(',');
                form.textBox129.Text = fireauto[0];
                form.textBox130.Text = fireauto[1];
                form.textBox131.Text = fireauto[2];
                form.textBox132.Text = fireauto[3];
                form.textBox133.Text = fireauto[4];

                kilk_auto = rdr[102].ToString().Split(',');
                form.textBox58.Text = kilk_auto[0];
                form.textBox57.Text = kilk_auto[1];
                form.textBox56.Text = kilk_auto[2];
                form.textBox55.Text = kilk_auto[3];
                form.textBox54.Text = kilk_auto[4];

                code_stvol = rdr[103].ToString().Split(',');
                form.textBox134.Text = code_stvol[0];
                form.textBox135.Text = code_stvol[1];
                form.textBox136.Text = code_stvol[2];

                kilk_stvol = rdr[105].ToString().Split(',');
                form.textBox61.Text = kilk_stvol[0];
                form.textBox60.Text = kilk_stvol[1];
                form.textBox59.Text = kilk_stvol[2];

                code_rech = rdr[106].ToString().Split(',');
                form.textBox137.Text = code_rech[0];
                form.textBox138.Text = code_rech[1];
                form.textBox139.Text = code_rech[2];

                code_pervini = rdr[108].ToString().Split(',');
                form.textBox140.Text = code_pervini[0];
                form.textBox141.Text = code_pervini[1];
                form.textBox142.Text = code_pervini[2];

                code_djerelo = rdr[110].ToString().Split(',');
                form.textBox143.Text = code_djerelo[0];
                form.textBox144.Text = code_djerelo[1];
                form.textBox145.Text = code_djerelo[2];

                form.textBox62.Text = rdr[112].ToString();
                form.textBox63.Text = rdr[113].ToString();
                form.textBox64.Text = rdr[114].ToString();

                /*7 раздел*/
                bool res7 = DateTime.TryParse(rdr[115].ToString(), out DateTime date8);
                if(res7)
                form.dateTimePicker9.Value = date8;
                form.textBox146.Text = rdr[116].ToString();
                form.textBox147.Text = rdr[118].ToString();

                code_zahodi = rdr[120].ToString().Split(',');
                form.textBox148.Text = code_zahodi[0];
                form.textBox149.Text = code_zahodi[1];
                form.textBox159.Text = rdr[122].ToString();
                form.dateTimePicker10.Value = DateTime.Parse(rdr[123].ToString());
                form.textBox150.Text = rdr[124].ToString();
                // region_items.Add(rdr[1].ToString());
               
            }

        }
    }

}

 