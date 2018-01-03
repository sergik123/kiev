using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word=Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SQLite;
using System.Data.Common;
using System.IO;
using System.Threading;

namespace WindowsFormsApplication1
{
    
    public partial class Form1 : Form
    {
        SQLiteConnection con_db;
	    SQLiteCommand cmd_db;
	    SQLiteDataReader rdr;
List<string> region_items = new List<string>();
        Dictionary<string, string> dict = new Dictionary<string, string>();
        bool state_list = false;
        bool state_list2 = false;
        string id_raion = "";
        string id_np = "";
        int current;
        List<TextBox> tb = new List<TextBox>();
        List<DateTimePicker> tb_date = new List<DateTimePicker>();
        List<MaskedTextBox> tb_masked = new List<MaskedTextBox>();
        List<ComboBox> tb_combo = new List<ComboBox>();
        bool flag_edit = false;
        string error_log="";
        public Form1()
        {
            try
            {
               // Thread.Sleep(100);
                InitializeComponent();
               
            }
            catch (Exception ex)
            {
                error_log = Environment.CurrentDirectory + "/MyDataBase/errorlog.txt";
                String errLog = "";
                if (File.Exists(error_log)) errLog = File.ReadAllText(error_log);
                errLog += "\r\n\r\n" + ex.ToString();
                File.WriteAllText(error_log, errLog);
                MessageBox.Show("Произошла критическая ошибка. Подробнее смотрите в логе ошибок: " + error_log);
                Application.ExitThread();
            }
            


            /* dict["code_region"]="";
             dict["name_region"]="";
             dict["code_raion"]="";
             dict["name_raion"]="";
             dict["code_region_item"]="";
             dict["number_cartka"]="";
             dict["main_dop"]="";
             dict["date_viniknenya"]="";

             dict["code_adress"]="";
             dict["name_adress"]="";
             dict["fire_code"]="";
             dict["fire_item"]="";
             dict["code_forma"]="";
             dict["name_forma"]="";
             dict["code_riziku"]="";
             dict["name_riziku"]="";
             dict["code_object"]="";
             dict["name_object"]="";
             dict["poverhovist"]="";
             dict["code_poverh"]="";
             dict["name_poverh"]="";
             dict["code_stoikist"]="";
             dict["name_stoikist"]="";
             dict["code_category"]="";
             dict["name_category"]="";
             dict["code_place"]="";
             dict["name_place"]="";
             dict["code_virib"]="";
             dict["item_virib"]="";
             dict["code_pricini"]="";
             dict["name_pricini"]="";

             dict["viavleno"]="";
             dict["via_ditei"]="";
             dict["zag_vnaslidok"]="";
             dict["zag_ditei"]="";
             dict["zag_fire"]="";
             dict["zag_names"]="";
             dict["zag_vik"]="";
             dict["zag_stat_code"]="";
             dict["zag_stat_name"]="";
             dict["code_status"]="";
             dict["name_status"]="";
             dict["code_moment"]="";
             dict["moment"]="";
             dict["code_umovi"]="";
             dict["name_umovi"]="";*/
           
        }

      
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text += " V.1.0";
            foreach (var ctl in groupBox1.Controls)
            {
                if (ctl is TextBox) tb.Add((TextBox)ctl);
                if (ctl is DateTimePicker) tb_date.Add((DateTimePicker)ctl);
                if (ctl is MaskedTextBox) tb_masked.Add((MaskedTextBox)ctl);
                if (ctl is ComboBox) tb_combo.Add((ComboBox)ctl);

            }
            foreach (var ctl in groupBox2.Controls)
            {
                if (ctl is TextBox) tb.Add((TextBox)ctl);
                if (ctl is ComboBox) tb_combo.Add((ComboBox)ctl);
            }
            foreach (var ctl in groupBox3.Controls)
            {
                if (ctl is TextBox) tb.Add((TextBox)ctl);
                if (ctl is DateTimePicker) tb_date.Add((DateTimePicker)ctl);
                if (ctl is MaskedTextBox) tb_masked.Add((MaskedTextBox)ctl);
                if (ctl is ComboBox) tb_combo.Add((ComboBox)ctl);
            }
            foreach (var ctl in groupBox4.Controls)
            {
                if (ctl is TextBox) tb.Add((TextBox)ctl);
                if (ctl is ComboBox) tb_combo.Add((ComboBox)ctl);
            }
            foreach (var ctl in groupBox5.Controls)
            {
                if (ctl is TextBox) tb.Add((TextBox)ctl);
                if (ctl is ComboBox) tb_combo.Add((ComboBox)ctl);
            }
            foreach (var ctl in groupBox6.Controls)
            {
                if (ctl is TextBox) tb.Add((TextBox)ctl);
                if (ctl is ComboBox) tb_combo.Add((ComboBox)ctl);
            }
            foreach (var ctl in groupBox7.Controls)
            {
                if (ctl is TextBox) tb.Add((TextBox)ctl);
                if (ctl is DateTimePicker) tb_date.Add((DateTimePicker)ctl);
                if (ctl is MaskedTextBox) tb_masked.Add((MaskedTextBox)ctl);
                if (ctl is ComboBox) tb_combo.Add((ComboBox)ctl);
            }
            foreach (var ctl in panel1.Controls)
            {
                if (ctl is TextBox) tb.Add((TextBox)ctl);
                if (ctl is ComboBox) tb_combo.Add((ComboBox)ctl);

            }


            connection();
            ReadDb();
            textBox69.Text = CodeAllContent("current_region", "id_current", "code_current_region", "1");
            textBox151.Text = CodeAllContent("current_region", "id_current", "code_current_region", "1");
            if (textBox69.Text == "" || comboBox83.Items.Count==0 || comboBox85.Items.Count == 0)
            {
                panel2.Visible = true;
               // ReadDb();
            }
            else
            {
                panel2.Visible = false;
                настройкаToolStripMenuItem.Enabled = true;
               // ReadDb();
            }
            dateTimePicker10.Value = DateTime.Today;
            maskedTextBox6.Text = DateTime.Today.ToString();
        }

 private void button1_Click(object sender, EventArgs e)
 {
            panel6.Visible = true;
            progressBar1.Value += 1;
           string error = "Заповніть наступні поля: ";
            //  try
            //  {
            //  Stream myStream;
         //   SaveFileDialog saveFileDialog1 = new SaveFileDialog();

          //  saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
         //   saveFileDialog1.FilterIndex = 2;
         //   saveFileDialog1.RestoreDirectory = true;

           
           /* ComboBox[] combo = { comboBox1, comboBox2, comboBox3, comboBox5, comboBox6, comboBox7, comboBox9, comboBox10, comboBox11, comboBox12, comboBox24,
                comboBox35, comboBox36, comboBox45, comboBox46,comboBox57,comboBox66,comboBox67,comboBox72,comboBox78,comboBox79,comboBox80,comboBox81};*/
            TextBox[] textbox = { textBox69, textBox1, textBox70, textBox2, textBox71, textBox4, textBox72, textBox79, textBox80, textBox81, textBox9,textBox102,textBox150};

            if (dateTimePicker1.Value.ToString() == "01.01.1900 0:00:00")
            {
                MessageBox.Show("Заповніть дату виникнення пожежі");
                return;
            }
            if (dateTimePicker2.Value.ToString() == "01.01.1900 0:00:00")
            {
                MessageBox.Show("Заповніть дату повідомлення про пожежу");
                return;
            }
            if (dateTimePicker8.Value.ToString() == "01.01.1900 0:00:00")
            {
                if (textBox3.Text == "0")
                {
                    MessageBox.Show("Заповніть дату ліквідації пожежі");
                    return;
                }
               
               
            }
           
           
            if (textBox102.Text == "1")
            {
                if (dateTimePicker3.Value.ToShortTimeString() == "0:00")
                {
                    error += "час повідомлення про пожежу,";
                }
                if (dateTimePicker4.Value.ToString() == "0:00")
                {
                    error += "час прибуття 1-го підрозділу,";
                }
               
            }
            if (textBox102.Text == "2")
            {
                if (dateTimePicker3.Value.ToShortTimeString() == "0:00")
                {
                    error += "час повідомлення про пожежу,";
                }
                if (dateTimePicker4.Value.ToShortTimeString() == "0:00")
                {
                    error += "час прибуття 1-го підрозділу,";
                }
                if (dateTimePicker6.Value.ToString() == "01.01.1900 0:00:00")
                {
                    error += "дата локалізації пожежі,";
                }
                if (dateTimePicker5.Value.ToShortTimeString() == "0:00")
                {
                    error += "час локалізації пожежі,";
                }
                if (dateTimePicker7.Value.ToShortTimeString() == "0:00")
                {
                    error += "час ліквідації пожежі";
                }

            }
            if (error!= "Заповніть наступні поля: ")
            {
                MessageBox.Show(error);
                return;
            }
            if (textBox146.Text != "")
            {
                if (dateTimePicker9.Value.ToString() == "01.01.1900 0:00:00")
                {
                    MessageBox.Show("Заповніть дату останньої перевірки");
                    return;
                }
            }
            if (dateTimePicker1.Value > dateTimePicker2.Value)
            {
                MessageBox.Show("Дата виникнення пожежі не може бути більшою ніж дата повідомлення про пожежу");
                return;
            }
            if (dateTimePicker1.Value < dateTimePicker9.Value)
            {
                MessageBox.Show("Дата виникнення пожежі не може бути меньшою ніж дата останної перевірки");
                return;
            }
            if (dateTimePicker9.Value > dateTimePicker10.Value)
            {
                MessageBox.Show("Дата останної перевірки не може бути більшою ніж дата заповнення картки");
                return;
            }
            if (dateTimePicker6.Value > dateTimePicker8.Value)
            {
                MessageBox.Show("Дата локалізації пожежі не може бути більшою ніж дата ліквідації пожежі");
                return;
            }

            if (validateField(textbox))
            {
                

                object filename = Environment.CurrentDirectory + "/MyDataBase/Proekt_kartky.docx";
                object filename2 = Environment.CurrentDirectory + "/export/Proekt_kartky("+textBox2.Text+","+textBox3.Text+ ").docx";
                var doc = new Word.Application();
                var worddoc = doc.Documents.Open(filename);

            /*1. Загальні дані*/
            var region_text = "";
            var raion_text = "";
            var type_text = "";
            var number_text = "";
            if (textBox69.Text!= "")
            {
                region_text = textBox69.Text;
                CodeFromBase("region", "code_region", "name_region", "{region}", region_text, worddoc);
            }
            else
            {
                ReplaceWord("{n1}", " ", worddoc);
            }
                
            if(textBox1.Text!= "")
            {
                raion_text = textBox1.Text;
                CodeFromBase("current_raion", "code_raion", "name_raion", "{raion}", raion_text, worddoc);
                    //ReplaceWord("{n2}", raion_text, worddoc);
                }
            else
            {
                ReplaceWord("{n2}", " ", worddoc);
            }

            if (textBox70.Text!="")
            {
                type_text = textBox70.Text;
                CodeFromBase("type_region", "code_region_item", "type_region_item", "{type}", type_text, worddoc);
            }
            else
            {
                ReplaceWord("{n3}", " ", worddoc);
            }
                
            if (textBox2.Text!= "")
            {
                number_text = textBox2.Text + "," + textBox3.Text;
                ReplaceWord("{n4}", number_text, worddoc);
                    dict.Add("number_cartka", textBox2.Text);
                    dict.Add("main_dop", textBox3.Text);
                }
            else
            {
                ReplaceWord("{n4}", " ", worddoc);
            }
               

                var day = dateTimePicker1.Value.Day;
                var month = dateTimePicker1.Value.Month;
                var year = dateTimePicker1.Value.Year;
                dict.Add("date_viniknenya", dateTimePicker1.Value.ToShortDateString());
                ReplaceWord("{n1}", region_text, worddoc);
                
                ReplaceWord("{n2}", raion_text, worddoc);
                ReplaceWord("{n3}", type_text, worddoc);
                ReplaceWord("{number}", number_text, worddoc);
                ReplaceWord("{day}", day, worddoc);
                ReplaceWord("{month}", month, worddoc);
                ReplaceWord("{year}", year, worddoc);
                ReplaceWord("{n5}", day, worddoc);
                ReplaceWord("{n6}", month, worddoc);
                ReplaceWord("{n7}", year, worddoc);
                /*2. Інформація про об'єкт пожежі*/
            var code_adr = "";
            var adress = "";
            var object_fire = "";
            var forma_vlasnosti = "";
            var stupin_risiku = "";
            var pidkontrol = "";
            var poverhovist = "";
            var vognestoikist = "";
            var category = "";
            var micse = "";
            var virib = "";
            var pricina = "";
            var poverh = "";
            if (textBox71.Text != "")
                {
                    adress = textBox4.Text;
                    code_adr = textBox71.Text;
                    ReplaceWord("{adress}", adress, worddoc);
                    ReplaceWord("{code_adr}", code_adr, worddoc);
                    dict.Add("code_adress", code_adr);
                    dict.Add("name_adress", adress);
                }
               
            if (textBox72.Text!="")
            {
                object_fire = textBox72.Text;
                    // CodeFromBase("fire_objects", "fire_code", "fire_item", "{object}", object_fire, worddoc);
                    ReplaceWord("{object}", textBox73.Text, worddoc);
                    ReplaceWord("{n9}", textBox72.Text, worddoc);
                    dict.Add("fire_code", textBox72.Text);
                    dict.Add("fire_item", textBox73.Text);
                }
            else
            {
                    ReplaceWord("{n9}", " ", worddoc);
                    ReplaceWord("{object}", " ", worddoc);
                    dict.Add("fire_code", "");
                    dict.Add("fire_item", "");
                }
              
                ReplaceWord("{n9}", object_fire, worddoc);
            if (textBox74.Text!="")
            {
                forma_vlasnosti = textBox74.Text;
                CodeFromBase("forma_vlasnosti", "code_forma", "name_forma", "{forma}", forma_vlasnosti, worddoc);
            }
            else
            {
                ReplaceWord("{n10}", " ", worddoc);
                ReplaceWord("{forma}", " ", worddoc);
                    dict.Add("code_forma", "");
                    dict.Add("name_forma", "");
                }
              
                ReplaceWord("{n10}", forma_vlasnosti, worddoc);
            if (textBox75.Text!="")
            {
                stupin_risiku = textBox75.Text;
                CodeFromBase("stupin_riziku", "code_riziku", "name_riziku", "{stupin}", stupin_risiku, worddoc);
            }
            else
            {
                ReplaceWord("{n11}", " ", worddoc);
                ReplaceWord("{stupin}", " ", worddoc);
                    dict.Add("code_riziku", "");
                    dict.Add("name_riziku", "");
                }
                
                ReplaceWord("{n11}", stupin_risiku, worddoc);
            if (textBox76.Text!="")
            {
                pidkontrol = textBox76.Text;
                CodeFromBase("pidkontrol_object", "code_object", "name_object", "{pidkontrol}", pidkontrol, worddoc);
            }
            else
            {
                ReplaceWord("{n12}", " ", worddoc);
                ReplaceWord("{pidkontrol}", " ", worddoc);
                    dict.Add("code_object", "");
                    dict.Add("name_object", "");
                }
               
                ReplaceWord("{n12}", pidkontrol, worddoc);
            if (textBox68.Text!="")
            {
                    if(textBox68.Text=="201" || textBox68.Text == "202" || textBox68.Text == "203" || textBox68.Text == "204")
                    {
                        CodeFromBase("poverhovist", "code_poverh", "name_poverh", "{poverh}", textBox68.Text, worddoc);
                        // poverhovist = textBox68.Text;

                        // ReplaceWord("{n14}", poverhovist, worddoc);
                    }
                    else
                    {
                        // CodeFromBase("poverhovist", "code_poverh", "name_poverh", "{poverh}", textBox68.Text, worddoc);
                        // poverhovist = "test";
                        poverhovist = textBox68.Text;
                        dict.Add("code_poverh", textBox68.Text);
                        dict.Add("name_poverh", "");
                    }
                    ReplaceWord("{n14}", textBox68.Text, worddoc);

                }
            else
            {
                   
                ReplaceWord("{n14}", " ", worddoc);
                    dict.Add("code_poverh", textBox68.Text);
                    dict.Add("name_poverh", "");
                }
            if (textBox67.Text != "поверх пожежі")
            {
                poverh = textBox67.Text;
                ReplaceWord("{n13}", poverh, worddoc);
            }
            else
            {
                ReplaceWord("{n13}", " ", worddoc);
            }

                
             ReplaceWord("{poverh_build}", poverh, worddoc);
                dict.Add("poverhovist", poverh);
                ReplaceWord("{poverh}", poverhovist, worddoc);
            if (textBox77.Text!="")
            {
                vognestoikist = textBox77.Text;
                CodeFromBase("vognestoikist", "code_stoikist", "name_stoikist", "{vognegasnist}", vognestoikist, worddoc);
            }
            else
            {
                ReplaceWord("{n15}", " ", worddoc);
                ReplaceWord("{vognegasnist}", " ", worddoc);
                    dict.Add("code_stoikist", "");
                    dict.Add("name_stoikist", "");
                }
                
                ReplaceWord("{n15}", vognestoikist, worddoc);
            if (textBox78.Text!="")
            {
                category = textBox78.Text;
                CodeFromBase("category_nebespeki", "code_category", "name_category", "{category}", category, worddoc);
            }
            else
            {
                ReplaceWord("{n16}"," ", worddoc);
                ReplaceWord("{category}", " ", worddoc);
                    dict.Add("code_category", "");
                    dict.Add("name_category", "");
                }
                
                ReplaceWord("{n16}", category, worddoc);
            if (textBox79.Text!="")
            {
                micse = textBox79.Text;
               // CodeFromBase("place_fire", "code_place", "name_place", "{misce}", micse, worddoc);
                    ReplaceWord("{n17}", textBox79.Text, worddoc);
                    ReplaceWord("{misce}", textBox156.Text, worddoc);
                    dict.Add("code_place", textBox79.Text);
                    dict.Add("name_place", textBox156.Text);
            }
            else
            {
                    ReplaceWord("{n17}", " ", worddoc);
                    ReplaceWord("{misce}", " ", worddoc);
                    dict.Add("code_place", "");
                    dict.Add("name_place", "");
            }
                
                ReplaceWord("{n17}", micse, worddoc);
            if (textBox80.Text!="")
            {
                virib = textBox80.Text;
                    // CodeFromBase("virib_iniciator", "code_virib", "item_virib", "{virib}", virib, worddoc);
                    ReplaceWord("{n18}", textBox80.Text, worddoc);
                    ReplaceWord("{virib}", textBox157.Text, worddoc);
                    dict.Add("code_virib", textBox80.Text);
                    dict.Add("item_virib", textBox157.Text);
                }
            else
            {
                ReplaceWord("{n18}", " ", worddoc);

                    ReplaceWord("{virib}", " ", worddoc);
                    dict.Add("code_virib", "");
                    dict.Add("item_virib", "");
                }
                ReplaceWord("{n18}", virib, worddoc);
            if (textBox81.Text!="")
            {
                pricina = textBox81.Text;
                    // CodeFromBase("pricini_fire", "code_pricini", "name_pricini", "{pricina}", pricina, worddoc);
                    ReplaceWord("{n19}", textBox81.Text, worddoc);
                    ReplaceWord("{pricina}", textBox158.Text, worddoc);
                    dict.Add("code_pricini", textBox81.Text);
                    dict.Add("name_pricini", textBox158.Text);
                }
            else
            {
                ReplaceWord("{n19}", " ", worddoc);
                    dict.Add("code_pricini", "");
                    dict.Add("name_pricini", "");
                }
                
                ReplaceWord("{n19}", pricina, worddoc);

            /*3. Наслідки пожежі*/
            var zagiblih = " ";
            var ditei = " ";
            var zag_vnaslidok = " ";
            var zag_ditei = " ";
            var zag_fireman = " ";
            var die1 = " ";
            var die2 = " ";
            var die3 = " ";
            var die4 = " ";
            var die5 = " ";
            if (textBox5.Text!= "")
                zagiblih = textBox5.Text;
                ReplaceWord("{zagiblih}", zagiblih, worddoc);
                dict.Add("viavleno", zagiblih);
                ReplaceWord("{n20}", zagiblih, worddoc);
            if (textBox6.Text!= "")
                ditei = textBox6.Text;
                ReplaceWord("{ditei}", ditei, worddoc);
                dict.Add("via_ditei", ditei);
                ReplaceWord("{n21}", ditei, worddoc);
            if (textBox8.Text!= "")
                zag_vnaslidok = textBox8.Text;
                ReplaceWord("{zag_vnaslidok}", zag_vnaslidok, worddoc);
                dict.Add("zag_vnaslidok", zag_vnaslidok);
                ReplaceWord("{n22}", zag_vnaslidok, worddoc);
            if (textBox65.Text!= "")
                zag_ditei = textBox65.Text;
                ReplaceWord("{zag_ditei}", zag_ditei, worddoc);
                dict.Add("zag_ditei", zag_ditei);
                ReplaceWord("{n23}", zag_ditei, worddoc);
            if (textBox82.Text!= "")
                zag_fireman = textBox82.Text;
                ReplaceWord("{zag_fireman}", zag_fireman, worddoc);
                dict.Add("zag_fire", zag_fireman);
                ReplaceWord("{n24}", zag_fireman, worddoc);

            if (textBox18.Text != "")
                  die1 = textBox18.Text;
                ReplaceWord("{die1}", die1, worddoc);
             
                if (textBox20.Text != "")
                   die2 = textBox20.Text;
                 ReplaceWord("{die2}", die2, worddoc);
             
                if (textBox22.Text != "")
                die3 = textBox22.Text;
                ReplaceWord("{die3}", die3, worddoc);
              
                if (textBox24.Text != "")
                die4 = textBox24.Text;
                ReplaceWord("{die4}", die4, worddoc);
               
                if (textBox26.Text != "")
                 die5 = textBox26.Text;
                ReplaceWord("{die5}", die5, worddoc);

                dict.Add("zag_names", die1+","+die2+","+die3+","+die4+","+die5);
                var vik = " ";
                if (textBox17.Text!= "")
            {
                vik += textBox17.Text;
                ReplaceWord("{n25}", textBox17.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n25}", " ", worddoc);
            }
                    

                if (textBox19.Text != "")
            {
                vik += "," + textBox19.Text;
                ReplaceWord("{n26}", textBox19.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n26}"," ", worddoc);
            }
                    
                if(textBox21.Text!= "")
            {
                vik += "," + textBox21.Text;
                ReplaceWord("{n27}", textBox21.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n27}", " ", worddoc);
            }
                   
                if (textBox23.Text != "")
            {
                vik += "," + textBox23.Text;
                ReplaceWord("{n28}", textBox23.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n28}", " ", worddoc);
            }
                    
                if (textBox25.Text != "")
            {
                vik += "," + textBox25.Text;
                ReplaceWord("{n29}", textBox25.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n29}", " ", worddoc);
            }
                dict.Add("zag_vik", textBox17.Text+","+textBox19.Text+","+textBox21.Text+","+textBox23.Text+","+textBox25.Text);

                ReplaceWord("{vik}", vik, worddoc);

                var stat = " ";
                if (textBox10.Text!="")
            {
                    if (textBox10.Text =="1")
                    {
                        stat += "Чоловіча";
                    }else if(textBox10.Text == "2")
                    {
                        stat += "Жіноча";
                    }else if(textBox10.Text == "3")
                    {
                        stat += "Стать не встановлено";
                    }
               
                ReplaceWord("{n30}", textBox10.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n30}", " ", worddoc);
            }
                 
                if (textBox83.Text!="")
            {
                    if (textBox83.Text == "1")
                    {
                        stat +=","+" Чоловіча";
                    }
                    else if (textBox83.Text == "2")
                    {
                        stat +=","+" Жіноча";
                    }
                    else if (textBox83.Text == "3")
                    {
                        stat +=","+" Стать не встановлено";
                    }
                   // stat += "," + textBox83.Text;
                ReplaceWord("{n31}", textBox83.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n31}", " ", worddoc);
            }
                   
                if (textBox84.Text!="")
            {
                    if (textBox84.Text == "1")
                    {
                        stat +=","+" Чоловіча";
                    }
                    else if (textBox84.Text == "2")
                    {
                        stat +=","+" Жіноча";
                    }
                    else if (textBox84.Text == "3")
                    {
                        stat +=","+" Стать не встановлено";
                    }
                    //stat += "," + textBox84.Text;
                ReplaceWord("{n32}", textBox84.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n32}", " ", worddoc);
            }
                    
                if (textBox85.Text!="")
            {
                    if (textBox85.Text == "1")
                    {
                        stat +=","+" Чоловіча";
                    }
                    else if (textBox85.Text == "2")
                    {
                        stat +=","+" Жіноча";
                    }
                    else if (textBox85.Text == "3")
                    {
                        stat +=","+" Стать не встановлено";
                    }
                    //stat += "," + textBox85.Text;
                ReplaceWord("{n33}", textBox85.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n33}", " ", worddoc);
            }
                   
                if (textBox86.Text!="")
            {
                    if (textBox86.Text == "1")
                    {
                        stat +=","+" Чоловіча";
                    }
                    else if (textBox86.Text == "2")
                    {
                        stat +=","+" Жіноча";
                    }
                    else if (textBox86.Text == "3")
                    {
                        stat +=","+" Стать не встановлено";
                    }
                   // stat += "," + textBox86.Text;
                ReplaceWord("{n34}", textBox86.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n34}", " ", worddoc);
            }
                dict.Add("zag_stat_code", textBox10.Text+","+textBox83.Text+","+textBox84.Text+","+textBox85.Text+","+textBox86.Text);
                dict.Add("zag_stat_name", stat);
                ReplaceWord("{stat}", stat, worddoc);

                var status = " ";
                if (textBox87.Text!="")
            {
                status = CodeAllContent("social_status", "code_status", "name_status",textBox87.Text);
                
                    ReplaceWord("{n35}", textBox87.Text, worddoc);
                    ReplaceWord("{status1}", status, worddoc);
                }
            else
            {
                ReplaceWord("{n35}", " ", worddoc);
                    ReplaceWord("{status1}", " ", worddoc);
                }
                   
                if (textBox88.Text!="")
            {
                status = ","+CodeAllContent("social_status", "code_status", "name_status", textBox88.Text);
                    // CodeFromBase("social_status", "name_status", "code_status", "{n36}", textBox88.Text, worddoc);
                    ReplaceWord("{n36}", textBox88.Text, worddoc);
                    ReplaceWord("{status2}", status, worddoc);
                }
            else
            {
                ReplaceWord("{n36}", " ", worddoc);
                ReplaceWord("{status2}", " ", worddoc);
            }
                   
                if (textBox89.Text!="")
            {
                status =","+CodeAllContent("social_status", "code_status", "name_status", textBox89.Text);
                    // CodeFromBase("social_status", "name_status", "code_status", "{n37}", textBox89.Text, worddoc);
                    ReplaceWord("{n37}", textBox89.Text, worddoc);
                    ReplaceWord("{status3}", status, worddoc);
                }
            else
            {
                ReplaceWord("{n37}", " ", worddoc);
                    ReplaceWord("{status3}", " ", worddoc);
                }
                   
                if (textBox90.Text!="")
            {
                status =","+CodeAllContent("social_status", "code_status", "name_status", textBox90.Text);
                    // CodeFromBase("social_status", "name_status", "code_status", "{n38}",textBox90.Text, worddoc);
                    ReplaceWord("{n38}", textBox90.Text, worddoc);
                    ReplaceWord("{status4}", status, worddoc);
                }
            else
            {
                ReplaceWord("{n38}", " ", worddoc);
                    ReplaceWord("{status4}"," ", worddoc);
                }
                    
                if (textBox91.Text!="")
            {
                status = ","+CodeAllContent("social_status", "code_status", "name_status", textBox91.Text);
                    // CodeFromBase("social_status", "name_status", "code_status", "{n39}", textBox91.Text, worddoc);
                    ReplaceWord("{n39}", textBox91.Text, worddoc);
                    ReplaceWord("{status5}", status, worddoc);
                }
            else
            {
                ReplaceWord("{n39}", " ", worddoc);
                    ReplaceWord("{status5}", " ", worddoc);
                }
                dict.Add("code_status", textBox87.Text + "," + textBox88.Text + "," + textBox89.Text + "," + textBox90.Text + "," + textBox91.Text);
                dict.Add("name_status", status);
                //ReplaceWord("{status}", status, worddoc);

            var moment = " ";
            if (textBox92.Text!="")
            {
                    moment += CodeAllContent("moment_smerti", "code_moment", "moment", textBox92.Text); 
                //CodeFromBase("moment_smerti", "moment", "code_moment", "{n40}", textBox92.Text, worddoc);
                    ReplaceWord("{n40}", textBox92.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n40}", " ", worddoc);
            }
               
            if (textBox93.Text!="")
            {
                moment += "," + CodeAllContent("moment_smerti", "code_moment", "moment", textBox93.Text);
                    //CodeFromBase("moment_smerti", "moment", "code_moment", "{n41}", textBox93.Text, worddoc);
                    ReplaceWord("{n41}", textBox93.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n41}", " ", worddoc);
            }
              
            if (textBox94.Text!="")
            {
                moment += "," + CodeAllContent("moment_smerti", "code_moment", "moment", textBox94.Text);
                    // CodeFromBase("moment_smerti", "moment", "code_moment", "{n42}", textBox94.Text, worddoc);
                    ReplaceWord("{n42}", textBox94.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n42}", " ", worddoc);
            }
                
            if (textBox95.Text!="")
            {
                moment += "," + CodeAllContent("moment_smerti", "code_moment", "moment", textBox95.Text);
               // CodeFromBase("moment_smerti", "moment", "code_moment", "{n43}",textBox95.Text, worddoc);
                    ReplaceWord("{n43}", textBox95.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n43}", " ", worddoc);
            }
                
            if (textBox96.Text!="")
            {
                moment += "," + CodeAllContent("moment_smerti", "code_moment", "moment", textBox96.Text);
                    // CodeFromBase("moment_smerti", "moment", "code_moment", "{n44}", textBox96.Text, worddoc);
                    ReplaceWord("{n44}", textBox96.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n44}", " ", worddoc);
            }
                dict.Add("code_moment", textBox92.Text + "," + textBox93.Text + "," + textBox94.Text + "," + textBox95.Text + "," + textBox96.Text);
                dict.Add("moment", moment);
                ReplaceWord("{moment}", moment, worddoc);

            var umovi = " ";
            if (textBox97.Text!="")
            {
                umovi = CodeAllContent("umova_smerti", "code_umova", "name_umova", textBox97.Text);
                    ReplaceWord("{n45}", textBox97.Text, worddoc);
                    ReplaceWord("{umovi1}", umovi, worddoc);
                    // CodeFromBase("umova_smerti", "name_umova", "code_umova", "{n45}", textBox97.Text, worddoc);

                }
            else
            {
                ReplaceWord("{n45}", " ", worddoc);
                    ReplaceWord("{umovi1}", " ", worddoc);
                }
                
            if (textBox98.Text!="")
            {
                umovi =","+CodeAllContent("umova_smerti", "code_umova", "name_umova", textBox98.Text);
                    //CodeFromBase("umova_smerti", "name_umova", "code_umova", "{n46}", textBox98.Text, worddoc);
                    ReplaceWord("{n46}", textBox98.Text, worddoc);
                    ReplaceWord("{umovi2}", umovi, worddoc);
                }
            else
            {
                ReplaceWord("{n46}", " ", worddoc);
                    ReplaceWord("{umovi2}", " ", worddoc);
                }
                
            if (textBox99.Text!="")
            {
                umovi = ","+CodeAllContent("umova_smerti", "code_umova", "name_umova", textBox99.Text);
                    ReplaceWord("{n47}", textBox99.Text, worddoc);
                    ReplaceWord("{umovi3}", umovi, worddoc);
                    //CodeFromBase("umova_smerti", "name_umova", "code_umova", "{n47}", textBox99.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n47}", " ", worddoc);
                    ReplaceWord("{umovi3}", " ", worddoc);
                }
               
            if (textBox100.Text!="")
            {
                umovi = ","+CodeAllContent("umova_smerti", "code_umova", "name_umova", textBox100.Text);
                    ReplaceWord("{n48}", textBox100.Text, worddoc);
                    ReplaceWord("{umovi4}", umovi, worddoc);
                    //CodeFromBase("umova_smerti", "name_umova", "code_umova", "{n48}", textBox100.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n48}", " ", worddoc);
                    ReplaceWord("{umovi4}", " ", worddoc);
                }
                
            if (textBox101.Text!="")
            {
                umovi = ","+CodeAllContent("umova_smerti", "code_umova", "name_umova", textBox101.Text);
                    ReplaceWord("{n49}", textBox101.Text, worddoc);
                    ReplaceWord("{umovi5}", umovi, worddoc);
                    //CodeFromBase("umova_smerti", "name_umova", "code_umova", "{n49}", textBox101.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n49}", " ", worddoc);
                ReplaceWord("{umovi5}", " ", worddoc);
            }
                dict.Add("code_umovi", textBox97.Text + "," + textBox98.Text + "," + textBox99.Text + "," + textBox100.Text + "," + textBox101.Text);
                dict.Add("name_umovi", umovi);
               // ReplaceWord("{umovi}", umovi, worddoc);

            var tr = " ";
            if(textBox12.Text!= "")
                tr = textBox12.Text;
            ReplaceWord("{tr}", tr, worddoc);
            dict.Add("travm", tr);
            ReplaceWord("{n50}", tr, worddoc);
            var tr_d = " ";
            if (textBox11.Text != "")
                tr_d = textBox11.Text;
            ReplaceWord("{tr_d}", tr_d, worddoc);
            dict.Add("travm_ditei", tr_d);
            ReplaceWord("{n51}", tr_d, worddoc);
            var tr_fire = " ";
            if (textBox66.Text != "")
                tr_fire = textBox66.Text;
            ReplaceWord("{tr_fire}", tr_fire, worddoc);
            dict.Add("travm_fire", tr_fire);
            ReplaceWord("{n52}", tr_fire, worddoc);

            var zbitok = " ";
            if (textBox9.Text != "")
                zbitok = textBox9.Text;
            ReplaceWord("{zbitok}", zbitok, worddoc);
            dict.Add("pramiy", zbitok);
            ReplaceWord("{n53}", zbitok, worddoc);
            var pobich = " ";
            if (textBox15.Text != "")
                pobich = textBox15.Text;
            ReplaceWord("{pobich}", pobich, worddoc);
            dict.Add("pobichniy", pobich);
            ReplaceWord("{n54}", pobich, worddoc);
            var zn = " ";
            if (textBox7.Text != "")
                zn = textBox7.Text;
            ReplaceWord("{zn}", zn, worddoc);
            dict.Add("zn_bud", zn);
            ReplaceWord("{n55}", zn, worddoc);
            var posh = " ";
            if (textBox14.Text != "")
                posh = textBox14.Text;
            ReplaceWord("{posh}", posh, worddoc);
            dict.Add("posh_bud", posh);
            ReplaceWord("{n56}", posh, worddoc);
            var zn_t = " ";
            if (textBox13.Text != "")
                zn_t = textBox13.Text;
            ReplaceWord("{zn_t}", zn_t, worddoc);
            dict.Add("zn_tehnika", zn_t);
            ReplaceWord("{n57}", zn_t, worddoc);
            var posh_t = " ";
            if (textBox16.Text != "")
                posh_t = textBox16.Text;
            ReplaceWord("{posh_t}", posh_t, worddoc);
            dict.Add("posh_tehnika", posh_t);
            ReplaceWord("{n58}", posh_t, worddoc);
            var zn_z = " ";
            if (textBox28.Text != "")
                zn_z = textBox28.Text;
            ReplaceWord("{zn_z}", zn_z, worddoc);
            dict.Add("zn_zerno", zn_z);
            ReplaceWord("{n59}", zn_z, worddoc);
            var zn_h = " ";
            if (textBox27.Text != "")
                zn_h = textBox27.Text;
            ReplaceWord("{zn_h}", zn_h, worddoc);
            dict.Add("zn_koreni", zn_h);
            ReplaceWord("{n60}", zn_h, worddoc);
            var zn_v = " ";
            if (textBox36.Text != "")
                zn_v = textBox36.Text;
            ReplaceWord("{zn_v}", zn_v, worddoc);
            dict.Add("zn_valki", zn_v);
            ReplaceWord("{n61}", zn_v, worddoc);
            var zn_korm = " ";
            if (textBox30.Text != "")
                zn_korm = textBox30.Text;
            ReplaceWord("{zn_korm}", zn_korm, worddoc);
            dict.Add("zn_korm", zn_korm);
            ReplaceWord("{n62}", zn_korm, worddoc);
            var zn_torf = " ";
            if (textBox29.Text != "")
                zn_torf = textBox29.Text;
            ReplaceWord("{zn_torf}", zn_torf, worddoc);
            dict.Add("zn_torf", zn_torf);
            ReplaceWord("{n63}", zn_torf, worddoc);
            var posh_torf = " ";
            if (textBox35.Text != "")
                posh_torf = textBox35.Text;
            ReplaceWord("{posh_torf}", posh_torf, worddoc);
            dict.Add("posh_torf", posh_torf);
            ReplaceWord("{n64}", posh_torf, worddoc);
            var zn_tvarin = " ";
            if (textBox32.Text != "")
                zn_tvarin = textBox32.Text;
            ReplaceWord("{zn_tvarin}", zn_tvarin, worddoc);
            dict.Add("zag_tvarin", zn_tvarin);
            ReplaceWord("{n65}", zn_tvarin, worddoc);
            var zn_ptici = " ";
            if (textBox31.Text != "")
                zn_ptici = textBox31.Text;
            ReplaceWord("{zn_ptici}", zn_ptici, worddoc);
            dict.Add("zag_ptici", zn_ptici);
            ReplaceWord("{n66}", zn_ptici, worddoc);
            var dodatkovo_info = " ";
            if (textBox33.Text != "")
                dodatkovo_info = textBox33.Text;
            ReplaceWord("{dodatkovo_info}", dodatkovo_info, worddoc);
            dict.Add("dop_info", dodatkovo_info);
                /*4. Врятовано на пожежі*/
                var vr_l = " ";
            if (textBox37.Text != "")
                vr_l = textBox37.Text;
            ReplaceWord("{vr_l}", vr_l, worddoc);
            dict.Add("vr_ludei", vr_l);
            ReplaceWord("{n67}", vr_l, worddoc);
            var vr_d = " ";
            if (textBox34.Text != "")
                vr_d = textBox34.Text;
            ReplaceWord("{vr_d}", vr_d, worddoc);
            dict.Add("vr_ditei", vr_d);
            ReplaceWord("{n68}", vr_d, worddoc);
            var vr_t = "";
            if (textBox39.Text != "")
                vr_t = textBox39.Text;
            ReplaceWord("{vr_t}", vr_t, worddoc);
            dict.Add("vr_tvarin", vr_t);
            ReplaceWord("{n69}", vr_t, worddoc);
            var vr_p = " ";
            if (textBox38.Text != "")
                vr_p = textBox38.Text;
            ReplaceWord("{vr_p}", vr_p, worddoc);
            dict.Add("vr_ptici", vr_p);
            ReplaceWord("{n70}", vr_p, worddoc);
            var vr_b = " ";
            if (textBox41.Text != "")
                vr_b = textBox41.Text;
            ReplaceWord("{vr_b}", vr_b, worddoc);
            dict.Add("vr_bud", vr_b);
            ReplaceWord("{n71}", vr_b, worddoc);
            var vr_auto = " ";
            if (textBox40.Text != "")
                vr_auto = textBox40.Text;
            ReplaceWord("{vr_auto}", vr_auto, worddoc);
            dict.Add("vr_tehnika", vr_auto);
            ReplaceWord("{n72}", vr_auto, worddoc);
            var vr_cultur = " ";
            if (textBox44.Text != "")
                vr_cultur = textBox44.Text;
            ReplaceWord("{vr_cultur}", vr_cultur, worddoc);
            dict.Add("vr_zerno", vr_cultur);
            ReplaceWord("{n73}", vr_cultur, worddoc);
            var vr_h = " ";
            if (textBox46.Text != "")
                vr_h = textBox46.Text;
            ReplaceWord("{vr_h}", vr_h, worddoc);
            dict.Add("vr_koreni", vr_h);
            ReplaceWord("{n74}", vr_h, worddoc);
            var vr_v = " ";
            if (textBox45.Text != "")
                vr_v = textBox45.Text;
            ReplaceWord("{vr_v}", vr_v, worddoc);
            dict.Add("vr_valki", vr_v);
            ReplaceWord("{n75}", vr_v, worddoc);
            var vr_k = " ";
            if (textBox43.Text != "")
                vr_k = textBox43.Text;
            ReplaceWord("{vr_k}", vr_k, worddoc);
            dict.Add("vr_korm", vr_k);
            ReplaceWord("{n76}", vr_k, worddoc);
            var vr_torf = " ";
            if (textBox42.Text != "")
                vr_torf = textBox42.Text;
            ReplaceWord("{vr_torf}", vr_torf, worddoc);
            dict.Add("vr_torf", vr_torf);
            ReplaceWord("{n77}", vr_torf, worddoc);
            var vr_dodatkovo = " ";
            if (textBox47.Text != "")
                vr_dodatkovo = textBox47.Text;
            ReplaceWord("{vr_dodatkovo}", vr_dodatkovo, worddoc);
                dict.Add("vr_dop", vr_dodatkovo);
                var vr_mat = " ";
            if (textBox48.Text != "")
                vr_mat = textBox48.Text;
            ReplaceWord("{vr_mat}", vr_mat, worddoc);
                dict.Add("vr_mat", vr_mat);
                ReplaceWord("{n78}", vr_mat, worddoc);
            /*5. Розвиток і гасіння пожежі*/
            var day_p = dateTimePicker2.Value.Day;
            var month_p = dateTimePicker2.Value.Month;
            var year_p = dateTimePicker2.Value.Year;
            dict.Add("data_pov", dateTimePicker2.Value.ToShortDateString());
            ReplaceWord("{day_p}", day_p, worddoc);
            ReplaceWord("{month_p}", month_p, worddoc);
            ReplaceWord("{year_p}", year_p, worddoc);

            ReplaceWord("{n79}", day_p, worddoc);
            ReplaceWord("{n80}", month_p, worddoc);
            ReplaceWord("{n81}", year_p, worddoc);

            var hours_p = dateTimePicker3.Value.Hour;
            var minuts_p = dateTimePicker3.Value.Minute;
                if(hours_p!=00 || minuts_p!= 00)
                {
                    dict.Add("time_pov", hours_p + ":" + minuts_p);
                    ReplaceWord("{hours_p}", hours_p, worddoc);
                    ReplaceWord("{minuts_p}", minuts_p, worddoc);
                    ReplaceWord("{n82}", hours_p, worddoc);
                    ReplaceWord("{n83}", minuts_p, worddoc);
                }
                else if (hours_p == 00 && minuts_p == 00)
                {
                    dict.Add("time_pov", "");
                    ReplaceWord("{hours_p}", "", worddoc);
                    ReplaceWord("{minuts_p}", "", worddoc);
                    ReplaceWord("{n82}", "", worddoc);
                    ReplaceWord("{n83}", "", worddoc);
                }
            

           
            var hours_pr= dateTimePicker4.Value.Hour;
            var minuts_pr = dateTimePicker4.Value.Minute;
                if(hours_pr!=00 || minuts_pr != 00)
                {
                    dict.Add("time_pributa", hours_pr + ":" + minuts_pr);
                    ReplaceWord("{hours_pr}", hours_pr, worddoc);
                    ReplaceWord("{minuts_pr}", minuts_pr, worddoc);
                    ReplaceWord("{n84}", hours_pr, worddoc);
                    ReplaceWord("{n85}", minuts_pr, worddoc);
                }
                else if (hours_pr == 00 || minuts_pr == 00)
                {
                    dict.Add("time_pributa", "");
                    ReplaceWord("{hours_pr}", "", worddoc);
                    ReplaceWord("{minuts_pr}", "", worddoc);
                    ReplaceWord("{n84}", "", worddoc);
                    ReplaceWord("{n85}", "", worddoc);
                }
           

           
            var info_likv = " ";
            if (textBox102.Text!="")
            {
                info_likv = textBox102.Text;
                CodeFromBase("info_fire", "code_fire_likvid", "name_fire_likvid", "{info_likv}", info_likv, worddoc);
            }
            else
            {
                ReplaceWord("{n86}", " ", worddoc);
            }
            
            ReplaceWord("{n86}", info_likv, worddoc);

            var day_local = dateTimePicker6.Value.Day;
            var month_local = dateTimePicker6.Value.Month;
            var year_local = dateTimePicker6.Value.Year;

            if (dateTimePicker6.Value.ToString()!= "01.01.1900 0:00:00")
                {
                    dict.Add("data_lokal", dateTimePicker6.Value.ToShortDateString());
                    ReplaceWord("{day_local}", day_local, worddoc);
                    ReplaceWord("{month_local}", month_local, worddoc);
                    ReplaceWord("{year_local}", year_local, worddoc);

                    ReplaceWord("{n87}", day_local, worddoc);
                    ReplaceWord("{n88}", month_local, worddoc);
                    ReplaceWord("{n89}", year_local, worddoc);
                }
                else
                {
                    dict.Add("data_lokal", "");
                    ReplaceWord("{day_local}", "", worddoc);
                    ReplaceWord("{month_local}", "", worddoc);
                    ReplaceWord("{year_local}", "", worddoc);

                    ReplaceWord("{n87}", "", worddoc);
                    ReplaceWord("{n88}", "", worddoc);
                    ReplaceWord("{n89}", "", worddoc);
                }

          

            var hours_local = dateTimePicker5.Value.Hour;
            var minuts_local = dateTimePicker5.Value.Minute;
                if(hours_local!=00 || minuts_local != 00)
                {
                    ReplaceWord("{hours_local}", hours_local, worddoc);
                    ReplaceWord("{minuts_local}", minuts_local, worddoc);
                    dict.Add("time_lokal", hours_local + ":" + minuts_local);

                    ReplaceWord("{n90}", hours_local, worddoc);
                    ReplaceWord("{n91}", minuts_local, worddoc);
                }
                else if (hours_local == 00 && minuts_local == 00)
                {
                    ReplaceWord("{hours_local}", "", worddoc);
                    ReplaceWord("{minuts_local}", "", worddoc);
                    dict.Add("time_lokal", "");

                    ReplaceWord("{n90}", "", worddoc);
                    ReplaceWord("{n91}", "", worddoc);
                }
                var day_likvid = dateTimePicker8.Value.Day;
                var month_likvid = dateTimePicker8.Value.Month;
                var year_likvid = dateTimePicker8.Value.Year;

                if (dateTimePicker8.Value.ToString() != "01.01.1900 0:00:00")
                {
                    ReplaceWord("{day_likvid}", day_likvid, worddoc);
                    ReplaceWord("{month_likvid}", month_likvid, worddoc);
                    ReplaceWord("{year_likvid}", year_likvid, worddoc);
                    dict.Add("data_likvid", dateTimePicker8.Value.ToShortDateString());

                    ReplaceWord("{n92}", day_likvid, worddoc);
                    ReplaceWord("{n93}", month_likvid, worddoc);
                    ReplaceWord("{n94}", year_likvid, worddoc);
                }
                else
                {
                    ReplaceWord("{day_likvid}", "", worddoc);
                    ReplaceWord("{month_likvid}", "", worddoc);
                    ReplaceWord("{year_likvid}", "", worddoc);
                    dict.Add("data_likvid", "");

                    ReplaceWord("{n92}", "", worddoc);
                    ReplaceWord("{n93}", "", worddoc);
                    ReplaceWord("{n94}", "", worddoc);
                }
                   
           

            var hours_likvid = dateTimePicker7.Value.Hour;
            var minuts_likvid = dateTimePicker7.Value.Minute;
                if(hours_likvid!=00 || minuts_likvid != 00)
                {
                    ReplaceWord("{hours_likvid}", hours_likvid, worddoc);
                    ReplaceWord("{minuts_likvid}", minuts_likvid, worddoc);
                    dict.Add("time_likvid", hours_likvid + ":" + minuts_likvid);

                    ReplaceWord("{n95}", hours_likvid, worddoc);
                    ReplaceWord("{n96}", minuts_likvid, worddoc);
                }
                else if (hours_likvid == 00 && minuts_likvid == 00)
                {
                    ReplaceWord("{hours_likvid}", "", worddoc);
                    ReplaceWord("{minuts_likvid}", "", worddoc);
                    dict.Add("time_likvid", "");

                    ReplaceWord("{n95}", "", worddoc);
                    ReplaceWord("{n96}", "", worddoc);
                }
            

            var umovi_vpliv = " ";
            if (textBox103.Text!="")
            {
                    umovi_vpliv=CodeAllContent("poshirenya_fire", "code_poshireni", "umovi_poshireni",textBox103.Text); 
                    ReplaceWord("{n97}", textBox103.Text, worddoc);
                    ReplaceWord("{umovi_vpliv1}", umovi_vpliv, worddoc);
                    // CodeFromBase("poshirenya_fire", "umovi_poshireni", "code_poshireni", "{n97}", textBox103.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n97}", " ", worddoc);
                ReplaceWord("{umovi_vpliv1}", " ", worddoc);
                }
                
            if (textBox104.Text!="")
            {
                umovi_vpliv =","+CodeAllContent("poshirenya_fire", "code_poshireni", "umovi_poshireni", textBox104.Text);
                    ReplaceWord("{n98}", textBox104.Text, worddoc);
                    ReplaceWord("{umovi_vpliv2}", umovi_vpliv, worddoc);
                    //  CodeFromBase("poshirenya_fire", "umovi_poshireni", "code_poshireni", "{n98}", textBox104.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n98}", " ", worddoc);
                ReplaceWord("{umovi_vpliv2}", " ", worddoc);
            }
              
            if (textBox105.Text!="")
            {
                umovi_vpliv=","+CodeAllContent("poshirenya_fire", "code_poshireni", "umovi_poshireni", textBox105.Text);
                    ReplaceWord("{n99}", textBox105.Text, worddoc);
                    ReplaceWord("{umovi_vpliv3}", umovi_vpliv, worddoc);
                    // CodeFromBase("poshirenya_fire", "umovi_poshireni", "code_poshireni", "{n99}", textBox105.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n99}", " ", worddoc);
                ReplaceWord("{umovi_vpliv3}", " ", worddoc);
            }
               
            if (textBox106.Text!="")
            {
                umovi_vpliv =","+CodeAllContent("poshirenya_fire", "code_poshireni", "umovi_poshireni", textBox106.Text);
                    ReplaceWord("{n100}", textBox106.Text, worddoc);
                    ReplaceWord("{umovi_vpliv4}", umovi_vpliv, worddoc);
                    //CodeFromBase("poshirenya_fire", "umovi_poshireni", "code_poshireni", "{n100}", textBox106.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n100}", " ", worddoc);
                ReplaceWord("{umovi_vpliv4}", " ", worddoc);
            }
                
            if (textBox107.Text!="")
            {
                umovi_vpliv = ","+CodeAllContent("poshirenya_fire", "code_poshireni", "umovi_poshireni", textBox107.Text);
                    ReplaceWord("{n101}", textBox107.Text, worddoc);
                    ReplaceWord("{umovi_vpliv5}", umovi_vpliv, worddoc);
                    //CodeFromBase("poshirenya_fire", "umovi_poshireni", "code_poshireni", "{n101}", textBox107.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n101}", " ", worddoc);
                ReplaceWord("{umovi_vpliv5}", " ", worddoc);
            }
                dict.Add("code_umovi_posh", textBox103.Text + "," + textBox104.Text + "," + textBox105.Text + "," + textBox106.Text + "," + textBox107.Text);
                dict.Add("name_umovi_posh",umovi_vpliv);
           // ReplaceWord("{umovi_vpliv}", umovi_vpliv, worddoc);

            var umovi_uskl = " ";
            if (textBox112.Text!="")
            {
                    umovi_uskl= CodeAllContent("uskladnenya_fire", "code_uskl", "name_uskl", textBox112.Text);
                ReplaceWord("{n102}", textBox112.Text, worddoc);
                    ReplaceWord("{umovi_uskl1}", umovi_uskl, worddoc);
                    //CodeFromBase("uskladnenya_fire", "name_uskl", "code_uskl", "{n102}", textBox112.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n102}", " ", worddoc);
                    ReplaceWord("{umovi_uskl1}", " ", worddoc);
                }
               
            if (textBox111.Text!="")
            {
                umovi_uskl = "," + CodeAllContent("uskladnenya_fire", "code_uskl", "name_uskl", textBox111.Text);
                    ReplaceWord("{n103}", textBox111.Text, worddoc);
                    ReplaceWord("{umovi_uskl2}", umovi_uskl, worddoc);
                    //CodeFromBase("uskladnenya_fire", "name_uskl", "code_uskl", "{n103}", textBox111.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n103}", " ", worddoc);
                    ReplaceWord("{umovi_uskl2}", " ", worddoc);
                }
                
            if (textBox110.Text!="")
            {
                umovi_uskl = "," + CodeAllContent("uskladnenya_fire", "code_uskl", "name_uskl", textBox110.Text);
                    ReplaceWord("{n104}", textBox110.Text, worddoc);
                    ReplaceWord("{umovi_uskl3}", umovi_uskl, worddoc);
                    //CodeFromBase("uskladnenya_fire", "name_uskl", "code_uskl", "{n104}", textBox110.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n104}", " ", worddoc);
                    ReplaceWord("{umovi_uskl3}", " ", worddoc);
                }
                
            if (textBox109.Text!="")
            {
                umovi_uskl = "," + CodeAllContent("uskladnenya_fire", "code_uskl", "name_uskl", textBox109.Text);
                    ReplaceWord("{n105}", textBox109.Text, worddoc);
                    ReplaceWord("{umovi_uskl4}", umovi_uskl, worddoc);
                    //CodeFromBase("uskladnenya_fire", "name_uskl", "code_uskl", "{n105}", textBox109.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n105}", " ", worddoc);
                    ReplaceWord("{umovi_uskl4}", " ", worddoc);
                }
                
            if (textBox108.Text!="")
            {
                umovi_uskl = "," + CodeAllContent("uskladnenya_fire", "code_uskl", "name_uskl", textBox108.Text);
                    ReplaceWord("{n106}", textBox108.Text, worddoc);
                    ReplaceWord("{umovi_uskl5}", umovi_uskl, worddoc);
                    //CodeFromBase("uskladnenya_fire", "name_uskl", "code_uskl", "{n106}", textBox108.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n106}", " ", worddoc);
                    ReplaceWord("{umovi_uskl5}", " ", worddoc);
                }
                dict.Add("code_umovi_uskl", textBox112.Text + "," + textBox111.Text + "," + textBox110.Text + "," + textBox109.Text + "," + textBox108.Text);
                dict.Add("name_umovi_uskl", umovi_uskl);
              //  ReplaceWord("{umovi_uskl}", umovi_uskl, worddoc);
           

            var spz = " ";
            if (textBox113.Text!="")
            {
                spz = textBox113.Text;
                CodeFromBase("nayavnist_spz", "code_spz", "name_spz", "{spz}", spz, worddoc);
            }
            else
            {
                ReplaceWord("{n107}", " ", worddoc);
                ReplaceWord("{spz}", " ", worddoc);
                    dict.Add("code_spz", "");
                    dict.Add("name_spz", "");
                }
               
            ReplaceWord("{n107}", spz, worddoc);

            var system_fire = " ";
            if (textBox118.Text!="")
            {
                system_fire += CodeAllContent("system_protipojeji", "code_system", "name_system", textBox118.Text);
                ReplaceWord("{n108}", textBox118.Text, worddoc);
                    // CodeFromBase("system_protipojeji", "name_system", "code_system", "{n108}", textBox118.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n108}", " ", worddoc);
            }
                
            if (textBox117.Text!="")
            {
                system_fire += "," + CodeAllContent("system_protipojeji", "code_system", "name_system", textBox117.Text);
                ReplaceWord("{n109}", textBox117.Text, worddoc);
                    // CodeFromBase("system_protipojeji", "name_system", "code_system", "{n109}", textBox117.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n109}", " ", worddoc);
            }
               
            if (textBox116.Text!="")
            {
                system_fire += "," + CodeAllContent("system_protipojeji", "code_system", "name_system", textBox116.Text);
                ReplaceWord("{n110}", textBox116.Text, worddoc);
                //CodeFromBase("system_protipojeji", "name_system", "code_system", "{n110}", textBox116.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n110}", " ", worddoc);
            }
               
            if (textBox115.Text!="")
            {
                system_fire += "," + CodeAllContent("system_protipojeji", "code_system", "name_system", textBox115.Text);
                ReplaceWord("{n111}", textBox115.Text, worddoc);
                //CodeFromBase("system_protipojeji", "name_system", "code_system", "{n111}", textBox115.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n111}", " ", worddoc);
            }
                
            if (textBox114.Text!="")
            {
                system_fire += "," + CodeAllContent("system_protipojeji", "code_system", "name_system", textBox114.Text);
                ReplaceWord("{n112}", textBox114.Text, worddoc);
                    //CodeFromBase("system_protipojeji", "name_system", "code_system", "{n112}", textBox114.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n112}", " ", worddoc);
            }
                dict.Add("code_system", textBox118.Text + "," + textBox117.Text + "," + textBox116.Text + "," + textBox115.Text + "," + textBox114.Text);
                dict.Add("name_system", system_fire);
                ReplaceWord("{system_fire}", system_fire, worddoc);

            var res_system = " ";
            if (textBox123.Text!="")
            {
                    res_system = CodeAllContent("resultat_dii_system", "code_res", "name_res", textBox123.Text);
                ReplaceWord("{n113}", textBox123.Text, worddoc);
                    ReplaceWord("{res_system1}", res_system, worddoc);
                    //CodeFromBase("resultat_dii_system", "name_res", "code_res", "{n113}", textBox123.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n113}", " ", worddoc);
                    ReplaceWord("{res_system1}", " ", worddoc);
                }
             
            if (textBox122.Text!="")
            {
                res_system = "," + CodeAllContent("resultat_dii_system", "code_res", "name_res", textBox122.Text);
                    ReplaceWord("{n114}", textBox122.Text, worddoc);
                    ReplaceWord("{res_system2}", res_system, worddoc);
                    //CodeFromBase("resultat_dii_system", "name_res", "code_res", "{n114}", textBox122.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n114}", " ", worddoc);
                    ReplaceWord("{res_system2}", " ", worddoc);
                }
                
            if (textBox121.Text!="")
            {
                res_system = "," + CodeAllContent("resultat_dii_system", "code_res", "name_res", textBox121.Text);
                    ReplaceWord("{n115}", textBox121.Text, worddoc);
                    ReplaceWord("{res_system3}", res_system, worddoc);
                    // CodeFromBase("resultat_dii_system", "name_res", "code_res", "{n115}", textBox121.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n115}", " ", worddoc);
                    ReplaceWord("{res_system3}", " ", worddoc);
                }
               
            if (textBox120.Text!="")
            {
                res_system = "," + CodeAllContent("resultat_dii_system", "code_res", "name_res", textBox120.Text);
                    ReplaceWord("{n116}", textBox120.Text, worddoc);
                    ReplaceWord("{res_system4}", res_system, worddoc);
                    //CodeFromBase("resultat_dii_system", "name_res", "code_res", "{n116}", textBox120.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n116}", " ", worddoc);
                    ReplaceWord("{res_system4}", " ", worddoc);
                }
               
            if (textBox119.Text!="")
            {
                res_system = "," + CodeAllContent("resultat_dii_system", "code_res", "name_res", textBox119.Text);
                    ReplaceWord("{n117}", textBox119.Text, worddoc);
                    ReplaceWord("{res_system5}", res_system, worddoc);
                    //CodeFromBase("resultat_dii_system", "name_res", "code_res", "{n117}", textBox119.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n117}", " ", worddoc);
                    ReplaceWord("{res_system5}", " ", worddoc);
                }
                dict.Add("code_dii", textBox123.Text + "," + textBox122.Text + "," + textBox121.Text + "," + textBox120.Text + "," + textBox119.Text);
                dict.Add("name_dii", res_system);
               // ReplaceWord("{res_system}", res_system, worddoc);

            /*6. Сили та засоби гасіння пожежі*/
            var uch_fire = " ";
            if (textBox128.Text!="")
            {
                    uch_fire += CodeAllContent("uchasnik_fire", "code_uchasnik", "name_uchasnik", textBox128.Text);
                ReplaceWord("{n118}", textBox128.Text, worddoc);
                    //CodeFromBase("uchasnik_fire", "name_uchasnik", "code_uchasnik", "{n118}", textBox128.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n118}", " ", worddoc);
            }
                
            if (textBox127.Text!="")
            {
                uch_fire += "," + CodeAllContent("uchasnik_fire", "code_uchasnik", "name_uchasnik", textBox127.Text);
                    ReplaceWord("{n119}", textBox127.Text, worddoc);
                    //CodeFromBase("uchasnik_fire", "name_uchasnik", "code_uchasnik", "{n119}", textBox127.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n119}", " ", worddoc);
            }
                
            if (textBox126.Text!="")
            {
                uch_fire += "," + CodeAllContent("uchasnik_fire", "code_uchasnik", "name_uchasnik", textBox126.Text);
                    ReplaceWord("{n120}", textBox126.Text, worddoc);
                    //CodeFromBase("uchasnik_fire", "name_uchasnik", "code_uchasnik", "{n120}", textBox126.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n120}", " ", worddoc);
            }
                
            if (textBox125.Text!="")
            {
                uch_fire += "," + CodeAllContent("uchasnik_fire", "code_uchasnik", "name_uchasnik", textBox125.Text);
                    ReplaceWord("{n121}", textBox125.Text, worddoc);
                    // CodeFromBase("uchasnik_fire", "name_uchasnik", "code_uchasnik", "{n121}", textBox125.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n121}", " ", worddoc);
            }
                
            if (textBox124.Text!="")
            {
                uch_fire += "," + CodeAllContent("uchasnik_fire", "code_uchasnik", "name_uchasnik", textBox124.Text);
                    ReplaceWord("{n122}", textBox124.Text, worddoc);
                //CodeFromBase("uchasnik_fire", "name_uchasnik", "code_uchasnik", "{n122}", textBox124.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n122}", " ", worddoc);
            }
                dict.Add("code_uch", textBox128.Text + "," + textBox127.Text + "," + textBox126.Text + "," + textBox125.Text + "," + textBox124.Text);
                dict.Add("name_uch", uch_fire);
                ReplaceWord("{uch_fire}", uch_fire, worddoc);

            var kilkist_uch = " ";
            if (textBox49.Text != "")
            {
                kilkist_uch += textBox49.Text;
                ReplaceWord("{n123}", textBox49.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n123}", " ", worddoc);
            }
                
            if (textBox50.Text != "")
            {
                kilkist_uch += "," + textBox50.Text;
                ReplaceWord("{n124}", textBox50.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n124}", " ", worddoc);
            }
                
            if (textBox51.Text != "")
            {
                kilkist_uch += "," + textBox51.Text;
                ReplaceWord("{n125}", textBox51.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n125}", " ", worddoc);
            }
                
            if (textBox52.Text != "")
            {
                kilkist_uch += "," + textBox52.Text;
                ReplaceWord("{n126}", textBox52.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n126}", " ", worddoc);
            }
                
            if (textBox53.Text != "")
            {
                kilkist_uch += "," + textBox53.Text;
                ReplaceWord("{n127}", textBox53.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n127}", " ", worddoc);
            }
                dict.Add("kilk_uch", textBox49.Text+","+textBox50.Text+","+textBox51.Text+","+textBox52.Text+","+textBox53.Text);
                ReplaceWord("{kilkist_uch}", kilkist_uch, worddoc);

            var fire_tehnika = " ";
            if (textBox129.Text!="")
            {
                    fire_tehnika += CodeAllContent("fire_auto", "code_auto", "name_auto", textBox129.Text);
                ReplaceWord("{n128}", textBox129.Text, worddoc);
               //CodeFromBase("fire_auto", "name_auto", "code_auto", "{n128}", textBox129.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n128}", " ", worddoc);
            }
                
            if (textBox130.Text!="")
            {
                fire_tehnika += "," + CodeAllContent("fire_auto", "code_auto", "name_auto", textBox130.Text);
                    ReplaceWord("{n129}", textBox130.Text, worddoc);
                    //CodeFromBase("fire_auto", "name_auto", "code_auto", "{n129}",textBox130.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n129}", " ", worddoc);
            }
                
            if (textBox131.Text!="")
            {
                fire_tehnika += "," + CodeAllContent("fire_auto", "code_auto", "name_auto", textBox131.Text);
                    ReplaceWord("{n130}", textBox131.Text, worddoc);
                    //CodeFromBase("fire_auto", "name_auto", "code_auto", "{n130}", textBox131.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n130}", " ", worddoc);
            }
                
            if (textBox132.Text!="")
            {
                fire_tehnika += "," + CodeAllContent("fire_auto", "code_auto", "name_auto", textBox132.Text);
                    ReplaceWord("{n131}", textBox132.Text, worddoc);
                   // CodeFromBase("fire_auto", "name_auto", "code_auto", "{n131}", textBox132.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n131}", " ", worddoc);
            }
                
            if (textBox133.Text!="")
            {
                fire_tehnika += "," + CodeAllContent("fire_auto", "code_auto", "name_auto", textBox133.Text);
                    ReplaceWord("{n132}", textBox133.Text, worddoc);
                    //CodeFromBase("fire_auto", "name_auto", "code_auto", "{n132}",textBox133.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n132}", " ", worddoc);
            }
                dict.Add("code_fireauto", textBox129.Text + "," + textBox130.Text + "," + textBox131.Text + "," + textBox132.Text + "," + textBox133.Text);
                dict.Add("name_fireauto", fire_tehnika);
                ReplaceWord("{fire_tehnika}", fire_tehnika, worddoc);

            var kilkist_fire = " ";
            if (textBox58.Text != "")
            {
                kilkist_fire += textBox58.Text;
                ReplaceWord("{n133}", textBox58.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n133}", " ", worddoc);
            }
                
            if (textBox57.Text != "")
            {
                kilkist_fire += "," + textBox57.Text;
                ReplaceWord("{n134}", textBox57.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n134}", " ", worddoc);
            }
               
            if (textBox56.Text != "")
            {
                kilkist_fire += "," + textBox56.Text;
                ReplaceWord("{n135}", textBox56.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n135}", " ", worddoc);
            }
                
            if (textBox55.Text != "")
            {
                kilkist_fire += "," + textBox55.Text;
                ReplaceWord("{n136}", textBox55.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n136}", " ", worddoc);
            }
               
            if (textBox54.Text != "")
            {
                kilkist_fire += "," + textBox54.Text;
                ReplaceWord("{n137}", textBox54.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n137}", " ", worddoc);
            }
                dict.Add("kilk_auto", textBox58.Text+","+textBox57.Text+","+textBox56.Text+","+textBox55.Text+","+textBox54.Text);
                ReplaceWord("{kilkist_fire}", kilkist_fire, worddoc);

            var fire_stvol = " ";
            if (textBox134.Text!="")
            {
                    fire_stvol += CodeAllContent("fire_stvoli", "code_stvol", "name_stvol", textBox134.Text);
                ReplaceWord("{n139}", textBox134.Text, worddoc);
                //CodeFromBase("fire_stvoli", "name_stvol", "code_stvol", "{n139}", textBox134.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n139}", " ", worddoc);
            }
                
            if (textBox135.Text!="")
            {
                fire_stvol += "," + CodeAllContent("fire_stvoli", "code_stvol", "name_stvol", textBox135.Text);
                    ReplaceWord("{n140}", textBox135.Text, worddoc);
                    //CodeFromBase("fire_stvoli", "name_stvol", "code_stvol", "{n140}", textBox135.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n140}", " ", worddoc);
            }
              
            if (textBox136.Text!="")
            {
                fire_stvol += "," + CodeAllContent("fire_stvoli", "code_stvol", "name_stvol", textBox136.Text);
                    ReplaceWord("{n141}", textBox136.Text, worddoc);
                    //CodeFromBase("fire_stvoli", "name_stvol", "code_stvol", "{n141}", textBox136.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n141}", " ", worddoc);
            }
                dict.Add("code_firestvol", textBox134.Text + "," + textBox135.Text + "," + textBox136.Text);
                dict.Add("name_firestvol", fire_stvol);
                ReplaceWord("{fire_stvol}", fire_stvol, worddoc);

            var kilkist_stvol = " ";
            if (textBox61.Text != "")
            {
                kilkist_stvol += textBox61.Text;
                ReplaceWord("{n142}", textBox61.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n142}", " ", worddoc);
            }
                
            if (textBox60.Text != "")
            {
                kilkist_stvol += "," + textBox60.Text;
                ReplaceWord("{n143}", textBox60.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n143}", " ", worddoc);
            }
                
            if (textBox59.Text != "")
            {
                kilkist_stvol += "," + textBox59.Text;
                ReplaceWord("{n144}", textBox59.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n144}", " ", worddoc);
            }
                dict.Add("kilk_firestvol", textBox61.Text+","+textBox60.Text+","+textBox59.Text);
                ReplaceWord("{kilkist_stvol}", kilkist_stvol, worddoc);

            var vogn_rech = " ";
            if (textBox137.Text!="")
            {
                    vogn_rech += CodeAllContent("vognegasni_rechovini", "code_rechovini", "name_rechovini", textBox137.Text);
                ReplaceWord("{n145}", textBox137.Text, worddoc);
                    //CodeFromBase("vognegasni_rechovini", "name_rechovini", "code_rechovini", "{n145}", textBox137.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n145}", " ", worddoc);
            }
                
            if (textBox138.Text!="")
            {
                vogn_rech += "," + CodeAllContent("vognegasni_rechovini", "code_rechovini", "name_rechovini", textBox138.Text);
                    ReplaceWord("{n146}", textBox138.Text, worddoc);
                    //CodeFromBase("vognegasni_rechovini", "name_rechovini", "code_rechovini", "{n146}", textBox138.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n146}", " ", worddoc);
            }
                
            if (textBox139.Text!="")
            {
                vogn_rech += "," + CodeAllContent("vognegasni_rechovini", "code_rechovini", "name_rechovini", textBox139.Text);
                    ReplaceWord("{n147}", textBox139.Text, worddoc);
               //CodeFromBase("vognegasni_rechovini", "name_rechovini", "code_rechovini", "{n147}", textBox139.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n147}", " ", worddoc);
            }
                dict.Add("code_rechovini", textBox137.Text + "," + textBox138.Text + "," + textBox139.Text);
                dict.Add("name_rechovini", vogn_rech);
                ReplaceWord("{vogn_rech}", vogn_rech, worddoc);

            var perv_zacobi = " ";
            if (textBox140.Text!="")
            {
                    perv_zacobi += CodeAllContent("zacobi_fire", "code_zasobi", "name_zasobi", textBox140.Text);
                ReplaceWord("{n148}", textBox140.Text, worddoc);
                    //CodeFromBase("zacobi_fire", "name_zasobi", "code_zasobi", "{n148}", textBox140.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n148}", " ", worddoc);
            }
                
            if (textBox141.Text!="")
            {
                perv_zacobi += "," + CodeAllContent("zacobi_fire", "code_zasobi", "name_zasobi", textBox141.Text);
                    ReplaceWord("{n149}", textBox141.Text, worddoc);
                    //CodeFromBase("zacobi_fire", "name_zasobi", "code_zasobi", "{n149}", textBox141.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n149}", " ", worddoc);
            }
               
            if (textBox142.Text!="")
            {
                perv_zacobi += "," + CodeAllContent("zacobi_fire", "code_zasobi", "name_zasobi", textBox142.Text);
                    ReplaceWord("{n150}", textBox142.Text, worddoc);
                    // CodeFromBase("zacobi_fire", "name_zasobi", "code_zasobi", "{n150}", textBox142.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n150}", " ", worddoc);
            }
                dict.Add("code_pervini", textBox140.Text + "," + textBox141.Text + "," + textBox142.Text);
                dict.Add("name_pervini", perv_zacobi);
                ReplaceWord("{perv_zacobi}", perv_zacobi, worddoc);

            var djerelo_voda = " ";
            if (textBox143.Text!="")
            {
                    djerelo_voda += CodeAllContent("djerela_vodopostachanaya", "code_djerela", "name_djerela",textBox143.Text);
                    ReplaceWord("{n151}", textBox143.Text, worddoc);
                    // CodeFromBase("djerela_vodopostachanaya", "name_djerela", "code_djerela", "{n151}", textBox143.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n151}", " ", worddoc);
            }
                
            if (textBox144.Text!="")
            {
                djerelo_voda += "," + CodeAllContent("djerela_vodopostachanaya", "code_djerela", "name_djerela", textBox144.Text);
                    ReplaceWord("{n152}", textBox144.Text, worddoc);
                    //  CodeFromBase("djerela_vodopostachanaya", "name_djerela", "code_djerela", "{n152}", textBox144.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n152}", " ", worddoc);
            }
                
            if (textBox145.Text!="")
            {
                djerelo_voda += "," + CodeAllContent("djerela_vodopostachanaya", "code_djerela", "name_djerela", textBox145.Text);
                    ReplaceWord("{n153}", textBox145.Text, worddoc);
                    // CodeFromBase("djerela_vodopostachanaya", "name_djerela", "code_djerela", "{n153}", textBox145.Text, worddoc);
                }
            else
            {
                ReplaceWord("{n153}", " ", worddoc);
            }
                dict.Add("code_djerela", textBox143.Text + "," + textBox144.Text + "," + textBox145.Text);
                dict.Add("name_djerela", djerelo_voda);
                ReplaceWord("{djerelo_voda}", djerelo_voda, worddoc);

            var gds = " ";
            if (textBox62.Text != "")
                gds = textBox62.Text;
            ReplaceWord("{gds}", gds, worddoc);
                dict.Add("vikr_gds",gds);
                ReplaceWord("{n154}", gds, worddoc);
            var kilk_lanok = " ";
            if (textBox63.Text != "")
                kilk_lanok = textBox63.Text;
            ReplaceWord("{kilk_lanok}", kilk_lanok, worddoc);
            dict.Add("kilk_gds", kilk_lanok);
            ReplaceWord("{n155}", kilk_lanok, worddoc);
            var all_work = " ";
            if (textBox64.Text != "")
                all_work = textBox64.Text;
            ReplaceWord("{all_work}", all_work, worddoc);
            dict.Add("time_gds", all_work);
            ReplaceWord("{n156}", all_work, worddoc);
            /*7. Заходи державного нагляду*/
            var day_last = dateTimePicker9.Value.Day;
            var month_last = dateTimePicker9.Value.Month;
            var year_last = dateTimePicker9.Value.Year;
            if (dateTimePicker9.Value.ToString() != "01.01.1900 0:00:00")
            {
                    dict.Add("data_perevirki", dateTimePicker9.Value.ToShortDateString());

                    ReplaceWord("{day_last}", day_last, worddoc);
                    ReplaceWord("{month_last}", month_last, worddoc);
                    ReplaceWord("{year_last}", year_last, worddoc);

                    ReplaceWord("{n157}", day_last, worddoc);
                    ReplaceWord("{n158}", month_last, worddoc);
                    ReplaceWord("{n159}", year_last, worddoc);
                }
                else
                {
                    dict.Add("data_perevirki", "");

                    ReplaceWord("{day_last}", "", worddoc);
                    ReplaceWord("{month_last}", "", worddoc);
                    ReplaceWord("{year_last}", "", worddoc);

                    ReplaceWord("{n157}", "", worddoc);
                    ReplaceWord("{n158}", "", worddoc);
                    ReplaceWord("{n159}", "", worddoc);
                }
                   
            var vid_perev = "";
            if (textBox146.Text!="")
            {
                vid_perev = textBox146.Text;
                CodeFromBase("perevirka_fire", "code_perevirka", "name_perevirka", "{vid_perev}", vid_perev, worddoc);
            }
            else
            {
                ReplaceWord("{n160}", " ", worddoc);
                ReplaceWord("{vid_perev}", " ", worddoc);
                    dict.Add("code_perevirka", "");
                    dict.Add("name_perevirka", "");
                }
                
            ReplaceWord("{n160}", vid_perev, worddoc);
            var umova_dialnist = "";
            if (textBox147.Text!="")
            {
                umova_dialnist = textBox147.Text;
                CodeFromBase("gospodar_dialnist", "code_dialnist", "name_dialnist", "{umova_dialnist}", umova_dialnist, worddoc);
            }
            else
            {
                ReplaceWord("{n161}", " ", worddoc);
                ReplaceWord("{umova_dialnist}", " ", worddoc);
                    dict.Add("code_dialnist", "");
                    dict.Add("name_dialnist", "");
                }
               
            ReplaceWord("{n161}", umova_dialnist, worddoc);
            var fact_zahodi = "";
            if (textBox148.Text!="")
            {
                    fact_zahodi += CodeAllContent("zahodi_pojeji", "code_zahodi", "name_zahodi", textBox148.Text);
                ReplaceWord("{n162}", textBox148.Text, worddoc);
                
                //CodeFromBase("zahodi_pojeji", "code_zahodi", "name_zahodi", "{n162}", textBox148.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n162}", " ", worddoc);
            }
            if (textBox149.Text!="")
            {
                fact_zahodi += ",кримінальне провадження за " + textBox159.Text;
                    ReplaceWord("{n163}", textBox149.Text, worddoc);
                //CodeFromBase("zahodi_pojeji", "code_zahodi", "name_zahodi", "{n163}", textBox149.Text, worddoc);
            }
            else
            {
                ReplaceWord("{n163}", " ", worddoc);
            }
                dict.Add("code_zahodi", textBox148.Text+","+textBox149.Text);
                dict.Add("name_zahodi", fact_zahodi);
              //  dict.Add("name_dialnist", "");
                ReplaceWord("{fact_zahodi}", fact_zahodi, worddoc);
                dict.Add("number_kku", textBox159.Text);
                var day_zap = dateTimePicker10.Value.Day;
            var month_zap = dateTimePicker10.Value.Month;
            var year_zap = dateTimePicker10.Value.Year;
            dict.Add("data_zapovnenya", dateTimePicker10.Value.ToShortDateString());

            ReplaceWord("{day_zap}", day_zap, worddoc);
            ReplaceWord("{month_zap}", month_zap, worddoc);
            ReplaceWord("{year_zap}", year_zap, worddoc);

            ReplaceWord("{n164}", day_zap, worddoc);
            ReplaceWord("{n165}", month_zap, worddoc);
            ReplaceWord("{n166}", year_zap, worddoc);

                dict.Add("pid_osibi", textBox150.Text);
                //   if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                //    {
                // if ((saveFileDialog1.OpenFile()) != null)
                //  {
                worddoc.SaveAs(filename2);

                       
                  //  }
              //  }
               //myStream.Close();  
                worddoc.Close();
                panel6.Visible = false;
            MessageBox.Show("Картка обліку пожежі створена");

               
                SaveInDb save = new SaveInDb(dict,flag_edit);
                dict.Clear();
                flag_edit = false;
            }
            else
            {
                MessageBox.Show("Заповніть всі необхідні поля");
            }
            //}
            /*catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }*/

            // worddoc.Close();
        }
 public void CodeFromBase(string tableitem, string namefield, string code, object stubToReplace, object text, Word.Document wordDocument)
        {
            string select_code = "SELECT `" + code + "` FROM " + tableitem + " WHERE " + namefield + "= '" + text.ToString()+"'";
            string content = "";
            //"SELECT `code_region` FROM region WHERE name_region='Донецька область'"



            //connection();
            try
            {
               // if (con_db.State == ConnectionState.Open)
               // {
                    cmd_db = new SQLiteCommand(select_code, con_db);
                    rdr = cmd_db.ExecuteReader();

                    while (rdr.Read())
                    {
                    content = rdr[0].ToString();
                        ReplaceWord(stubToReplace, rdr[0].ToString(), wordDocument);
                        
                    }
                dict.Add(code, content);
                // }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
         //   Disconnect();
            dict.Add(namefield, text.ToString());
        }
public string CodeAllContent(string tableitem, string namefield, string code,  object text)
        {
            string content = "";
            string select_code = "SELECT `" + code + "` FROM " + tableitem + " WHERE " + namefield + "= '" + text.ToString() + "'";
          //  connection();
            try
            {
                if (con_db.State == ConnectionState.Open)
                {
                    cmd_db = new SQLiteCommand(select_code, con_db);
                    rdr = cmd_db.ExecuteReader();

                    while (rdr.Read())
                    {
                        content = rdr[0].ToString();
                       // dict.Add(code, rdr[0].ToString());

                    }
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
         //   Disconnect();
           // dict.Add(namefield, text.ToString());
            return content;
        }
 public void ReplaceWord(object stubToReplace, object text, Word.Document wordDocument)
    {
        var range = wordDocument.Content;
        range.Find.ClearFormatting();
        
        range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
            
      
           
    }
        public void ReadDb() {
          //  connection();
            try {
                if (con_db.State == ConnectionState.Open)
                {
                    /*таблица 1 регионы*/
                    cmd_db = new SQLiteCommand("SELECT * from region", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox84.Items.Clear();
                    
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        //comboBox1.Items.Add((rdr[1].ToString()));
                        comboBox84.Items.Add((rdr[1].ToString()));
                       
                    }

                    /*таблица 2 тип населенного пункта*/
                    cmd_db = new SQLiteCommand("SELECT * from type_region", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox2.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox2.Items.Add((rdr[1].ToString()));
                    }
                    /*таблица 3 район*/
                    comboBox83.Items.Clear();
                    listBox1.Items.Clear();
                    cmd_db = new SQLiteCommand("SELECT * from current_raion", con_db);
                    rdr = cmd_db.ExecuteReader();
                    
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                       
                        comboBox83.Items.Add((rdr[2].ToString()));
                        listBox1.Items.Add(rdr[1].ToString()+";"+rdr[2].ToString());
                    }
                    /*таблица 3-1 коди населенных пунктов*/
                    comboBox85.Items.Clear();
                    listBox2.Items.Clear();
                    cmd_db = new SQLiteCommand("SELECT * from current_np", con_db);
                    rdr = cmd_db.ExecuteReader();

                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox85.Items.Add((rdr[2].ToString()));
                        listBox2.Items.Add(rdr[1].ToString() + ";" + rdr[2].ToString());
                    }

                    /*таблица 4 тип форма власності*/
                    cmd_db = new SQLiteCommand("SELECT * from forma_vlasnosti", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox5.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox5.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 5 Ступінь ризику господарської діяльності*/
                    cmd_db = new SQLiteCommand("SELECT * from stupin_riziku", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox6.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox6.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 6 Підконтрольність об’єкта*/
                    cmd_db = new SQLiteCommand("SELECT * from pidkontrol_object", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox7.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox7.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 7 Поверховість*/
                    cmd_db = new SQLiteCommand("SELECT * from poverhovist", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox8.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox8.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 8 Ступінь вогнестійкості будинку*/
                    cmd_db = new SQLiteCommand("SELECT * from vognestoikist", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox9.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox9.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 9 Категорія небезпеки*/
                    cmd_db = new SQLiteCommand("SELECT * from category_nebespeki", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox10.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox10.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 10 Місце виникнення пожежі*/
                    cmd_db = new SQLiteCommand("SELECT * from place_fire", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox11.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox11.Items.Add((rdr[2].ToString()+";"+rdr[1].ToString()));
                    }

                    /*таблица 12 Причина пожежі*/
                    cmd_db = new SQLiteCommand("SELECT * from pricini_fire", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox24.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox24.Items.Add((rdr[2].ToString()+";"+rdr[1].ToString()));
                        
                    }

                    /*таблица 14 соціальний статус*/
                    cmd_db = new SQLiteCommand("SELECT * from social_status", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox19.Items.Clear();
                    comboBox20.Items.Clear();
                    comboBox21.Items.Clear();
                    comboBox22.Items.Clear();
                    comboBox23.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox19.Items.Add((rdr[2].ToString()+";"+rdr[1].ToString()));
                        comboBox20.Items.Add((rdr[2].ToString()+ ";"+rdr[1].ToString()));
                        comboBox21.Items.Add((rdr[2].ToString() + ";" + rdr[1].ToString()));
                        comboBox22.Items.Add((rdr[2].ToString() + ";" + rdr[1].ToString()));
                        comboBox23.Items.Add((rdr[2].ToString() + ";" + rdr[1].ToString()));
                    }

                    /*таблица 15 Момент настання смерті*/
                    cmd_db = new SQLiteCommand("SELECT * from moment_smerti", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox25.Items.Clear();
                    comboBox26.Items.Clear();
                    comboBox27.Items.Clear();
                    comboBox28.Items.Clear();
                    comboBox29.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox25.Items.Add((rdr[1].ToString()));
                        comboBox26.Items.Add((rdr[1].ToString()));
                        comboBox27.Items.Add((rdr[1].ToString()));
                        comboBox28.Items.Add((rdr[1].ToString()));
                        comboBox29.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 16 Умови, що вплинули на загибель людей*/
                    cmd_db = new SQLiteCommand("SELECT * from umova_smerti", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox30.Items.Clear();
                    comboBox31.Items.Clear();
                    comboBox32.Items.Clear();
                    comboBox33.Items.Clear();
                    comboBox34.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox30.Items.Add((rdr[1].ToString()));
                        comboBox31.Items.Add((rdr[1].ToString()));
                        comboBox32.Items.Add((rdr[1].ToString()));
                        comboBox33.Items.Add((rdr[1].ToString()));
                        comboBox34.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 17 Інформація про ліквідацію пожежі*/
                    cmd_db = new SQLiteCommand("SELECT * from info_fire", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox35.Items.Clear();
                   
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox35.Items.Add((rdr[1].ToString()));
                       
                    }

                    /*таблица 18 Умови, що вплинули на поширення пожежі*/
                    cmd_db = new SQLiteCommand("SELECT * from poshirenya_fire", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox36.Items.Clear();
                    comboBox37.Items.Clear();
                    comboBox38.Items.Clear();
                    comboBox39.Items.Clear();
                    comboBox40.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox36.Items.Add((rdr[1].ToString()));
                        comboBox37.Items.Add((rdr[1].ToString()));
                        comboBox38.Items.Add((rdr[1].ToString()));
                        comboBox39.Items.Add((rdr[1].ToString()));
                        comboBox40.Items.Add((rdr[1].ToString()));

                    }

                    /*таблица 19 Умови, що ускладнювали гасіння пожежі*/
                    cmd_db = new SQLiteCommand("SELECT * from uskladnenya_fire", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox41.Items.Clear();
                    comboBox42.Items.Clear();
                    comboBox43.Items.Clear();
                    comboBox44.Items.Clear();
                    comboBox45.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox41.Items.Add((rdr[1].ToString()));
                        comboBox42.Items.Add((rdr[1].ToString()));
                        comboBox43.Items.Add((rdr[1].ToString()));
                        comboBox44.Items.Add((rdr[1].ToString()));
                        comboBox45.Items.Add((rdr[1].ToString()));

                    }

                    /*таблица 20 Наявність систем протипожежного захисту*/
                    cmd_db = new SQLiteCommand("SELECT * from nayavnist_spz", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox46.Items.Clear();
                   
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox46.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 21 Системи протипожежного захисту*/
                    cmd_db = new SQLiteCommand("SELECT * from system_protipojeji", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox47.Items.Clear();
                    comboBox48.Items.Clear();
                    comboBox49.Items.Clear();
                    comboBox50.Items.Clear();
                    comboBox51.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox47.Items.Add((rdr[1].ToString()));
                        comboBox48.Items.Add((rdr[1].ToString()));
                        comboBox49.Items.Add((rdr[1].ToString()));
                        comboBox50.Items.Add((rdr[1].ToString()));
                        comboBox51.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 22 Результат дії системи протипожежного захисту*/
                    cmd_db = new SQLiteCommand("SELECT * from resultat_dii_system", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox52.Items.Clear();
                    comboBox53.Items.Clear();
                    comboBox54.Items.Clear();
                    comboBox55.Items.Clear();
                    comboBox56.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox52.Items.Add((rdr[1].ToString()));
                        comboBox53.Items.Add((rdr[1].ToString()));
                        comboBox54.Items.Add((rdr[1].ToString()));
                        comboBox55.Items.Add((rdr[1].ToString()));
                        comboBox56.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 23 Перелік кодів учасників гасіння пожежі*/
                    cmd_db = new SQLiteCommand("SELECT * from uchasnik_fire", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox57.Items.Clear();
                    comboBox58.Items.Clear();
                    comboBox59.Items.Clear();
                    comboBox60.Items.Clear();
                    comboBox61.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox57.Items.Add((rdr[1].ToString()));
                        comboBox58.Items.Add((rdr[1].ToString()));
                        comboBox59.Items.Add((rdr[1].ToString()));
                        comboBox60.Items.Add((rdr[1].ToString()));
                        comboBox61.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 24 Перелік кодів пожежних автомобілів*/
                    cmd_db = new SQLiteCommand("SELECT * from fire_auto", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox62.Items.Clear();
                    comboBox63.Items.Clear();
                    comboBox64.Items.Clear();
                    comboBox65.Items.Clear();
                    comboBox66.Items.Clear();
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox62.Items.Add((rdr[1].ToString()));
                        comboBox63.Items.Add((rdr[1].ToString()));
                        comboBox64.Items.Add((rdr[1].ToString()));
                        comboBox65.Items.Add((rdr[1].ToString()));
                        comboBox66.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 25. Перелік кодів пожежних стволів*/
                    cmd_db = new SQLiteCommand("SELECT * from fire_stvoli", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox67.Items.Clear();
                    comboBox68.Items.Clear();
                    comboBox69.Items.Clear();
                   
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox67.Items.Add((rdr[1].ToString()));
                        comboBox68.Items.Add((rdr[1].ToString()));
                        comboBox69.Items.Add((rdr[1].ToString()));
              
                    }

                    /*таблица 26. Перелік кодів вогнегасних речовин*/
                    cmd_db = new SQLiteCommand("SELECT * from vognegasni_rechovini", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox70.Items.Clear();
                    comboBox71.Items.Clear();
                    comboBox72.Items.Clear();
                    
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox70.Items.Add((rdr[1].ToString()));
                        comboBox71.Items.Add((rdr[1].ToString()));
                        comboBox72.Items.Add((rdr[1].ToString()));

                    }

                    /*таблица 27. Перелік кодів первинних засобів пожежогасіння*/
                    cmd_db = new SQLiteCommand("SELECT * from zacobi_fire", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox73.Items.Clear();
                    comboBox74.Items.Clear();
                    comboBox75.Items.Clear();
                    
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox73.Items.Add((rdr[1].ToString()));
                        comboBox74.Items.Add((rdr[1].ToString()));
                        comboBox75.Items.Add((rdr[1].ToString()));

                    }

                    /*таблица 28. Перелік кодів джерел водопостачання*/
                    cmd_db = new SQLiteCommand("SELECT * from djerela_vodopostachanaya", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox76.Items.Clear();
                    comboBox77.Items.Clear();
                    comboBox78.Items.Clear();
                   
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox76.Items.Add((rdr[1].ToString()));
                        comboBox77.Items.Add((rdr[1].ToString()));
                        comboBox78.Items.Add((rdr[1].ToString()));

                    }

                    /*таблица 29. Перелік кодів видів перевірки об’єкта пожежі*/
                    cmd_db = new SQLiteCommand("SELECT * from perevirka_fire", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox79.Items.Clear();
                   
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox79.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 30. Перелік кодів умов, на підставі яких здійснюється діяльність об’єкта*/
                    cmd_db = new SQLiteCommand("SELECT * from gospodar_dialnist", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox80.Items.Clear();
                   
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        comboBox80.Items.Add((rdr[1].ToString()));
                    }

                    /*таблица 31. Перелік кодів вжитих заходів*/
                    cmd_db = new SQLiteCommand("SELECT * from zahodi_pojeji", con_db);
                    rdr = cmd_db.ExecuteReader();
                    comboBox81.Items.Clear();
                    comboBox82.Items.Clear();
                   
                    while (rdr.Read())
                    {
                        // region_items.Add(rdr[1].ToString());
                        if (int.Parse(rdr[2].ToString()) < 4)
                        {
                            comboBox81.Items.Add((rdr[1].ToString()));
                        }
                        else
                        {
                            comboBox82.Items.Add((rdr[1].ToString()));
                        }
                      
                        
                    }

                  
                }
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
        
       //  Disconnect();
}


  public void connection(){

		//try {

			string path=Environment.CurrentDirectory+"/MyDataBase/info.bytes.db";
             
               // if (con_db.State == ConnectionState.Closed)
               // {
                    con_db = new SQLiteConnection(string.Format("DATA Source={0};", path));
                    con_db.Open();
              //  }
			

		//} catch (Exception ex) {
              //  Console.WriteLine(ex.Message);
		//}
        // path = File.ReadAllLines(Application.dataPath + "/StreamingAssets/new1.xml")[0];
        //Debug.Log(path);
    
      

       // Path.ChangeExtension(Application.dataPath + "/StreamingAssets/new1.xml", "xls");

    }
  public void Disconnect()
    {
        con_db.Close();
    }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox72.Text = "";
            textBox73.Text = "";
            //    connection();
            /*таблица 3 об'єкт пожежі*/
            string[] current_fire_item = comboBox3.SelectedItem.ToString().Split(')');
            cmd_db = new SQLiteCommand("SELECT * from fire_objects WHERE `fire_name`=" + "\'" + current_fire_item[1] + "\'", con_db);
            rdr = cmd_db.ExecuteReader();
            comboBox4.Items.Clear();
            while (rdr.Read())
            {
                // region_items.Add(rdr[1].ToString());
                comboBox4.Items.Add((rdr[3].ToString()+";"+rdr[2].ToString()));
            }
            // comboBox4.SelectedIndex = 0;
           
         //   Disconnect();
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox80.Text = "";
            textBox157.Text = "";
            //connection();
            /*таблица 11 Виріб-ініціатор пожежі*/
            try
            {
                string[] current_fire_item = comboBox12.SelectedItem.ToString().Split(')');
                cmd_db = new SQLiteCommand("SELECT * from virib_iniciator WHERE `name_virib`=" + "\'" + current_fire_item[1] + "\'", con_db);
                rdr = cmd_db.ExecuteReader();
                comboBox13.Items.Clear();
                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    comboBox13.Items.Add((rdr[3].ToString() + ";" + rdr[2].ToString()));
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
            
           // comboBox13.SelectedIndex = 0;
          //  Disconnect();
        }

        private void comboBox46_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox113.Text = CodeAllContent("nayavnist_spz", "name_spz", "code_spz", comboBox46.SelectedItem);
            if (comboBox46.SelectedIndex == 0)
            {
                panel1.Visible = true;
                //panel2.Visible = true;
            }
            else
            {
                panel1.Visible = false;
                //panel2.Visible = false;
            }
        }

       

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button1.Visible = true;
            panel3.Visible = false;
            panel4.Visible = false;
            panel2.Visible = false;
            panel5.Visible = false;
            foreach (var item in tb)
            {
                if (item.Name != "textBox69")
                {
                    item.Clear();
                    item.Enabled = true;

                }
                else
                {
                    item.Select();
                }
                if (item.Name == "textBox3")
                {
                    item.Text = "0";
                }
               
                
            }
            foreach (var item in tb_date)
            {
                if (item.Name != "dateTimePicker9")
                {
                    item.Enabled = true;
                }
               
                if(item.Name== "dateTimePicker10")
                {
                    item.Value = DateTime.Today;
                }
                else
                {
                    item.Value = DateTime.Parse("01.01.1900");
                }
                
            }
            foreach (var item in tb_masked)
            {
                if (item.Name != maskedTextBox6.Name)
                {
                    item.Text = "";
                } 
            }
            foreach (var item in tb_combo)
            {
                //item.SelectionLength = 0;
                item.Text = "";


            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox1);
            textBox1.BackColor = Color.White;
            textBox1.ForeColor = Color.Black;

        }
        private void validateIntForm(TextBox textBox)
        {
            int x;
            try
            {
                if (textBox.Text != "")
                    x = int.Parse(textBox.Text);
            }
            catch (Exception)
            {

               // MessageBox.Show("В це поле можно ввести тільки цілі числа!");
                textBox.Text = "";
            }
        }
        private void validateFloatForm(TextBox textBox)
        {
            float x;
            try
            {
                if (textBox.Text != "")
                    x = float.Parse(textBox.Text);
            }
            catch (Exception)
            {

               // MessageBox.Show("В це поле можно ввести тільки цифри!");
                textBox.Text = "";
            }
           
        }


        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox2);
            textBox2.BackColor = Color.White;
            textBox2.ForeColor = Color.Black;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
           // validateIntForm(textBox3);
           
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox5);
            textBox5.BackColor = Color.White;
            textBox5.ForeColor = Color.Black;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox6);
            textBox6.BackColor = Color.White;
            textBox6.ForeColor = Color.Black;
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox8);
            textBox8.BackColor = Color.White;
            textBox8.ForeColor = Color.Black;
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox65);
            textBox65.BackColor = Color.White;
            textBox65.ForeColor = Color.Black;
        }

        private void textBox65_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox65);
            textBox65.BackColor = Color.White;
            textBox65.ForeColor = Color.Black;
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox12);
            textBox12.BackColor = Color.White;
            textBox12.ForeColor = Color.Black;
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox11);
            textBox11.BackColor = Color.White;
            textBox11.ForeColor = Color.Black;
        }

        private void textBox66_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox66);
            textBox66.BackColor = Color.White;
            textBox66.ForeColor = Color.Black;
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox9);
            textBox9.BackColor = Color.White;
            textBox9.ForeColor = Color.Black;
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox15);
            textBox15.BackColor = Color.White;
            textBox15.ForeColor = Color.Black;
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox7);
            textBox7.BackColor = Color.White;
            textBox7.ForeColor = Color.Black;
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox14);
            textBox14.BackColor = Color.White;
            textBox14.ForeColor = Color.Black;
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox13);
            textBox13.BackColor = Color.White;
            textBox13.ForeColor = Color.Black;
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox16);
            textBox16.BackColor = Color.White;
            textBox16.ForeColor = Color.Black;
        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox17);
            textBox17.BackColor = Color.White;
            textBox17.ForeColor = Color.Black;
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox17.Text = "";
            }
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox19);
            textBox19.BackColor = Color.White;
            textBox19.ForeColor = Color.Black;
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox19.Text = "";
            }
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox21);
            textBox21.BackColor = Color.White;
            textBox21.ForeColor = Color.Black;
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox21.Text = "";
            }
        }

        private void textBox23_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox23);
            textBox23.BackColor = Color.White;
            textBox23.ForeColor = Color.Black;
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox23.Text = "";
            }
        }

        private void textBox25_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox25);
            textBox25.BackColor = Color.White;
            textBox25.ForeColor = Color.Black;
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox25.Text = "";
            }
        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox28);
            textBox28.BackColor = Color.White;
            textBox28.ForeColor = Color.Black;
        }

        private void textBox27_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox27);
            textBox27.BackColor = Color.White;
            textBox27.ForeColor = Color.Black;
        }

        private void textBox36_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox36);
            textBox36.BackColor = Color.White;
            textBox36.ForeColor = Color.Black;
        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox30);
            textBox30.BackColor = Color.White;
            textBox30.ForeColor = Color.Black;
        }

        private void textBox29_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox29);
            textBox29.BackColor = Color.White;
            textBox29.ForeColor = Color.Black;
        }

        private void textBox35_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox35);
            textBox35.BackColor = Color.White;
            textBox35.ForeColor = Color.Black;
        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox32);
            textBox32.BackColor = Color.White;
            textBox32.ForeColor = Color.Black;
        }

        private void textBox31_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox31);
            textBox31.BackColor = Color.White;
            textBox31.ForeColor = Color.Black;
        }

        private void textBox49_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox49);
            textBox49.BackColor = Color.White;
            textBox49.ForeColor = Color.Black;
        }

        private void textBox50_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox50);
            textBox50.BackColor = Color.White;
            textBox50.ForeColor = Color.Black;
        }

        private void textBox51_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox51);
            textBox51.BackColor = Color.White;
            textBox51.ForeColor = Color.Black;
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox52);
            textBox52.BackColor = Color.White;
            textBox52.ForeColor = Color.Black;
        }

        private void textBox53_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox53);
            textBox53.BackColor = Color.White;
            textBox53.ForeColor = Color.Black;
        }

        private void textBox58_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox58);
            textBox58.BackColor = Color.White;
            textBox58.ForeColor = Color.Black;
        }

        private void textBox57_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox57);
            textBox57.BackColor = Color.White;
            textBox57.ForeColor = Color.Black;
        }

        private void textBox56_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox56);
            textBox56.BackColor = Color.White;
            textBox56.ForeColor = Color.Black;
        }

        private void textBox55_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox55);
            textBox55.BackColor = Color.White;
            textBox55.ForeColor = Color.Black;
        }

        private void textBox54_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox54);
            textBox54.BackColor = Color.White;
            textBox54.ForeColor = Color.Black;
        }

        private void textBox62_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox62);
            textBox62.BackColor = Color.White;
            textBox62.ForeColor = Color.Black;
            if (textBox62.Text == "" && textBox63.Text == "" && textBox64.Text == "")
            {
                textBox62.BackColor = Color.White;
                textBox63.BackColor = Color.White;
                textBox64.BackColor = Color.White;
                textBox62.ForeColor = Color.Black;
                textBox63.ForeColor = Color.Black;
                textBox64.ForeColor = Color.Black;
            }
        }

        private void textBox63_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox63);
            textBox63.BackColor = Color.White;
            textBox63.ForeColor = Color.Black;
            if (textBox62.Text == "" && textBox63.Text == "" && textBox64.Text == "")
            {
                textBox62.BackColor = Color.White;
                textBox63.BackColor = Color.White;
                textBox64.BackColor = Color.White;
                textBox62.ForeColor = Color.Black;
                textBox63.ForeColor = Color.Black;
                textBox64.ForeColor = Color.Black;
            }
        }

        private void textBox64_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox64);
            textBox64.BackColor = Color.White;
            textBox64.ForeColor = Color.Black;
            if (textBox62.Text == "" && textBox63.Text == "" && textBox64.Text == "")
            {
                textBox62.BackColor = Color.White;
                textBox63.BackColor = Color.White;
                textBox64.BackColor = Color.White;
                textBox62.ForeColor = Color.Black;
                textBox63.ForeColor = Color.Black;
                textBox64.ForeColor = Color.Black;
            }
            if (textBox64.Text == "0")
            {
                MessageBox.Show("В це поле неможливо записати 0");
                textBox64.Text = "";
            }

        }

        private void textBox37_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox37);
            textBox37.BackColor = Color.White;
            textBox37.ForeColor = Color.Black;
        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox34);
            textBox34.BackColor = Color.White;
            textBox34.ForeColor = Color.Black;
        }

        private void textBox39_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox39);
            textBox39.BackColor = Color.White;
            textBox39.ForeColor = Color.Black;
        }

        private void textBox38_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox38);
            textBox38.BackColor = Color.White;
            textBox38.ForeColor = Color.Black;
        }

        private void textBox41_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox41);
            textBox41.BackColor = Color.White;
            textBox41.ForeColor = Color.Black;
        }

        private void textBox40_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox40);
            textBox40.BackColor = Color.White;
            textBox40.ForeColor = Color.Black;
        }

        private void textBox44_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox44);
            textBox44.BackColor = Color.White;
            textBox44.ForeColor = Color.Black;
        }

        private void textBox46_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox46);
            textBox46.BackColor = Color.White;
            textBox46.ForeColor = Color.Black;
        }

        private void textBox45_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox45);
            textBox45.BackColor = Color.White;
            textBox45.ForeColor = Color.Black;
        }

        private void textBox43_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox43);
            textBox43.BackColor = Color.White;
            textBox43.ForeColor = Color.Black;
        }

        private void textBox42_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox42);
            textBox42.BackColor = Color.White;
            textBox42.ForeColor = Color.Black;
        }

        private void textBox48_TextChanged(object sender, EventArgs e)
        {
            validateFloatForm(textBox48);
            textBox48.BackColor = Color.White;
            textBox48.ForeColor = Color.Black;
        }

        private bool validateField(/*ComboBox[] combo,*/ TextBox[] textbox)
        {
            bool val=true;
            int result;
            float result1;
            string content_val = "";
            /*for (int i = 0; i < combo.Length; i++)
            {
                if (combo[i].SelectedItem == null)
                {
                    combo[i].BackColor = Color.Red;
                    val = false;
                }
            }*/
            if (textBox9.Text == "")
            {
                textBox9.BackColor = Color.Firebrick;
                textBox9.ForeColor = Color.White;
                val = false;
            }
            if (textBox3.Text=="0" && textBox102.Text == "")
            {
                val = false;
                textBox102.BackColor = Color.Firebrick;
            }
            if(textBox72.Text!="" && textBox73.Text == "")
            {
                val = false;
                textBox73.BackColor = Color.Firebrick;
            }
            if (textBox79.Text != "" && textBox156.Text == "")
            {
                val = false;
                textBox156.BackColor = Color.Firebrick;
            }
            if (textBox80.Text != "" && textBox157.Text == "")
            {
                val = false;
                textBox157.BackColor = Color.Firebrick;
            }
            if (textBox81.Text != "" && textBox158.Text == "")
            {
                val = false;
                textBox158.BackColor = Color.Firebrick;
            }

            if (textBox92.Text == "1" || textBox92.Text == "2")
            {
                if (textBox5.Text == "")
                {
                    val = false;
                    textBox5.BackColor = Color.Firebrick;

                }
                    
            }
            if (textBox113.Text == "1")
            {
                if (textBox118.Text=="")
                {
                    val = false;
                    textBox118.BackColor = Color.Firebrick;
                }
                if(textBox123.Text == "")
                {
                    val = false;
                    textBox123.BackColor = Color.Firebrick;
                }
            }
            else
            {
                textBox118.Text = "";
                textBox123.Text = "";
                textBox117.Text = "";
                textBox122.Text = "";
                textBox116.Text = "";
                textBox121.Text = "";
                textBox115.Text = "";
                textBox120.Text = "";
                textBox114.Text = "";
                textBox119.Text = "";
            }
            
            if(dateTimePicker9.Value.ToString()!= "01.01.1900 0:00:00")
            {
                if (textBox146.Text == "")
                {
                    val = false;
                    textBox146.BackColor = Color.Firebrick;
                }
            }
            if (textBox147.Text != "")
            {
                if (textBox76.Text == "" || textBox76.Text=="20")
                {
                    val = false;
                    textBox76.BackColor = Color.Firebrick;
                    textBox76.Text = "";
                }
            }
            bool val50 = int.TryParse(textBox102.Text, out int res50);
            if (res50 == 1)
            {
                if (textBox128.Text == "")
                {
                    val = false;
                    textBox128.BackColor = Color.Firebrick;
                }
                if (textBox129.Text == "")
                {
                    val = false;
                    textBox129.BackColor = Color.Firebrick;
                }
                if (textBox58.Text == "")
                {
                    val = false;
                    textBox58.BackColor = Color.Firebrick;
                }
            }
            if (res50 == 2)
            {
                if (textBox128.Text == "")
                {
                    val = false;
                    textBox128.BackColor = Color.Firebrick;
                }
                if (textBox49.Text == "")
                {
                    val = false;
                    textBox49.BackColor = Color.Firebrick;
                }

                if (textBox129.Text == "")
                {
                    val = false;
                    textBox129.BackColor = Color.Firebrick;
                }
                if (textBox58.Text == "")
                {
                    val = false;
                    textBox58.BackColor = Color.Firebrick;
                }
                if (textBox137.Text == "")
                {
                    val = false;
                    textBox137.BackColor = Color.Firebrick;
                }
            }
           
            
            bool val46 = float.TryParse(textBox48.Text, out float res46);
            if (res46 > 0 && (textBox37.Text == "" && textBox39.Text == "" && textBox38.Text == "" && textBox41.Text == "" && textBox40.Text == "" && textBox44.Text == "" && textBox46.Text == "" && textBox43.Text == "" && textBox42.Text == "" && textBox47.Text == ""))
            {
                val = false;
                textBox48.BackColor = Color.Firebrick;
            }
            bool val25 = float.TryParse(textBox9.Text, out float res25);
            if (res25 > 0 &&(textBox7.Text=="" && textBox14.Text=="" && textBox13.Text == "" && textBox16.Text == "" && textBox28.Text == "" && textBox27.Text == "" && textBox36.Text == "" && textBox30.Text == "" && textBox29.Text == "" && textBox35.Text == "" && textBox32.Text == "" && textBox31.Text == "" && textBox33.Text == ""))
            {
                val = false;
                textBox9.BackColor = Color.Firebrick;
            }
            bool val12 = int.TryParse(textBox12.Text, out int res12);
            bool val11 = int.TryParse(textBox11.Text, out int res11);
            bool val66 = int.TryParse(textBox66.Text, out int res66);
            if (textBox12.Text != "")
            {
                if (res12 < (res11 + res66))
                {
                    val = false;
                    if (res11>=res12)
                    {
                        textBox11.BackColor = Color.Firebrick;
                        textBox11.ForeColor = Color.White;
                    }
                    if (res66>=res12)
                    {
                        textBox66.BackColor = Color.Firebrick;
                        textBox66.ForeColor = Color.White;
                    }
                }
            }
            if (textBox12.Text == "" && (textBox11.Text != "" || textBox66.Text != ""))
            {
                val = false;
                textBox12.BackColor = Color.Firebrick;
            }

            bool val18 = int.TryParse(textBox8.Text, out int res18);
            bool val65 = int.TryParse(textBox65.Text, out int res65);
            bool val82 = int.TryParse(textBox82.Text, out int res82);
            if (textBox8.Text != "")
            {
                if (res18 < (res65 + res82))
                {
                    val = false;
                    if (res65 >= res18)
                    {
                        textBox65.BackColor = Color.Firebrick;
                        textBox65.ForeColor = Color.White;
                    }
                    if (res82 >= res18)
                    {
                        textBox82.BackColor = Color.Firebrick;
                        textBox82.ForeColor = Color.White;
                    }
                }
            }
          
            if(textBox8.Text=="" && (textBox65.Text != "" || textBox82.Text != ""))
            {
                val = false;
                textBox8.BackColor = Color.Firebrick;
            }
            if(textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox10.BackColor = Color.White;
                textBox87.BackColor = Color.White;
                textBox92.BackColor = Color.White;
                textBox97.BackColor = Color.White;

                textBox18.Text = "";
                textBox17.Text = "";
                textBox10.Text = "";
                textBox87.Text = "";
                textBox92.Text = "";
                textBox97.Text = "";

                textBox19.Text = "";
                textBox20.Text = "";  
                textBox83.Text = "";
                textBox88.Text = "";
                textBox93.Text = "";
                textBox98.Text = "";

                textBox21.Text = "";
                textBox22.Text = "";
                textBox84.Text = "";
                textBox89.Text = "";
                textBox94.Text = "";
                textBox99.Text = "";

                textBox23.Text  = "";
                textBox24.Text  = "";
                textBox85.Text  = "";
                textBox90.Text  = "";
                textBox95.Text  = "";
                textBox100.Text = "";

                textBox25.Text  = "";
                textBox26.Text  = "";
                textBox86.Text  = "";
                textBox91.Text  = "";
                textBox96.Text  = "";
                textBox101.Text = "";

            }
          
            if (textBox8.Text!="")
            {
                
                if (textBox10.Text == "")
                {
                    val = false;
                    textBox10.BackColor = Color.Firebrick;
                }
                else
                {
                    textBox10.BackColor = Color.White;
                }
                if (textBox87.Text == "")
                {
                    val = false;
                    textBox87.BackColor = Color.Firebrick;
                }
                else
                {
                    textBox87.BackColor = Color.White;
                }
                if (textBox92.Text == "")
                {
                    val = false;
                    textBox92.BackColor = Color.Firebrick;
                }
                else
                {
                    textBox92.BackColor = Color.White;
                }
                if (textBox97.Text == "")
                {
                    val = false;
                    textBox97.BackColor = Color.Firebrick;
                }
                else
                {
                    textBox97.BackColor = Color.White;
                }
               
            }
            
            bool val_13 = int.TryParse(textBox72.Text, out int res13);
                if(res13>=1 && res13<210)
                {
              
                if (textBox78.Text == "")
                    {
                    val = false;
                    textBox78.BackColor = Color.Firebrick;
                    }
                }          
            bool val_cur3 = int.TryParse(textBox67.Text, out int res3);
            bool val_cur2 = int.TryParse(textBox68.Text, out int res1);
     
            if (res3 < res1)
            {
                if(res1 != 201 && res1 != 202 && res1 != 203 && res1 != 204)
                {
                    if (textBox67.Text != "")
                    {
                        val = false;
                        textBox67.BackColor = Color.Aqua;
                       // textBox67.ForeColor = Color.White;
                    }
                    if (textBox68.Text != "")
                    {
                        val = false;
                        textBox68.BackColor = Color.Aqua;
                        //textBox68.ForeColor = Color.White;
                    }
                }
            }
            else
            {
                textBox67.BackColor = Color.White;
                textBox67.ForeColor = Color.Black;
               // textBox68.BackColor = Color.White;
               // textBox68.ForeColor = Color.Black;
            }


            bool val_cur10 = int.TryParse(textBox76.Text, out int res10);
            if(res10==11 || res10==12 || res10 == 13)
            {
                if (textBox146.Text == "")
                {
                    val = false;
                    textBox146.BackColor = Color.Firebrick;

                }
            }
            
            if (textBox75.Text != "")
            {
                if (textBox76.Text == "")
                {
                    val = false;
                    textBox76.BackColor = Color.Firebrick;
                }
            }
            else
            {
               // textBox76.BackColor = Color.White;
            }

            bool val_cur = int.TryParse(textBox72.Text,out int res);
            if ((res >= 1601 && res < 1606) || (res >= 1701 && res < 1708) || (res >= 1801 && res < 1819))
            {
                if (textBox67.Text!= "")
                {
                    //MessageBox.Show("Це поле повинне бути пустим!");
                    textBox67.BackColor = Color.Firebrick;
                    textBox67.ForeColor = Color.White;
                    val = false;
                }
                if(textBox68.Text!= "")
                {
                    textBox68.BackColor = Color.Firebrick;
                    textBox68.ForeColor = Color.White;
                    val = false;
                }
                if(textBox77.Text!= "")
                {
                    textBox77.BackColor = Color.Firebrick;
                    textBox77.ForeColor = Color.White;
                    val = false;
                }
                if(textBox78.Text!= "")
                {
                    textBox78.BackColor = Color.Firebrick;
                    textBox78.ForeColor = Color.White;
                    val = false;
                }
            }
           

            string select_code = "SELECT `number_cartka` FROM kartka_obliku WHERE code_raion=" + textBox1.Text + " AND number_cartka=" + textBox2.Text + " AND main_dop=" + textBox3.Text;
            //  connection();
            try
            {
              //  if (con_db.State == ConnectionState.Open)
              //  {
                    cmd_db = new SQLiteCommand(select_code, con_db);
                    rdr = cmd_db.ExecuteReader();

                    while (rdr.Read())
                    {
                        content_val = rdr[0].ToString();
                        // dict.Add(code, rdr[0].ToString());

                    }
               // }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
            try
            {
                if (content_val != "" && flag_edit == false)
                {
                    textBox2.BackColor = Color.Firebrick;
                    textBox2.ForeColor = Color.White;
                    textBox3.BackColor = Color.Firebrick;
                    textBox3.ForeColor = Color.White;
                    val = false;
                }
                if (textBox102.Text == "")
                {
                    dict.Add("code_fire_likvid", "");
                    dict.Add("name_fire_likvid", "");
                }
                for (int i = 0; i < textbox.Length; i++)
                {
                    if (textbox[i].Text == "" /*|| int.TryParse(textbox[i].Text, out result) == false*/)
                    {
                        if (textbox[i].Name == "textBox102")
                        {
                            if (textBox3.Text == "0")
                            {
                                textbox[i].BackColor = Color.Firebrick;
                                textbox[i].ForeColor = Color.White;

                                val = false;
                            }

                        }
                        else
                        {
                            textbox[i].BackColor = Color.Firebrick;
                            textbox[i].ForeColor = Color.White;
                            val = false;
                        }

                    }

                }
                if (textBox67.Text != "")
                {
                    if (textBox68.Text == "")
                    {
                        if (comboBox8.SelectedItem == null)
                        {
                            comboBox8.BackColor = Color.Firebrick;
                            val = false;
                        }
                    }
                    else
                    {
                        comboBox8.BackColor = Color.White;
                    }
                }
              
                if (textBox62.Text != "" || textBox63.Text != "" || textBox64.Text != "")
                {
                    if (textBox62.Text == "")
                    {
                        textBox62.BackColor = Color.Red;
                        val = false;
                    }
                    if (textBox63.Text == "")
                    {
                        textBox63.BackColor = Color.Red;
                        val = false;
                    }
                    if (textBox64.Text == "")
                    {
                        textBox64.BackColor = Color.Red;
                        val = false;
                    }
                }

                if (textBox128.Text != "")
                {
                    if (textBox49.Text == "")
                    {
                        val = false;
                        textBox49.BackColor = Color.Firebrick;
                    }
                }
                else if (textBox102.Text == "3")
                {
                    if (textBox49.Text != "")
                    {
                        val = false;
                        textBox128.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox128.BackColor = Color.White;
                        textBox49.BackColor = Color.White;
                    }
                }
                if (textBox127.Text != "")
                {
                    if (textBox50.Text == "")
                    {
                        val = false;
                        textBox50.BackColor = Color.Firebrick;
                    }
                }
                else
                {
                    if (textBox50.Text != "")
                    {
                        val = false;
                        textBox127.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox127.BackColor = Color.White;
                        textBox50.BackColor = Color.White;
                    }
                }
                if (textBox126.Text != "")
                {
                    if (textBox51.Text == "")
                    {
                        val = false;
                        textBox51.BackColor = Color.Firebrick;
                    }

                }
                else
                {
                    if (textBox51.Text != "")
                    {
                        val = false;
                        textBox126.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox126.BackColor = Color.White;
                        textBox51.BackColor = Color.White;
                    }
                }
                if (textBox125.Text != "")
                {
                    if (textBox52.Text == "")
                    {
                        val = false;
                        textBox52.BackColor = Color.Firebrick;
                    }
                }
                else
                {
                    if (textBox52.Text != "")
                    {
                        val = false;
                        textBox125.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox125.BackColor = Color.White;
                        textBox52.BackColor = Color.White;
                    }
                }
                if (textBox124.Text != "")
                {
                    if (textBox53.Text == "")
                    {
                        val = false;
                        textBox53.BackColor = Color.Firebrick;
                    }
                }
                else
                {
                    if (textBox53.Text != "")
                    {
                        val = false;
                        textBox124.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox124.BackColor = Color.White;
                        textBox53.BackColor = Color.White;
                    }
                }


                /*8888888888888888888888888888888888888888888888888888888*/

                if (textBox129.Text != "")
                {
                    if (textBox58.Text == "")
                    {
                        val = false;
                        textBox58.BackColor = Color.Firebrick;
                    }
                }
                else if (textBox102.Text == "3")
                {
                    if (textBox58.Text != "")
                    {
                        val = false;
                        textBox129.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox129.BackColor = Color.White;
                        textBox58.BackColor = Color.White;
                    }
                }
                if (textBox130.Text != "")
                {
                    if (textBox57.Text == "")
                    {
                        val = false;
                        textBox57.BackColor = Color.Firebrick;
                    }
                }
                else
                {
                    if (textBox57.Text != "")
                    {
                        val = false;
                        textBox130.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox130.BackColor = Color.White;
                        textBox57.BackColor = Color.White;
                    }
                }
                if (textBox131.Text != "")
                {
                    if (textBox56.Text == "")
                    {
                        val = false;
                        textBox56.BackColor = Color.Firebrick;
                    }

                }
                else
                {
                    if (textBox56.Text != "")
                    {
                        val = false;
                        textBox131.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox131.BackColor = Color.White;
                        textBox56.BackColor = Color.White;
                    }
                }
                if (textBox132.Text != "")
                {
                    if (textBox55.Text == "")
                    {
                        val = false;
                        textBox55.BackColor = Color.Firebrick;
                    }
                }
                else
                {
                    if (textBox55.Text != "")
                    {
                        val = false;
                        textBox132.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox132.BackColor = Color.White;
                        textBox55.BackColor = Color.White;
                    }
                }
                if (textBox133.Text != "")
                {
                    if (textBox54.Text == "")
                    {
                        val = false;
                        textBox54.BackColor = Color.Firebrick;
                    }
                }
                else
                {
                    if (textBox54.Text != "")
                    {
                        val = false;
                        textBox133.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox133.BackColor = Color.White;
                        textBox54.BackColor = Color.White;
                    }
                }


                /*3333333333333333333333333333*/
                if (textBox134.Text != "")
                {
                    if (textBox61.Text == "")
                    {
                        val = false;
                        textBox61.BackColor = Color.Firebrick;
                    }

                }
                else
                {
                    if (textBox61.Text != "")
                    {
                        val = false;
                        textBox134.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox134.BackColor = Color.White;
                        textBox61.BackColor = Color.White;
                    }
                }

                if (textBox135.Text != "")
                {
                    if (textBox60.Text == "")
                    {
                        val = false;
                        textBox60.BackColor = Color.Firebrick;
                    }
                }
                else
                {
                    if (textBox60.Text != "")
                    {
                        val = false;
                        textBox135.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox135.BackColor = Color.White;
                        textBox60.BackColor = Color.White;
                    }
                }
                if (textBox136.Text != "")
                {
                    if (textBox59.Text == "")
                    {
                        val = false;
                        textBox59.BackColor = Color.Firebrick;
                    }
                }
                else
                {
                    if (textBox59.Text != "")
                    {
                        val = false;
                        textBox136.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        textBox136.BackColor = Color.White;
                        textBox59.BackColor = Color.White;
                    }
                }

                /**4555555555555555555555555555555*/


                if (textBox134.Text != "")
                {
                    if (textBox137.Text == "")
                    {
                        val = false;
                        textBox137.BackColor = Color.Firebrick;
                    }

                }
                else if (textBox102.Text == "3")
                {
                    if (textBox137.Text != "")
                    {
                        val = false;
                        textBox134.BackColor = Color.Firebrick;
                        //textBox61.BackColor = Color.Firebrick;
                        //textBox137.BackColor = Color.White;
                    }
                    textBox137.BackColor = Color.White;

                }


                if (textBox135.Text != "")
                {
                    if (textBox138.Text == "")
                    {
                        val = false;
                        textBox138.BackColor = Color.Firebrick;
                    }

                }
                else
                {
                    if (textBox138.Text != "")
                    {
                        val = false;
                        textBox135.BackColor = Color.Firebrick;
                    }

                    // textBox60.BackColor = Color.White;
                    textBox138.BackColor = Color.White;
                }


                if (textBox136.Text != "")
                {
                    if (textBox139.Text == "")
                    {
                        val = false;
                        textBox139.BackColor = Color.Firebrick;
                    }

                }
                else
                {
                    if (textBox139.Text != "")
                    {
                        val = false;
                        textBox136.BackColor = Color.Firebrick;
                    }
                    textBox139.BackColor = Color.White;
                }
                if (textBox67.Text != "")
                {
                    if (comboBox8.SelectedItem == null)
                    {
                        if (textBox68.Text == "")
                        {
                            textBox68.BackColor = Color.Firebrick;
                            textBox68.ForeColor = Color.White;
                            val = false;
                        }
                    }
                    else
                    {
                        textBox68.BackColor = Color.White;
                        textBox68.ForeColor = Color.Black;
                    }
                }

                if (textBox8.Text != "" || textBox65.Text != "" || textBox82.Text != "")
                {

                    if (textBox20.Text != "" || textBox83.Text != "" || textBox88.Text != "" || textBox93.Text != "" || textBox98.Text != "")
                    {
                        if (textBox20.Text == "")
                        {
                            textBox20.BackColor = Color.Firebrick;
                            textBox20.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox83.Text == "")
                        {
                            textBox83.BackColor = Color.Firebrick;
                            textBox83.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox88.Text == "")
                        {
                            textBox88.BackColor = Color.Firebrick;
                            textBox88.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox93.Text == "")
                        {
                            textBox93.BackColor = Color.Firebrick;
                            textBox93.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox98.Text == "")
                        {
                            textBox98.BackColor = Color.Firebrick;
                            textBox98.ForeColor = Color.White;
                            val = false;
                        }
                    }
                    else
                    {
                        if (textBox20.Text == "" && textBox83.Text == "" && textBox88.Text == "" && textBox93.Text == "" && textBox98.Text == "")
                        {
                            textBox20.BackColor = Color.White;
                            textBox20.ForeColor = Color.Black;
                            textBox83.BackColor = Color.White;
                            textBox83.ForeColor = Color.Black;
                            textBox88.BackColor = Color.White;
                            textBox88.ForeColor = Color.Black;
                            textBox93.BackColor = Color.White;
                            textBox93.ForeColor = Color.Black;
                            textBox98.BackColor = Color.White;
                            textBox98.ForeColor = Color.Black;

                        }
                    }
                    if (textBox22.Text != "" || textBox84.Text != "" || textBox89.Text != "" || textBox94.Text != "" || textBox99.Text != "")
                    {
                        if (textBox22.Text == "")
                        {
                            textBox22.BackColor = Color.Firebrick;
                            textBox22.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox84.Text == "")
                        {
                            textBox84.BackColor = Color.Firebrick;
                            textBox84.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox89.Text == "")
                        {
                            textBox89.BackColor = Color.Firebrick;
                            textBox89.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox94.Text == "")
                        {
                            textBox94.BackColor = Color.Firebrick;
                            textBox94.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox99.Text == "")
                        {
                            textBox99.BackColor = Color.Firebrick;
                            textBox99.ForeColor = Color.White;
                            val = false;
                        }
                    }
                    else
                    {
                        if (textBox22.Text == "" && textBox84.Text == "" && textBox89.Text == "" && textBox94.Text == "" && textBox99.Text == "")
                        {
                            textBox22.BackColor = Color.White;
                            textBox22.ForeColor = Color.Black;
                            textBox84.BackColor = Color.White;
                            textBox84.ForeColor = Color.Black;
                            textBox89.BackColor = Color.White;
                            textBox89.ForeColor = Color.Black;
                            textBox94.BackColor = Color.White;
                            textBox94.ForeColor = Color.Black;
                            textBox99.BackColor = Color.White;
                            textBox99.ForeColor = Color.Black;

                        }
                    }
                    if (textBox24.Text != "" || textBox85.Text != "" || textBox90.Text != "" || textBox95.Text != "" || textBox100.Text != "")
                    {
                        if (textBox24.Text == "")
                        {
                            textBox24.BackColor = Color.Firebrick;
                            textBox24.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox85.Text == "")
                        {
                            textBox85.BackColor = Color.Firebrick;
                            textBox85.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox90.Text == "")
                        {
                            textBox90.BackColor = Color.Firebrick;
                            textBox90.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox95.Text == "")
                        {
                            textBox95.BackColor = Color.Firebrick;
                            textBox95.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox100.Text == "")
                        {
                            textBox100.BackColor = Color.Firebrick;
                            textBox100.ForeColor = Color.White;
                            val = false;
                        }

                    }
                    else
                    {
                        if (textBox24.Text == "" && textBox85.Text == "" && textBox90.Text == "" && textBox95.Text == "" && textBox100.Text == "")
                        {
                            textBox24.BackColor = Color.White;
                            textBox24.ForeColor = Color.Black;
                            textBox85.BackColor = Color.White;
                            textBox85.ForeColor = Color.Black;
                            textBox90.BackColor = Color.White;
                            textBox90.ForeColor = Color.Black;
                            textBox95.BackColor = Color.White;
                            textBox95.ForeColor = Color.Black;
                            textBox100.BackColor = Color.White;
                            textBox100.ForeColor = Color.Black;

                        }
                    }
                    if (textBox26.Text != "" || textBox86.Text != "" || textBox91.Text != "" || textBox96.Text != "" || textBox101.Text != "")
                    {
                        if (textBox26.Text == "")
                        {
                            textBox26.BackColor = Color.Firebrick;
                            textBox26.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox86.Text == "")
                        {
                            textBox86.BackColor = Color.Firebrick;
                            textBox86.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox91.Text == "")
                        {
                            textBox91.BackColor = Color.Firebrick;
                            textBox91.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox96.Text == "")
                        {
                            textBox96.BackColor = Color.Firebrick;
                            textBox96.ForeColor = Color.White;
                            val = false;
                        }
                        if (textBox101.Text == "")
                        {
                            textBox101.BackColor = Color.Firebrick;
                            textBox101.ForeColor = Color.White;
                            val = false;
                        }

                    }
                    else
                    {
                        if (textBox26.Text == "" && textBox86.Text == "" && textBox91.Text == "" && textBox96.Text == "" && textBox101.Text == "")
                        {
                            textBox26.BackColor = Color.White;
                            textBox26.ForeColor = Color.Black;
                            textBox86.BackColor = Color.White;
                            textBox86.ForeColor = Color.Black;
                            textBox91.BackColor = Color.White;
                            textBox91.ForeColor = Color.Black;
                            textBox96.BackColor = Color.White;
                            textBox96.ForeColor = Color.Black;
                            textBox101.BackColor = Color.White;
                            textBox101.ForeColor = Color.Black;

                        }
                    }
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
            return val;
        }

        private void comboBox1_Enter(object sender, EventArgs e)
        {
           // comboBox1.BackColor = Color.White;
        }

        private void comboBox2_Enter(object sender, EventArgs e)
        {
            comboBox2.BackColor = Color.White;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            textBox4.BackColor = Color.White;
            textBox4.ForeColor = Color.Black;
        }

        private void comboBox3_Enter(object sender, EventArgs e)
        {
            comboBox3.BackColor = Color.White;
        }

        private void comboBox5_Enter(object sender, EventArgs e)
        {
            comboBox5.BackColor = Color.White;
        }

        private void comboBox6_Enter(object sender, EventArgs e)
        {
            comboBox6.BackColor = Color.White;
        }

        private void comboBox7_Enter(object sender, EventArgs e)
        {
            comboBox7.BackColor = Color.White;
        }

        private void textBox67_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox67);
            textBox67.BackColor = Color.White;
            textBox67.ForeColor = Color.Black;
            bool res = int.TryParse(textBox72.Text, out int val);
            if((val>=1601 && val<1606)|| (val >= 1701 && val < 1708) || (val >= 1801 && val < 1819))
            {
                if (textBox67.Text != "")
                {
                   // MessageBox.Show("Це поле повинне бути пустим!");
                    textBox67.Text = "";
                }
            }
           /* bool val_cur = int.TryParse(textBox67.Text, out int res1);
            bool val_cur2 = int.TryParse(textBox68.Text, out int res2);

            if (res1 > res2)
            {
                if (res2 != 201 && res2 != 202 && res2 != 203 && res2 != 204)
                {

                        textBox68.BackColor = Color.White;
                        textBox68.ForeColor = Color.Black;  
                }

            }*/

        }

        private void textBox68_TextChanged(object sender, EventArgs e)
        {
            int x;
            try
            {
                if (textBox68.Text != "")
                {
                    comboBox8.BackColor = Color.White;
                    if (textBox68.Text != "-")
                        x = Math.Abs(int.Parse(textBox68.Text));
                }
                
            }
            catch (Exception)
            {

                MessageBox.Show("В це поле можно ввести тільки цифри!");
            }
           
            textBox68.BackColor = Color.White;
            textBox68.ForeColor = Color.Black;

            bool res = int.TryParse(textBox72.Text, out int val);
            if ((val >= 1601 && val < 1606) || (val >= 1701 && val < 1708) || (val >= 1801 && val < 1819))
            {
                if (textBox68.Text != "")
                {
                    MessageBox.Show("Це поле повинне бути пустим!");
                    textBox68.Text = "";
                }
            }
        }

        private void comboBox8_Enter(object sender, EventArgs e)
        {
            comboBox8.BackColor = Color.White;
            
        }

        private void comboBox9_Enter(object sender, EventArgs e)
        {
            comboBox9.BackColor = Color.White;
        }

        private void comboBox10_Enter(object sender, EventArgs e)
        {
            comboBox10.BackColor = Color.White;
        }

        private void comboBox11_Enter(object sender, EventArgs e)
        {
            comboBox11.BackColor = Color.White;
        }

        private void comboBox12_Enter(object sender, EventArgs e)
        {
            comboBox12.BackColor = Color.White;
        }

        private void comboBox24_Enter(object sender, EventArgs e)
        {
            comboBox24.BackColor = Color.White;
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox8.SelectedItem != null)
            {
                textBox68.BackColor = Color.White;
                textBox68.ForeColor = Color.Black;
            }
            textBox68.Text = CodeAllContent("poverhovist", "name_poverh", "code_poverh", comboBox8.SelectedItem);
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
             if(textBox18.Text=="")
            {
                textBox17.BackColor = Color.White;
                textBox17.ForeColor = Color.Black;
                comboBox14.BackColor = Color.White;
                comboBox19.BackColor = Color.White;
                comboBox25.BackColor = Color.White;
                comboBox30.BackColor = Color.White;
            }
            else
            {
                textBox18.BackColor = Color.White;
                textBox18.ForeColor = Color.Black;
            }
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox18.Text = "";
            }
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            if (textBox20.Text == "")
            {
                textBox19.BackColor = Color.White;
                textBox19.ForeColor = Color.Black;
                comboBox15.BackColor = Color.White;
                comboBox20.BackColor = Color.White;
                comboBox26.BackColor = Color.White;
                comboBox31.BackColor = Color.White;
            }
            else
            {
                textBox20.BackColor = Color.White;
                textBox20.ForeColor = Color.Black;
            }
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox20.Text = "";
            }
        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            if (textBox22.Text == "")
            {
                textBox21.BackColor = Color.White;
                textBox21.ForeColor = Color.Black;
                comboBox16.BackColor = Color.White;
                comboBox21.BackColor = Color.White;
                comboBox27.BackColor = Color.White;
                comboBox32.BackColor = Color.White;
            }
            else
            {
                textBox22.BackColor = Color.White;
                textBox22.ForeColor = Color.Black;
            }
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox22.Text = "";
            }
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            if (textBox24.Text == "")
            {
                textBox23.BackColor = Color.White;
                textBox23.ForeColor = Color.Black;
                comboBox17.BackColor = Color.White;
                comboBox22.BackColor = Color.White;
                comboBox28.BackColor = Color.White;
                comboBox33.BackColor = Color.White;
            }
            else
            {
                textBox24.BackColor = Color.White;
                textBox24.ForeColor = Color.Black;
            }
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox24.Text = "";
            }
        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {
            if (textBox26.Text == "")
            {
               
                textBox25.BackColor = Color.White;
                textBox25.ForeColor = Color.Black;
                comboBox18.BackColor = Color.White;
                comboBox23.BackColor = Color.White;
                comboBox29.BackColor = Color.White;
                comboBox34.BackColor = Color.White;
            }
            else
            {
                textBox26.BackColor = Color.White;
                textBox26.ForeColor = Color.Black;
            }
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox26.Text = "";
            }
        }

        private void comboBox14_Enter(object sender, EventArgs e)
        {
            comboBox14.BackColor = Color.White;
        }

        private void comboBox15_Enter(object sender, EventArgs e)
        {
            comboBox15.BackColor = Color.White;
        }

        private void comboBox16_Enter(object sender, EventArgs e)
        {
            comboBox16.BackColor = Color.White;
        }

        private void comboBox17_Enter(object sender, EventArgs e)
        {
            comboBox17.BackColor = Color.White;
        }

        private void comboBox18_Enter(object sender, EventArgs e)
        {
            comboBox18.BackColor = Color.White;
        }

        private void comboBox19_Enter(object sender, EventArgs e)
        {
            comboBox19.BackColor = Color.White;
        }

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] str = comboBox20.SelectedItem.ToString().Split(';');
            //comboBox20.BackColor = Color.White;
            textBox88.Text = CodeAllContent("social_status", "name_status", "code_status", str[1]);
            //textBox87.Text = CodeAllContent("social_status", "name_status", "code_status", comboBox19.SelectedItem);
            textBox88.BackColor = Color.White;
            textBox88.ForeColor = Color.Black;
        }

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] str = comboBox21.SelectedItem.ToString().Split(';');
            comboBox21.BackColor = Color.White;
            textBox89.Text = CodeAllContent("social_status", "name_status", "code_status", str[1]);
            textBox89.BackColor = Color.White;
            textBox89.ForeColor = Color.Black;
        }

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] str = comboBox22.SelectedItem.ToString().Split(';');
            comboBox22.BackColor = Color.White;
            textBox90.Text = CodeAllContent("social_status", "name_status", "code_status", str[1]);
            textBox90.BackColor = Color.White;
            textBox90.ForeColor = Color.Black;
        }

        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] str = comboBox23.SelectedItem.ToString().Split(';');
            comboBox23.BackColor = Color.White;
            textBox91.Text = CodeAllContent("social_status", "name_status", "code_status", str[1]);
            textBox91.BackColor = Color.White;
            textBox91.ForeColor = Color.Black;
        }

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox25.BackColor = Color.White;
            textBox92.Text = CodeAllContent("moment_smerti", "moment", "code_moment", comboBox25.SelectedItem);
            textBox92.BackColor = Color.White;
            textBox92.ForeColor = Color.Black;
        }

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox26.BackColor = Color.White;
            textBox93.Text = CodeAllContent("moment_smerti", "moment", "code_moment", comboBox26.SelectedItem);
            textBox93.BackColor = Color.White;
            textBox93.ForeColor = Color.Black;
        }

        private void comboBox27_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox27.BackColor = Color.White;
            textBox94.Text = CodeAllContent("moment_smerti", "moment", "code_moment", comboBox27.SelectedItem);
            textBox94.BackColor = Color.White;
            textBox94.ForeColor = Color.Black;
        }

        private void comboBox28_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox28.BackColor = Color.White;
            textBox95.Text = CodeAllContent("moment_smerti", "moment", "code_moment", comboBox28.SelectedItem);
            textBox95.BackColor = Color.White;
            textBox95.ForeColor = Color.Black;
        }

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox29.BackColor = Color.White;
            textBox96.Text = CodeAllContent("moment_smerti", "moment", "code_moment", comboBox29.SelectedItem);
            textBox96.BackColor = Color.White;
            textBox96.ForeColor = Color.Black;
        }

        private void comboBox30_Enter(object sender, EventArgs e)
        {
            comboBox30.BackColor = Color.White;
        }

        private void comboBox31_Enter(object sender, EventArgs e)
        {
            comboBox31.BackColor = Color.White;
        }

        private void comboBox32_Enter(object sender, EventArgs e)
        {
            comboBox32.BackColor = Color.White;
        }

        private void comboBox33_Enter(object sender, EventArgs e)
        {
            comboBox33.BackColor = Color.White;
        }

        private void comboBox34_Enter(object sender, EventArgs e)
        {
            comboBox34.BackColor = Color.White;
        }

        private void comboBox36_Enter(object sender, EventArgs e)
        {
            comboBox36.BackColor = Color.White;
        }

        private void comboBox45_Enter(object sender, EventArgs e)
        {
            comboBox45.BackColor = Color.White;
        }

        private void comboBox46_Enter(object sender, EventArgs e)
        {
            comboBox46.BackColor = Color.White;
        }

        private void comboBox35_Enter(object sender, EventArgs e)
        {
            comboBox35.BackColor = Color.White;
        }

        private void comboBox47_Enter(object sender, EventArgs e)
        {
            comboBox47.BackColor = Color.White;
        }

        private void comboBox52_Enter(object sender, EventArgs e)
        {
            comboBox52.BackColor = Color.White;
        }

        private void comboBox57_Enter(object sender, EventArgs e)
        {
            comboBox57.BackColor = Color.White;
        }

        private void comboBox66_Enter(object sender, EventArgs e)
        {
            comboBox66.BackColor = Color.White;
        }

        private void comboBox67_Enter(object sender, EventArgs e)
        {
            comboBox67.BackColor = Color.White;
        }

        private void textBox61_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox61);
            textBox61.BackColor = Color.White;
            textBox61.ForeColor = Color.Black;
        }

        private void textBox60_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox60);
            textBox60.BackColor = Color.White;
            textBox60.ForeColor = Color.Black;
        }

        private void textBox59_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox59);
            textBox59.BackColor = Color.White;
            textBox59.ForeColor = Color.Black;
        }

        private void comboBox72_Enter(object sender, EventArgs e)
        {
            comboBox72.BackColor = Color.White;
        }

        private void comboBox75_Enter(object sender, EventArgs e)
        {
            comboBox75.BackColor = Color.White;
        }

        private void comboBox78_Enter(object sender, EventArgs e)
        {
            comboBox78.BackColor = Color.White;
        }

        private void comboBox58_Enter(object sender, EventArgs e)
        {
            comboBox58.BackColor = Color.White;
        }

        private void comboBox65_Enter(object sender, EventArgs e)
        {
            comboBox65.BackColor = Color.White;
        }

        private void comboBox68_Enter(object sender, EventArgs e)
        {
            comboBox68.BackColor = Color.White;
        }

        private void comboBox59_Enter(object sender, EventArgs e)
        {
            comboBox59.BackColor = Color.White;
        }

        private void comboBox60_Enter(object sender, EventArgs e)
        {
            comboBox60.BackColor = Color.White;
        }

        private void comboBox61_Enter(object sender, EventArgs e)
        {
            comboBox61.BackColor = Color.White;
        }

        private void comboBox64_Enter(object sender, EventArgs e)
        {
            comboBox64.BackColor = Color.White;
        }

        private void comboBox63_Enter(object sender, EventArgs e)
        {
            comboBox63.BackColor = Color.White;
        }

        private void comboBox62_Enter(object sender, EventArgs e)
        {
            comboBox62.BackColor = Color.White;
        }

        private void comboBox69_Enter(object sender, EventArgs e)
        {
            comboBox69.BackColor = Color.White;
        }

        private void comboBox79_Enter(object sender, EventArgs e)
        {
            comboBox79.BackColor = Color.White;
        }

        private void comboBox80_Enter(object sender, EventArgs e)
        {
            comboBox80.BackColor = Color.White;
        }

        private void comboBox81_Enter(object sender, EventArgs e)
        {
            comboBox81.BackColor = Color.White;
        }

        private void label91_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Регіон", label91);
        }

        private void label92_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Район", label92);
        }

        private void label93_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Тип населеного пункту", label93);
        }

        private void label94_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Номер картки", label94);
        }

        private void label95_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Дата виникнення пожежі ", label95);
        }

        private void label97_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Адреса пожежі, з назвою об’єкту пожежі", label97);
        }

        private void label96_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Об’єкт пожежі ", label96);
        }

        private void label98_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Форма власності", label98);
        }

        private void label99_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Ступінь ризику господарської діяльності", label99);
        }

        private void label10_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Підконтрольність об’єкта", label10);
        }

        private void label12_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Ступінь вогнестійкості", label12);
        }

        private void label13_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Категорія небезпеки", label13);
        }

        private void label14_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Місце виникнення пожежі ", label14);
        }

        private void label15_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Виріб-ініціатор пожежі", label15);
        }

        private void label16_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Причина пожежі", label16);
        }

       

        private void textBox5_MouseHover(object sender, EventArgs e)
        {
            //toolTip1.Show("Виявлено загиблих на місці пожежі ", textBox5);
        }

       
        private void textBox6_MouseHover(object sender, EventArgs e)
        {
            //toolTip1.Show("Виявлено загиблих дітей і підлітків", textBox6);
        }

        private void textBox8_MouseHover(object sender, EventArgs e)
        {
           // toolTip1.Show("Загинуло внаслідок пожежі", textBox8);
        }

        private void label17_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Виявлено загиблих", label17);
        }

        private void label18_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Загинуло внаслідок пожежі", label18);
        }

        private void textBox82_MouseHover(object sender, EventArgs e)
        {
            //toolTip1.Show("Загинуло внаслідок пожежі дітей і підлітків", textBox82);
        }

        private void textBox10_MouseHover(object sender, EventArgs e)
        {
           // toolTip1.Show("Загинуло внаслідок пожежі особового складу", textBox65);
        }

        private void textBox12_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Травмовано на пожежі ", textBox12);
        }

        private void label19_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Травмовано на пожежі ", label19);
        }

        private void textBox11_MouseHover(object sender, EventArgs e)
        {
            //toolTip1.Show("Травмовано на пожежі дітей і підлітків ", textBox11);
        }

        private void textBox66_MouseHover(object sender, EventArgs e)
        {
            //toolTip1.Show("особового складу", textBox66);
        }

       

        private void textBox69_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Регіон";
        }

        private void label22_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Прямий збиток від пожежі", label22);
        }

        private void textBox9_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Прямий збиток від пожежі";
            bool val25 = float.TryParse(textBox9.Text, out float res25);
            if (res25 > 0 && textBox9.BackColor == Color.Firebrick)
            {
                MessageBox.Show("В одній із позицій 27-35 має бути проставлено значення чи текстова інформація");
            }
            textBox9.BackColor = Color.White;
            textBox9.ForeColor = Color.Black;

           
        }

        private void label23_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Побічний збиток від пожежі", label23);
        }

        private void textBox15_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Побічний збиток від пожежі";
            textBox15.BackColor = Color.White;
            textBox15.ForeColor = Color.Black;
            if (textBox9.Text == "")
            {
                MessageBox.Show("Введіть сумму прямого збитку!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox9.Select();
                textBox9.ScrollToCaret();
            }
        }

        private void label24_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Знищено будинків(споруд)", label24);
        }

        private void textBox7_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Знищено будинків(споруд)";
            textBox7.BackColor = Color.White;
            textBox7.ForeColor = Color.Black;
        }

        private void textBox14_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пошкоджено будинків(споруд)";
            textBox14.BackColor = Color.White;
            textBox14.ForeColor = Color.Black;
        }
        private void textBox13_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Знищено техніки";
            textBox13.BackColor = Color.White;
            textBox13.ForeColor = Color.Black;
        }

        private void textBox16_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пошкоджено техніки";
            textBox16.BackColor = Color.White;
            textBox16.ForeColor = Color.Black;
        }

        private void textBox28_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Знищено зернових та технічних культур";
            textBox28.BackColor = Color.White;
            textBox28.ForeColor = Color.Black;
        }

        private void textBox27_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Знищено хліба на корені";
            textBox27.BackColor = Color.White;
            textBox27.ForeColor = Color.Black;
        }

        private void textBox36_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Знищено хліба у валках";
            textBox36.BackColor = Color.White;
            textBox36.ForeColor = Color.Black;
        }

        private void textBox30_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Знищено кормів";
            textBox30.BackColor = Color.White;
            textBox30.ForeColor = Color.Black;
        }

        private void textBox29_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Знищено торфовищ";
            textBox29.BackColor = Color.White;
            textBox29.ForeColor = Color.Black;
        }

        private void textBox35_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пошкоджено торфовищ";
            textBox35.BackColor = Color.White;
            textBox35.ForeColor = Color.Black;
        }

        private void textBox32_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Загинуло тварин";
            textBox32.BackColor = Color.White;
            textBox32.ForeColor = Color.Black;
        }

        private void textBox31_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Загинуло птиці";
            textBox31.BackColor = Color.White;
            textBox31.ForeColor = Color.Black;
        }

        private void textBox33_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Додатково, що ще знищено(пошкоджено)";
            textBox33.BackColor = Color.White;
            textBox33.ForeColor = Color.Black;
        }

        private void textBox37_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано людей";
            textBox37.BackColor = Color.White;
            textBox37.ForeColor = Color.Black;
        }

        private void textBox34_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано дітей і підлітків ";
            textBox34.BackColor = Color.White;
            textBox34.ForeColor = Color.Black;
        }

        private void textBox39_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано тварин";
            textBox39.BackColor = Color.White;
            textBox39.ForeColor = Color.Black;
        }

        private void textBox38_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано  птиці";
            textBox38.BackColor = Color.White;
            textBox38.ForeColor = Color.Black;
        }

        private void textBox41_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано будинків";
            textBox41.BackColor = Color.White;
            textBox41.ForeColor = Color.Black;
        }

        private void textBox40_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано техніки";
            textBox40.BackColor = Color.White;
            textBox40.ForeColor = Color.Black;
        }

        private void textBox44_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано зернових і технічних культур";
            textBox44.BackColor = Color.White;
            textBox44.ForeColor = Color.Black;
        }

        private void textBox46_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано хліба на корені";
            textBox46.BackColor = Color.White;
            textBox46.ForeColor = Color.Black;
        }

        private void textBox45_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано хліба у валках";
            textBox45.BackColor = Color.White;
            textBox45.ForeColor = Color.Black;
        }

        private void textBox43_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано кормів";
            textBox43.BackColor = Color.White;
            textBox43.ForeColor = Color.Black;
        }

        private void textBox42_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано торфовищ";
            textBox42.BackColor = Color.White;
            textBox42.ForeColor = Color.Black;
        }

        private void textBox47_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Додаткова інформація (що ще врятовано)";
            textBox47.BackColor = Color.White;
            textBox47.ForeColor = Color.Black;
        }

        private void textBox48_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Врятовано матеріальних цінностей на суму";
            bool val25 = float.TryParse(textBox48.Text, out float res25);
            if (res25 > 0 && textBox48.BackColor == Color.Firebrick)
            {
                MessageBox.Show("В одній із позицій 36-45 має бути проставлено значення чи текстова інформація","Помилка",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            textBox48.BackColor = Color.White;
            textBox48.ForeColor = Color.Black;
           
        }

        private void textBox103_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на поширення пожежі";
            textBox103.BackColor = Color.White;
            textBox103.ForeColor = Color.Black;
        }

        private void textBox104_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на поширення пожежі";
            textBox104.BackColor = Color.White;
            textBox104.ForeColor = Color.Black;
        }

        private void textBox105_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на поширення пожежі";
            textBox105.BackColor = Color.White;
            textBox105.ForeColor = Color.Black;
        }

        private void textBox106_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на поширення пожежі";
            textBox106.BackColor = Color.White;
            textBox106.ForeColor = Color.Black;
        }

        private void textBox107_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на поширення пожежі";
            textBox107.BackColor = Color.White;
            textBox107.ForeColor = Color.Black;
        }

        private void textBox112_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що ускладнювали гасіння пожежі";
            textBox112.BackColor = Color.White;
            textBox112.ForeColor = Color.Black;
        }

        private void textBox111_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що ускладнювали гасіння пожежі";
            textBox111.BackColor = Color.White;
            textBox111.ForeColor = Color.Black;
        }

        private void textBox110_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що ускладнювали гасіння пожежі";
            textBox110.BackColor = Color.White;
            textBox110.ForeColor = Color.Black;
        }

        private void textBox109_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що ускладнювали гасіння пожежі";
            textBox109.BackColor = Color.White;
            textBox109.ForeColor = Color.Black;
        }

        private void textBox108_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що ускладнювали гасіння пожежі";
            textBox108.BackColor = Color.White;
            textBox108.ForeColor = Color.Black;
        }

        private void textBox113_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Наявність систем протипожежного захисту";
            textBox113.BackColor = Color.White;
            textBox113.ForeColor = Color.Black;
        }

        private void textBox118_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Системи протипожежного захисту";
            textBox118.BackColor = Color.White;
            textBox118.ForeColor = Color.Black;
        }

        private void textBox117_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Системи протипожежного захисту";
            textBox117.BackColor = Color.White;
            textBox117.ForeColor = Color.Black;
        }

        private void textBox116_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Системи протипожежного захисту";
            textBox116.BackColor = Color.White;
            textBox116.ForeColor = Color.Black;
        }

        private void textBox115_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Системи протипожежного захисту";
            textBox115.BackColor = Color.White;
            textBox115.ForeColor = Color.Black;
        }

        private void textBox114_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Системи протипожежного захисту";
            textBox114.BackColor = Color.White;
            textBox114.ForeColor = Color.Black;
        }

        private void textBox123_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Результат дії системи протипожежного захисту";
            textBox123.BackColor = Color.White;
            textBox123.ForeColor = Color.Black;
        }

        private void textBox122_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Результат дії системи протипожежного захисту";
            textBox122.BackColor = Color.White;
            textBox122.ForeColor = Color.Black;
        }

        private void textBox121_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Результат дії системи протипожежного захисту";
            textBox121.BackColor = Color.White;
            textBox121.ForeColor = Color.Black;
        }

        private void textBox120_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Результат дії системи протипожежного захисту";
            textBox120.BackColor = Color.White;
            textBox120.ForeColor = Color.Black;
        }

        private void textBox119_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Результат дії системи протипожежного захисту";
            textBox119.BackColor = Color.White;
            textBox119.ForeColor = Color.Black;
        }

        private void textBox128_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Учасник гасіння пожежі";
            textBox128.BackColor = Color.White;
            textBox128.ForeColor = Color.Black;
        }

        private void textBox127_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Учасник гасіння пожежі";
            textBox127.BackColor = Color.White;
            textBox127.ForeColor = Color.Black;
        }

        private void textBox126_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Учасник гасіння пожежі";
            textBox126.BackColor = Color.White;
            textBox126.ForeColor = Color.Black;
        }

        private void textBox125_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Учасник гасіння пожежі";
            textBox125.BackColor = Color.White;
            textBox125.ForeColor = Color.Black;
        }

        private void textBox124_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Учасник гасіння пожежі";
            textBox124.BackColor = Color.White;
            textBox124.ForeColor = Color.Black;
        }

        private void textBox49_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість учасників";
            textBox49.BackColor = Color.White;
            textBox49.ForeColor = Color.Black;
        }

        private void textBox50_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість учасників";
            textBox50.BackColor = Color.White;
            textBox50.ForeColor = Color.Black;
        }

        private void textBox51_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість учасників";
            textBox51.BackColor = Color.White;
            textBox51.ForeColor = Color.Black;
        }

        private void textBox52_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість учасників";
            textBox52.BackColor = Color.White;
            textBox52.ForeColor = Color.Black;
        }

        private void textBox53_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість учасників";
            textBox53.BackColor = Color.White;
            textBox53.ForeColor = Color.Black;
        }

        private void textBox129_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пожежні автомобілі";
            textBox129.BackColor = Color.White;
            textBox129.ForeColor = Color.Black;
        }

        private void textBox130_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пожежні автомобілі";
            textBox130.BackColor = Color.White;
            textBox130.ForeColor = Color.Black;
        }

        private void textBox131_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пожежні автомобілі";
            textBox131.BackColor = Color.White;
            textBox131.ForeColor = Color.Black;
        }

        private void textBox132_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пожежні автомобілі";
            textBox132.BackColor = Color.White;
            textBox132.ForeColor = Color.Black;
        }

        private void textBox133_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пожежні автомобілі";
            textBox133.BackColor = Color.White;
            textBox133.ForeColor = Color.Black;
        }

        private void textBox58_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість техніки";
            textBox58.BackColor = Color.White;
            textBox58.ForeColor = Color.Black;
        }

        private void textBox57_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість техніки";
            textBox57.BackColor = Color.White;
            textBox57.ForeColor = Color.Black;
        }

        private void textBox56_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість техніки";
            textBox56.BackColor = Color.White;
            textBox56.ForeColor = Color.Black;
        }

        private void textBox55_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість техніки";
            textBox55.BackColor = Color.White;
            textBox55.ForeColor = Color.Black;
        }

        private void textBox54_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість техніки";
            textBox54.BackColor = Color.White;
            textBox54.ForeColor = Color.Black;
        }

        private void textBox134_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пожежні стволи";
            textBox134.BackColor = Color.White;
            textBox134.ForeColor = Color.Black;
        }

        private void textBox135_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пожежні стволи";
            textBox135.BackColor = Color.White;
            textBox135.ForeColor = Color.Black;
        }

        private void textBox136_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Пожежні стволи";
            textBox136.BackColor = Color.White;
            textBox136.ForeColor = Color.Black;
        }

        private void textBox61_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість стволів";
            textBox61.BackColor = Color.White;
            textBox61.ForeColor = Color.Black;
        }

        private void textBox60_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість стволів";
            textBox60.BackColor = Color.White;
            textBox60.ForeColor = Color.Black;
        }

        private void textBox59_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість стволів";
            textBox59.BackColor = Color.White;
            textBox59.ForeColor = Color.Black;
        }

        private void textBox137_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вогнегасні речовини";
            textBox137.BackColor = Color.White;
            textBox137.ForeColor = Color.Black;
        }

        private void textBox138_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вогнегасні речовини";
            textBox138.BackColor = Color.White;
            textBox138.ForeColor = Color.Black;
        }

        private void textBox139_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вогнегасні речовини";
            textBox139.BackColor = Color.White;
            textBox139.ForeColor = Color.Black;
        }

        private void textBox140_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Первинні засоби пожежогасіння";
            textBox140.BackColor = Color.White;
            textBox140.ForeColor = Color.Black;
        }

        private void textBox141_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Первинні засоби пожежогасіння";
            textBox141.BackColor = Color.White;
            textBox141.ForeColor = Color.Black;
        }

        private void textBox142_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Первинні засоби пожежогасіння";
            textBox142.BackColor = Color.White;
            textBox142.ForeColor = Color.Black;
        }

        private void textBox143_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Джерела водопостачання";
            textBox143.BackColor = Color.White;
            textBox143.ForeColor = Color.Black;
        }

        private void textBox144_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Джерела водопостачання";
            textBox144.BackColor = Color.White;
            textBox144.ForeColor = Color.Black;
        }

        private void textBox145_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Джерела водопостачання";
            textBox145.BackColor = Color.White;
            textBox145.ForeColor = Color.Black;
        }

        private void textBox62_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Використання ГДЗС";
            textBox62.BackColor = Color.White;
            textBox62.ForeColor = Color.Black;
        }

        private void textBox63_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Кількість ланок";
            textBox63.BackColor = Color.White;
            textBox63.ForeColor = Color.Black;
        }

        private void textBox64_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Загальний час роботи ГДЗС";
            textBox64.BackColor = Color.White;
            textBox64.ForeColor = Color.Black;
        }

        private void textBox146_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вид перевірки";
            textBox146.BackColor = Color.White;
            textBox146.ForeColor = Color.Black;
            bool val_cur10 = int.TryParse(textBox76.Text, out int res10);
           /* if (res10 == 11 || res10 == 12 || res10 == 13)
            {
                if (textBox146.Text == "")
                {
                    MessageBox.Show("Це поле не може бути пустим!");
                }
            }*/
           
        }

        private void textBox147_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови господарської діяльністі";
            textBox147.BackColor = Color.White;
            textBox147.ForeColor = Color.Black;
        }

        private void textBox148_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вжиті заходи за фактом пожежі";
            textBox148.BackColor = Color.White;
            textBox148.ForeColor = Color.Black;
        }

        private void textBox149_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вжиті заходи за фактом пожежі";
            textBox149.BackColor = Color.White;
            textBox149.ForeColor = Color.Black;
        }

        private void textBox150_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "П.І.Б посадової особи, яка заповнила картку";
            textBox150.BackColor = Color.White;
            textBox150.ForeColor = Color.Black;
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Район";
            textBox1.BackColor = Color.White;
            textBox1.ForeColor = Color.Black;
        }

        private void textBox70_Enter(object sender, EventArgs e)
        {
            textBox70.BackColor = Color.White;
            textBox70.ForeColor = Color.Black;
            toolStripStatusLabel1.Text = "Тип населеного пункту";
            string code = CodeAllContent("current_raion", "code_raion", "code_raion", textBox1.Text);
            // if (textBox71.Text != "")

            if (textBox1.Text == "")
            {
                textBox1.Select();
                textBox1.ScrollToCaret();
                MessageBox.Show("Це поле обов'язкове для заповнення!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (code == "")
            {
                MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox1.Text = "";
                textBox1.BackColor = Color.Firebrick;
                textBox1.Select();
                textBox1.ScrollToCaret();
            }
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            textBox2.BackColor = Color.White;
            textBox2.ForeColor = Color.Black;
            toolStripStatusLabel1.Text = "Номер картки";

            string code = CodeAllContent("type_region", "code_region_item", "code_region_item", textBox70.Text);
            // if (textBox71.Text != "")
            if (textBox70.Text == "")
            {
                textBox70.Select();
                textBox70.ScrollToCaret();
                MessageBox.Show("Це поле обов'язкове для заповнення!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (code == "")
            {
                MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox70.Text = "";
                textBox70.BackColor = Color.Firebrick;
                textBox70.Select();
                textBox70.ScrollToCaret();
            }
        }

        private void textBox71_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Код населеного пункту";
            textBox71.BackColor = Color.White;
            textBox71.ForeColor = Color.Black;
            if (dateTimePicker1.Value.ToString() == "01.01.1900 0:00:00")
            {
                MessageBox.Show("Заповніть дату виникнення пожежі");
                dateTimePicker1.Select();
            }
        }

        private void textBox4_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Адреса пожежі, з назвою об’єкту пожежі";
            textBox4.BackColor = Color.White;
            textBox4.ForeColor = Color.Black;

            string code = CodeAllContent("current_np", "code_np", "code_np", textBox71.Text);
            // if (textBox71.Text != "")
            if (code == "")
            {
                MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox71.Text = "";
                textBox71.BackColor = Color.Firebrick;
                textBox71.Select();
                textBox71.ScrollToCaret();
            }
        }

        private void textBox72_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Об’єкт пожежі";
            textBox72.BackColor = Color.White;
            textBox72.ForeColor = Color.Black;
        }

        private void textBox74_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Форма власності";
            textBox74.BackColor = Color.White;
            textBox74.ForeColor = Color.Black;
        }

        private void textBox102_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Інформація про ліквідацію пожежі ";
            textBox102.BackColor = Color.White;
            textBox102.ForeColor = Color.Black;
        }

        private void textBox75_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Ступінь ризику господарської діяльності";
            textBox75.BackColor = Color.White;
            textBox75.ForeColor = Color.Black;
        }

        private void textBox76_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Підконтрольність об’єкта";
            textBox76.BackColor = Color.White;
            textBox76.ForeColor = Color.Black;
        }

        private void textBox76_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox76);
            if (textBox147.Text != "")
            {
                if (textBox76.Text == "20")
                {
                    MessageBox.Show("В це поле неможливо поставити код 20");
                    textBox76.Text = "";
                }
            }
            bool val_cur10 = int.TryParse(textBox76.Text, out int res10);
            if (res10 == 11 || res10 == 12 || res10 == 13)
            {
                if (textBox146.Text == "")
                {
                    textBox146.BackColor = Color.Firebrick;

                }
            }
            else
            {
                textBox146.BackColor = Color.White;
            }
            if (textBox76.Text == "")
            {
                dateTimePicker9.Enabled = false;
                dateTimePicker9.Value = DateTime.Parse("01.01.1900");
                maskedTextBox5.Enabled = false;
            }
            else
            {
                dateTimePicker9.Enabled = true;
                maskedTextBox5.Enabled = true;
            }
        }

        private void textBox67_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Поверхів";
            textBox67.BackColor = Color.White;
            textBox67.ForeColor = Color.Black;

            bool val_cur = int.TryParse(textBox67.Text, out int res);
            bool val_cur2 = int.TryParse(textBox68.Text, out int res1);
            
            if (res<res1)
            {
                if(res1!= 201 && res1!= 202 && res1!= 203 && res1!= 204)
                {
                   // Console.WriteLine(res1);
                    if (textBox67.Text != "")
                    {
                        MessageBox.Show("Поверх пожежі не може бути більшим, ніж поверхів у будинку");
                        textBox67.BackColor = Color.Firebrick;
                        textBox67.ForeColor = Color.White;
                        textBox67.Text = "";
                    }
                }
               
            }
            bool res2 = int.TryParse(textBox72.Text, out int val);
            if ((val >= 1601 && val < 1606) || (val >= 1701 && val < 1708) || (val >= 1801 && val < 1819))
            {
                if (textBox67.Text != "")
                {
                    MessageBox.Show("Це поле повинне бути пустим!");
                    textBox67.Text = "";
                }
            }

        }

        private void textBox68_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Поверх, де виникла пожежа";
            textBox68.BackColor = Color.White;
            textBox68.ForeColor = Color.Black;

            bool res = int.TryParse(textBox72.Text, out int val);
            if ((val >= 1601 && val < 1606) || (val >= 1701 && val < 1708) || (val >= 1801 && val < 1819))
            {
                if (textBox68.Text != "")
                {
                    MessageBox.Show("Це поле повинне бути пустим!");
                    textBox68.Text = "";
                }
            }
        }

        private void textBox77_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Ступінь вогнестійкості";
            textBox77.BackColor = Color.White;
            textBox77.ForeColor = Color.Black;
        }

        private void textBox78_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Категорія небезпеки";
            textBox78.BackColor = Color.White;
            textBox78.ForeColor = Color.Black;

           /* bool val_13 = int.TryParse(textBox72.Text, out int res13);
                if (res13 >= 1 && res13 < 210)
                {
                    if (textBox78.Text == "")
                    {
                        MessageBox.Show("Це поле обов'язкове для заповнення!");
                    }
                }*/
        }
        

        private void textBox79_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Місце виникнення пожежі";
            textBox79.BackColor = Color.White;
            textBox79.ForeColor = Color.Black;
        }

        private void textBox80_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Виріб-ініціатор пожежі";
            textBox80.BackColor = Color.White;
            textBox80.ForeColor = Color.Black;
        }

        private void textBox81_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Причина пожежі";
            textBox81.BackColor = Color.White;
            textBox81.ForeColor = Color.Black;
        }

        private void textBox5_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Виявлено загиблих";
            textBox5.BackColor = Color.White;
            textBox5.ForeColor = Color.Black;

        }

        private void textBox6_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Виявлено загиблих дітей і підлітків";
            textBox6.BackColor = Color.White;
            textBox6.ForeColor = Color.Black;
        }

        private void textBox8_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = " Загинуло внаслідок пожежі";
            textBox8.BackColor = Color.White;
            textBox8.ForeColor = Color.Black;
        }

        private void textBox65_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Загинуло внаслідок пожежі дітей і підлітків";
            if (textBox65.BackColor == Color.Firebrick)
            {
                MessageBox.Show("Сума цього та наступного поля має бути меншою або дорівнювати загальній кількості травмованих!", "Помилка", MessageBoxButtons.OK,MessageBoxIcon.Error);
                textBox65.Text = "";
            }
            textBox65.BackColor = Color.White;
            textBox65.ForeColor = Color.Black;
        }

        private void textBox82_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Загинуло внаслідок пожежі особового складу";
            if (textBox82.BackColor == Color.Firebrick)
            {
                MessageBox.Show("Сума цього та наступного поля має бути меншою або дорівнювати загальній кількості травмованих!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox82.Text = "";
            }
            textBox82.BackColor = Color.White;
            textBox82.ForeColor = Color.Black;
        }

        private void textBox18_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "П.І.Б. загиблого унаслідок пожежі";
            textBox18.BackColor = Color.White;
            textBox18.ForeColor = Color.Black;
            
        }

        private void textBox20_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "П.І.Б. загиблого унаслідок пожежі";
            textBox20.BackColor = Color.White;
            textBox20.ForeColor = Color.Black;
        }

        private void textBox22_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "П.І.Б. загиблого унаслідок пожежі";
            textBox22.BackColor = Color.White;
            textBox22.ForeColor = Color.Black;
        }

        private void textBox24_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "П.І.Б. загиблого унаслідок пожежі";
            textBox24.BackColor = Color.White;
            textBox24.ForeColor = Color.Black;
        }

        private void textBox26_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "П.І.Б. загиблого унаслідок пожежі";
            textBox26.BackColor = Color.White;
            textBox26.ForeColor = Color.Black;
        }

        private void textBox17_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вік загиблого унаслідок пожежі";
            textBox17.BackColor = Color.White;
            textBox17.ForeColor = Color.Black;
        }

        private void textBox19_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вік загиблого унаслідок пожежі";
            textBox19.BackColor = Color.White;
            textBox19.ForeColor = Color.Black;
        }

        private void textBox21_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вік загиблого унаслідок пожежі";
            textBox21.BackColor = Color.White;
            textBox21.ForeColor = Color.Black;
        }

        private void textBox23_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вік загиблого унаслідок пожежі";
            textBox23.BackColor = Color.White;
            textBox23.ForeColor = Color.Black;
        }

        private void textBox25_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Вік загиблого унаслідок пожежі";
        }

        private void textBox10_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Стать загиблого унаслідок пожежі";
            textBox10.BackColor = Color.White;
            textBox10.ForeColor = Color.Black;
        }

        private void textBox83_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Стать загиблого унаслідок пожежі";
            textBox83.BackColor = Color.White;
            textBox83.ForeColor = Color.Black;
        }

        private void textBox84_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Стать загиблого унаслідок пожежі";
            textBox84.BackColor = Color.White;
            textBox84.ForeColor = Color.Black;
        }

        private void textBox85_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Стать загиблого унаслідок пожежі";
            textBox85.BackColor = Color.White;
            textBox85.ForeColor = Color.Black;
        }

        private void textBox86_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Стать загиблого унаслідок пожежі";
            textBox86.BackColor = Color.White;
            textBox86.ForeColor = Color.Black;
        }

        private void textBox87_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Соціальний статус загиблого унаслідок пожежі";
            textBox87.BackColor = Color.White;
            textBox87.ForeColor = Color.Black;
        }

        private void textBox88_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Соціальний статус загиблого унаслідок пожежі";
            textBox88.BackColor = Color.White;
            textBox88.ForeColor = Color.Black;
        }

        private void textBox89_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Соціальний статус загиблого унаслідок пожежі";
            textBox89.BackColor = Color.White;
            textBox89.ForeColor = Color.Black;
        }

        private void textBox90_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Соціальний статус загиблого унаслідок пожежі";
            textBox90.BackColor = Color.White;
            textBox90.ForeColor = Color.Black;
        }

        private void textBox91_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Соціальний статус загиблого унаслідок пожежі";
            textBox91.BackColor = Color.White;
            textBox91.ForeColor = Color.Black;
        }

        private void textBox92_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Момент настання смерті";
            textBox92.BackColor = Color.White;
            textBox92.ForeColor = Color.Black;
        }

        private void textBox93_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Момент настання смерті";
            textBox93.BackColor = Color.White;
            textBox93.ForeColor = Color.Black;
        }

        private void textBox94_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Момент настання смерті";
            textBox94.BackColor = Color.White;
            textBox94.ForeColor = Color.Black;
        }

        private void textBox95_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Момент настання смерті";
            textBox95.BackColor = Color.White;
            textBox95.ForeColor = Color.Black;
        }

        private void textBox96_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Момент настання смерті";
            textBox96.BackColor = Color.White;
            textBox96.ForeColor = Color.Black;
        }

        private void textBox97_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на загибель людей";
            textBox97.BackColor = Color.White;
            textBox97.ForeColor = Color.Black;
        }

        private void textBox98_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на загибель людей";
            textBox98.BackColor = Color.White;
            textBox98.ForeColor = Color.Black;
        }

        private void textBox99_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на загибель людей";
            textBox99.BackColor = Color.White;
            textBox99.ForeColor = Color.Black;
        }

        private void textBox100_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на загибель людей";
            textBox100.BackColor = Color.White;
            textBox100.ForeColor = Color.Black;
        }

        private void textBox101_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Умови, що вплинули на загибель людей";
            textBox101.BackColor = Color.White;
            textBox101.ForeColor = Color.Black;
        }

        private void textBox12_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Травмовано на пожежі";
            textBox12.BackColor = Color.White;
            textBox12.ForeColor = Color.Black;
            
        }

        private void textBox11_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Травмовано на пожежі дітей і підлітків";
            if (textBox11.BackColor == Color.Firebrick)
            {
                MessageBox.Show("Сума цього та наступного поля має бути меншою або дорівнювати загальній кількості травмованих!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox11.Text = "";
            }
            textBox11.BackColor = Color.White;
            textBox11.ForeColor = Color.Black;
            
        }

        private void textBox66_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Травмовано на пожежі особового складу";
            if (textBox66.BackColor == Color.Firebrick)
            {
                MessageBox.Show("Сума цього та наступного поля має бути меншою або дорівнювати загальній кількості травмованих!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox66.Text = "";
            }
            textBox66.BackColor = Color.White;
            textBox66.ForeColor = Color.Black;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //textBox69.Text = CodeAllContent("region", "name_region", "code_region", comboBox1.SelectedItem);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox70.Text = CodeAllContent("type_region", "type_region_item", "code_region_item", comboBox2.SelectedItem);
            textBox70.BackColor = Color.White;
            textBox70.ForeColor = Color.Black;
        }

       

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] str = comboBox4.SelectedItem.ToString().Split(';');
            textBox72.Text = CodeAllContent("fire_objects", "fire_item", "fire_code", str[1]);
            bool res = int.TryParse(textBox72.Text, out int code);
            if (code == 1904 || code == 1818 || code == 1815 || code == 1519 || code == 1404 || code == 1309 || code == 1216 || code == 1114 || code == 1003 || code == 911 || code == 804 || code == 712 || code == 606 || code == 512 || code == 408
               || code == 310 || code == 308 || code == 209 || code == 117 || code == 16)
            {
                textBox73.Text = "";
            }
            else
            {
                textBox73.Text = str[1];
            }
         
            textBox72.BackColor = Color.White;
            textBox72.ForeColor = Color.Black;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox74.Text = CodeAllContent("forma_vlasnosti", "name_forma", "code_forma", comboBox5.SelectedItem);
            textBox74.BackColor = Color.White;
            textBox74.ForeColor = Color.Black;
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox75.Text = CodeAllContent("stupin_riziku", "name_riziku", "code_riziku", comboBox6.SelectedItem);
            textBox75.BackColor = Color.White;
            textBox75.ForeColor = Color.Black;
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox76.Text = CodeAllContent("pidkontrol_object", "name_object", "code_object", comboBox7.SelectedItem);
            textBox76.BackColor = Color.White;
            textBox76.ForeColor = Color.Black;
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox77.Text = CodeAllContent("vognestoikist", "name_stoikist", "code_stoikist", comboBox9.SelectedItem);
            textBox77.BackColor = Color.White;
            textBox77.ForeColor = Color.Black;
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox78.Text = CodeAllContent("category_nebespeki", "name_category", "code_category", comboBox10.SelectedItem);
            textBox78.BackColor = Color.White;
            textBox78.ForeColor = Color.Black;
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string[] str = comboBox11.SelectedItem.ToString().Split(';');
                textBox79.Text = CodeAllContent("place_fire", "name_place", "code_place", str[1]);
                textBox79.BackColor = Color.White;
                textBox79.ForeColor = Color.Black;
                if (textBox79.Text == "108" || textBox79.Text == "37")
                {
                    textBox156.Text = "";
                }
                else
                {
                    textBox156.Text = str[1];
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
          
            
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string[] str = comboBox13.SelectedItem.ToString().Split(';');
                textBox80.Text = CodeAllContent("virib_iniciator", "item_virib", "code_virib", str[1]);
                bool res = int.TryParse(textBox80.Text, out int code);
                if (code == 109 || code == 114 || code == 123 || code == 129 || code == 132 || code == 133 || code == 208 || code == 215 || code == 223 || code == 228 || code == 229 || code == 303
                    || code == 409 || code == 513 || code == 610 || code == 705 || code == 804 || code == 806 || code == 904 || code == 1004 || code == 1103 || code == 1210 || code == 1226)
                {
                    textBox157.Text = "";
                }
                else
                {
                    textBox157.Text = str[1];
                }

                textBox80.BackColor = Color.White;
                textBox80.ForeColor = Color.Black;
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
        }

        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string[] str = comboBox24.SelectedItem.ToString().Split(';');
                textBox81.Text = CodeAllContent("pricini_fire", "name_pricini", "code_pricini", str[1]);
                if (textBox81.Text == "8" || textBox81.Text == "33" || textBox81.Text == "43")
                {
                    textBox158.Text = "";
                }
                else
                {
                    textBox158.Text = str[1];
                }

                textBox81.BackColor = Color.White;
                textBox81.ForeColor = Color.Black;
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox10.Text = (comboBox14.SelectedIndex + 1).ToString();
            textBox10.BackColor = Color.White;
            textBox10.ForeColor = Color.Black;
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox83.Text = (comboBox15.SelectedIndex + 1).ToString();
            textBox83.BackColor = Color.White;
            textBox83.ForeColor = Color.Black;
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox84.Text = (comboBox16.SelectedIndex + 1).ToString();
            textBox84.BackColor = Color.White;
            textBox84.ForeColor = Color.Black;
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox85.Text = (comboBox17.SelectedIndex + 1).ToString();
            textBox85.BackColor = Color.White;
            textBox85.ForeColor = Color.Black;
        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox86.Text = (comboBox18.SelectedIndex + 1).ToString();
            textBox86.BackColor = Color.White;
            textBox86.ForeColor = Color.Black;
        }

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] str = comboBox19.SelectedItem.ToString().Split(';');
            textBox87.Text = CodeAllContent("social_status", "name_status", "code_status", str[1]);
            textBox87.BackColor = Color.White;
            textBox87.ForeColor = Color.Black;
        }

        private void comboBox30_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox97.Text = CodeAllContent("umova_smerti", "name_umova", "code_umova", comboBox30.SelectedItem);
            textBox97.BackColor = Color.White;
            textBox97.ForeColor = Color.Black;
        }

        private void comboBox31_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox98.Text = CodeAllContent("umova_smerti", "name_umova", "code_umova", comboBox31.SelectedItem);
            textBox98.BackColor = Color.White;
            textBox98.ForeColor = Color.Black;
        }

        private void comboBox32_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox99.Text = CodeAllContent("umova_smerti", "name_umova", "code_umova", comboBox32.SelectedItem);
            textBox99.BackColor = Color.White;
            textBox99.ForeColor = Color.Black;
        }

        private void comboBox33_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox100.Text = CodeAllContent("umova_smerti", "name_umova", "code_umova", comboBox33.SelectedItem);
            textBox100.BackColor = Color.White;
            textBox100.ForeColor = Color.Black;
        }

        private void comboBox34_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox101.Text = CodeAllContent("umova_smerti", "name_umova", "code_umova", comboBox34.SelectedItem);
            textBox101.BackColor = Color.White;
            textBox101.ForeColor = Color.Black;
        }

        private void comboBox35_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox102.Text = CodeAllContent("info_fire", "name_fire_likvid", "code_fire_likvid", comboBox35.SelectedItem);
            textBox102.BackColor = Color.White;
            textBox102.ForeColor = Color.Black;
        }

        private void comboBox36_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox103.Text = CodeAllContent("poshirenya_fire", "umovi_poshireni", "code_poshireni", comboBox36.SelectedItem);
            textBox103.BackColor = Color.White;
            textBox103.ForeColor = Color.Black;
        }

        private void comboBox37_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox104.Text = CodeAllContent("poshirenya_fire", "umovi_poshireni", "code_poshireni", comboBox37.SelectedItem);
            textBox104.BackColor = Color.White;
            textBox104.ForeColor = Color.Black;
        }

        private void comboBox38_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox105.Text = CodeAllContent("poshirenya_fire", "umovi_poshireni", "code_poshireni", comboBox38.SelectedItem);
            textBox105.BackColor = Color.White;
            textBox105.ForeColor = Color.Black;
        }

        private void comboBox39_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox106.Text = CodeAllContent("poshirenya_fire", "umovi_poshireni", "code_poshireni", comboBox39.SelectedItem);
            textBox106.BackColor = Color.White;
            textBox106.ForeColor = Color.Black;
        }

        private void comboBox40_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox107.Text = CodeAllContent("poshirenya_fire", "umovi_poshireni", "code_poshireni", comboBox40.SelectedItem);
            textBox107.BackColor = Color.White;
            textBox107.ForeColor = Color.Black;
        }

        private void comboBox45_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox112.Text = CodeAllContent("uskladnenya_fire", "name_uskl", "code_uskl", comboBox45.SelectedItem);
            textBox112.BackColor = Color.White;
            textBox112.ForeColor = Color.Black;
        }

        private void comboBox44_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox111.Text = CodeAllContent("uskladnenya_fire", "name_uskl", "code_uskl", comboBox44.SelectedItem);
            textBox111.BackColor = Color.White;
            textBox111.ForeColor = Color.Black;
        }

        private void comboBox43_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox110.Text = CodeAllContent("uskladnenya_fire", "name_uskl", "code_uskl", comboBox43.SelectedItem);
            textBox110.BackColor = Color.White;
            textBox110.ForeColor = Color.Black;
        }

        private void comboBox42_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox109.Text = CodeAllContent("uskladnenya_fire", "name_uskl", "code_uskl", comboBox42.SelectedItem);
            textBox109.BackColor = Color.White;
            textBox109.ForeColor = Color.Black;
        }

        private void comboBox41_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox108.Text = CodeAllContent("uskladnenya_fire", "name_uskl", "code_uskl", comboBox41.SelectedItem);
            textBox108.BackColor = Color.White;
            textBox108.ForeColor = Color.Black;
        }

        private void comboBox47_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox118.Text = CodeAllContent("system_protipojeji", "name_system", "code_system", comboBox47.SelectedItem);
            textBox118.BackColor = Color.White;
            textBox118.ForeColor = Color.Black;
        }

        private void comboBox48_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox117.Text = CodeAllContent("system_protipojeji", "name_system", "code_system", comboBox48.SelectedItem);
            textBox117.BackColor = Color.White;
            textBox117.ForeColor = Color.Black;
        }

        private void comboBox49_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox116.Text = CodeAllContent("system_protipojeji", "name_system", "code_system", comboBox49.SelectedItem);
            textBox116.BackColor = Color.White;
            textBox116.ForeColor = Color.Black;
        }

        private void comboBox50_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox115.Text = CodeAllContent("system_protipojeji", "name_system", "code_system", comboBox50.SelectedItem);
            textBox115.BackColor = Color.White;
            textBox115.ForeColor = Color.Black;
        }

        private void comboBox51_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox114.Text = CodeAllContent("system_protipojeji", "name_system", "code_system", comboBox51.SelectedItem);
            textBox114.BackColor = Color.White;
            textBox114.ForeColor = Color.Black;
        }

        private void comboBox52_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox123.Text = CodeAllContent("resultat_dii_system", "name_res", "code_res", comboBox52.SelectedItem);
            textBox123.BackColor = Color.White;
            textBox123.ForeColor = Color.Black;
        }

        private void comboBox53_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox122.Text = CodeAllContent("resultat_dii_system", "name_res", "code_res", comboBox53.SelectedItem);
            textBox122.BackColor = Color.White;
            textBox122.ForeColor = Color.Black;
        }

        private void comboBox54_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox121.Text = CodeAllContent("resultat_dii_system", "name_res", "code_res", comboBox54.SelectedItem);
            textBox121.BackColor = Color.White;
            textBox121.ForeColor = Color.Black;
        }

        private void comboBox55_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox120.Text = CodeAllContent("resultat_dii_system", "name_res", "code_res", comboBox55.SelectedItem);
            textBox120.BackColor = Color.White;
            textBox120.ForeColor = Color.Black;
        }

        private void comboBox56_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox119.Text = CodeAllContent("resultat_dii_system", "name_res", "code_res", comboBox56.SelectedItem);
            textBox119.BackColor = Color.White;
            textBox119.ForeColor = Color.Black;
        }

        private void comboBox57_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox128.Text = CodeAllContent("uchasnik_fire", "name_uchasnik", "code_uchasnik", comboBox57.SelectedItem);
            textBox128.BackColor = Color.White;
            textBox128.ForeColor = Color.Black;
        }

        private void comboBox58_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox127.Text = CodeAllContent("uchasnik_fire", "name_uchasnik", "code_uchasnik", comboBox58.SelectedItem);
            textBox127.BackColor = Color.White;
            textBox127.ForeColor = Color.Black;
        }

        private void comboBox59_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox126.Text = CodeAllContent("uchasnik_fire", "name_uchasnik", "code_uchasnik", comboBox59.SelectedItem);
            textBox126.BackColor = Color.White;
            textBox126.ForeColor = Color.Black;
        }

        private void comboBox60_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox125.Text = CodeAllContent("uchasnik_fire", "name_uchasnik", "code_uchasnik", comboBox60.SelectedItem);
            textBox125.BackColor = Color.White;
            textBox125.ForeColor = Color.Black;
        }

        private void comboBox61_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox124.Text = CodeAllContent("uchasnik_fire", "name_uchasnik", "code_uchasnik", comboBox61.SelectedItem);
            textBox124.BackColor = Color.White;
            textBox124.ForeColor = Color.Black;
        }

        private void comboBox66_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox129.Text = CodeAllContent("fire_auto", "name_auto", "code_auto", comboBox66.SelectedItem);
            textBox129.BackColor = Color.White;
            textBox129.ForeColor = Color.Black;
        }

        private void comboBox65_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox130.Text = CodeAllContent("fire_auto", "name_auto", "code_auto", comboBox65.SelectedItem);
            textBox130.BackColor = Color.White;
            textBox130.ForeColor = Color.Black;
        }

        private void comboBox64_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox131.Text = CodeAllContent("fire_auto", "name_auto", "code_auto", comboBox64.SelectedItem);
            textBox131.BackColor = Color.White;
            textBox131.ForeColor = Color.Black;
        }

        private void comboBox63_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox132.Text = CodeAllContent("fire_auto", "name_auto", "code_auto", comboBox63.SelectedItem);
            textBox132.BackColor = Color.White;
            textBox132.ForeColor = Color.Black;
        }

        private void comboBox62_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox133.Text = CodeAllContent("fire_auto", "name_auto", "code_auto", comboBox62.SelectedItem);
            textBox133.BackColor = Color.White;
            textBox133.ForeColor = Color.Black;
        }

        private void comboBox67_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox134.Text = CodeAllContent("fire_stvoli", "name_stvol", "code_stvol", comboBox67.SelectedItem);
            textBox134.BackColor = Color.White;
            textBox134.ForeColor = Color.Black;
        }

        private void comboBox68_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox135.Text = CodeAllContent("fire_stvoli", "name_stvol", "code_stvol", comboBox68.SelectedItem);
            textBox135.BackColor = Color.White;
            textBox135.ForeColor = Color.Black;
        }

        private void comboBox69_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox136.Text = CodeAllContent("fire_stvoli", "name_stvol", "code_stvol", comboBox69.SelectedItem);
            textBox136.BackColor = Color.White;
            textBox136.ForeColor = Color.Black;
        }

        private void comboBox72_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox137.Text = CodeAllContent("vognegasni_rechovini", "name_rechovini", "code_rechovini", comboBox72.SelectedItem);
            textBox137.BackColor = Color.White;
            textBox137.ForeColor = Color.Black;
        }

        private void comboBox71_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox138.Text = CodeAllContent("vognegasni_rechovini", "name_rechovini", "code_rechovini", comboBox71.SelectedItem);
            textBox138.BackColor = Color.White;
            textBox138.ForeColor = Color.Black;
        }

        private void comboBox70_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox139.Text = CodeAllContent("vognegasni_rechovini", "name_rechovini", "code_rechovini", comboBox70.SelectedItem);
            textBox139.BackColor = Color.White;
            textBox139.ForeColor = Color.Black;
        }

        private void comboBox75_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox140.Text = CodeAllContent("zacobi_fire", "name_zasobi", "code_zasobi", comboBox75.SelectedItem);
            textBox140.BackColor = Color.White;
            textBox140.ForeColor = Color.Black;
        }

        private void comboBox74_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox141.Text = CodeAllContent("zacobi_fire", "name_zasobi", "code_zasobi", comboBox74.SelectedItem);
            textBox141.BackColor = Color.White;
            textBox141.ForeColor = Color.Black;
        }

        private void comboBox73_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox142.Text = CodeAllContent("zacobi_fire", "name_zasobi", "code_zasobi", comboBox73.SelectedItem);
            textBox142.BackColor = Color.White;
            textBox142.ForeColor = Color.Black;
        }

        private void comboBox78_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox143.Text = CodeAllContent("djerela_vodopostachanaya", "name_djerela", "code_djerela", comboBox78.SelectedItem);
            textBox143.BackColor = Color.White;
            textBox143.ForeColor = Color.Black;
        }

        private void comboBox77_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox144.Text = CodeAllContent("djerela_vodopostachanaya", "name_djerela", "code_djerela", comboBox77.SelectedItem);
            textBox144.BackColor = Color.White;
            textBox144.ForeColor = Color.Black;
        }

        private void comboBox76_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox145.Text = CodeAllContent("djerela_vodopostachanaya", "name_djerela", "code_djerela", comboBox76.SelectedItem);
            textBox145.BackColor = Color.White;
            textBox145.ForeColor = Color.Black;
        }

        private void comboBox79_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox146.Text = CodeAllContent("perevirka_fire", "name_perevirka", "code_perevirka", comboBox79.SelectedItem);
            textBox146.BackColor = Color.White;
            textBox146.ForeColor = Color.Black;
        }

        private void comboBox80_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox147.Text = CodeAllContent("gospodar_dialnist", "name_dialnist", "code_dialnist", comboBox80.SelectedItem);
            textBox147.BackColor = Color.White;
            textBox147.ForeColor = Color.Black;
        }

        private void comboBox81_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox148.Text = CodeAllContent("zahodi_pojeji", "name_zahodi", "code_zahodi", comboBox81.SelectedItem);
            textBox148.BackColor = Color.White;
            textBox148.ForeColor = Color.Black;
        }

        private void comboBox82_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox149.Text = CodeAllContent("zahodi_pojeji", "name_zahodi", "code_zahodi", comboBox82.SelectedItem);
            textBox149.BackColor = Color.White;
            textBox149.ForeColor = Color.Black;
        }

        private void textBox72_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox72);
            bool state = false;
            int code = 0;
            if (textBox72.Text == "")
            {
                textBox73.Text = "";
            }
            state = int.TryParse(textBox72.Text, out int code1);
            if (state)
            {
                code = int.Parse(textBox72.Text);
                if(code== 1904|| code == 1818 || code == 1815 || code == 1519 || code == 1404 || code == 1309|| code == 1216 || code ==1114|| code==1003|| code==911|| code==804 || code==712|| code==606 || code == 512 || code == 408 
                   || code == 310 || code == 308 || code == 209 || code == 117 || code == 16)
                {
                    textBox73.Text = "";
                }
                else
                {
                    textBox73.Text = CodeAllContent("fire_objects", "fire_code", "fire_item", textBox72.Text);
                }
            }
            bool val_cur = int.TryParse(textBox72.Text, out int res);
            if ((res >= 1601 && res < 1606) || (res >= 1701 && res < 1708) || (res >= 1801 && res < 1819))
            {
                if (textBox67.Text != "")
                {
                    //MessageBox.Show("Це поле повинне бути пустим!");
                    textBox67.BackColor = Color.Firebrick;
                    textBox67.ForeColor = Color.White;
                    
                }
                if (textBox68.Text != "")
                {
                    textBox68.BackColor = Color.Firebrick;
                    textBox68.ForeColor = Color.White;
                   
                }
                if (textBox77.Text != "")
                {
                    textBox77.BackColor = Color.Firebrick;
                    textBox77.ForeColor = Color.White;
                    
                }
                if (textBox78.Text != "")
                {
                    textBox78.BackColor = Color.Firebrick;
                    textBox78.ForeColor = Color.White;
                   
                }
            }
            if(res>=1 && res < 210)
            {
                if (textBox78.Text == "")
                {
                    textBox78.BackColor = Color.Firebrick;
                }
            }
            else
            {
                textBox78.BackColor = Color.White;
            }



        }

        private void textBox70_TextChanged(object sender, EventArgs e)
        {
            bool state=false;
            int code=0;
            
            state =int.TryParse(textBox70.Text,  out int code1);
            if (state)
            {
                code = int.Parse(textBox70.Text);
                if (code > 0 && code < 7)
                {
                    if (textBox70.Text == "1")
                    {
                        textBox4.Text = "м.";
                    }
                    else if (textBox70.Text == "2")
                    {
                        textBox4.Text = "c.м.т";
                    }
                    else if (textBox70.Text == "3")
                    {
                        textBox4.Text = "c.";
                    }
                    else if (textBox70.Text == "4")
                    {
                        textBox4.Text = "п.м. міського н.п";
                    }
                    else if (textBox70.Text == "5")
                    {
                        textBox4.Text = "п.м. сільського н.п.";
                    }
                    else if (textBox70.Text == "6")
                    {
                        textBox4.Text = "відселена зона";
                    }
                    else
                    {
                        textBox4.Text = "";
                    }
                }
                else
                {
                    MessageBox.Show("В це поле можно внести тільки коди від 1 до 6", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox70.Text = "";
                }
            }
            else
            {
                if (textBox70.Text != "")
                {
                    MessageBox.Show("В це поле можно внести тільки цифри", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox70.Text = "";
                }
               
            }
            
           
        }

        private void textBox69_TextChanged(object sender, EventArgs e)
        {
            bool state = false;
            int code = 0;

            state = int.TryParse(textBox69.Text, out int code1);
            if (state)
            {
                code = int.Parse(textBox69.Text);
                if(code> 27 || code<=0)
                {
                    MessageBox.Show("В це поле можно внести тільки коди від 1 до 27", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox69.Text = "";
                }
            }
            else
            {
                if (textBox69.Text != "")
                {
                    MessageBox.Show("В це поле можно внести тільки цифри", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox69.Text = "";
                }
                    
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker1.MaxDate = dateTimePicker10.Value;
            maskedTextBox1.Text = dateTimePicker1.Value.ToString();
        }

        private void textBox71_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox71);
        }

        private void textBox72_Leave(object sender, EventArgs e)
        {
           
            if (!(comboBox3.Focused || comboBox4.Focused))
            {
                string code = CodeAllContent("fire_objects", "fire_code", "fire_item", textBox72.Text);
                if (code == "")
                {
                    MessageBox.Show("Такого кода не існує", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox73.Text = "";
                    textBox72.Text = "";
                    textBox72.Select();
                    textBox72.ScrollToCaret();
                }
            }
           
            

        }

        private void textBox74_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox74, 6, 1, 6);
        }

        private void textBox75_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox75);
        }

        private void textBox75_Leave(object sender, EventArgs e)
        {
            string code = CodeAllContent("stupin_riziku", "code_riziku", "name_riziku", textBox75.Text);
            if(textBox75.Text!="")
            if (code == "")
            {
                MessageBox.Show("В це поле можно внести тільки коди 10,20,31,32", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox75.Text = "";
                textBox75.BackColor = Color.Firebrick;
                    textBox75.Select();
                    textBox75.ScrollToCaret();
            }
        }

        private void textBox76_Leave(object sender, EventArgs e)
        {
            string code = CodeAllContent("pidkontrol_object", "code_object", "name_object", textBox76.Text);
            if (textBox76.Text != "")
            {
                if (code == "")
                {
                    MessageBox.Show("В це поле можно внести тільки коди 11,12,13,20", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox76.Text = "";
                    textBox76.BackColor = Color.Firebrick;
                    textBox76.Select();
                    textBox76.ScrollToCaret();
                }
            }else if (textBox75.Text != "")
            {
                if (textBox76.Text == "" && comboBox7.Focused==false)
                {
                    MessageBox.Show("Це поле не повинне бути пустим!");
                    textBox76.BackColor = Color.Firebrick;
                    textBox76.Select();
                    textBox76.ScrollToCaret();
                }
            }
        }

        private void textBox77_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox77);
            if (textBox77.Text == "0" || textBox77.Text == "9")
            {
                MessageBox.Show("В це поле можно внести тільки коди від 1 до 8", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox77.Text = "";
            }
            bool res = int.TryParse(textBox72.Text, out int val);
            if ((val >= 1601 && val < 1606) || (val >= 1701 && val < 1708) || (val >= 1801 && val < 1819))
            {
                if (textBox77.Text != "")
                {
                    MessageBox.Show("Це поле повинне бути пустим!");
                    textBox77.Text = "";
                }
            }

        }

        private void textBox78_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox78, 10, 1, 10);
            bool res = int.TryParse(textBox72.Text, out int val);
            if ((val >= 1601 && val < 1606) || (val >= 1701 && val < 1708) || (val >= 1801 && val < 1819))
            {
                if (textBox78.Text != "")
                {
                    MessageBox.Show("Це поле повинне бути пустим!");
                    textBox78.Text = "";
                }
            }
        }

        private void textBox79_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox79, 108, 1, 108);
           
            if (textBox79.Text == "108" || textBox79.Text == "37")
            {
                textBox156.Text = "";
            }
            else
            {
                textBox156.Text = CodeAllContent("place_fire", "code_place", "name_place", textBox79.Text);
            }
        }

        private void textBox80_Leave(object sender, EventArgs e)
        {
            if(!(comboBox12.Focused || comboBox13.Focused))
            {
                string code = CodeAllContent("virib_iniciator", "code_virib", "item_virib", textBox80.Text);
                //if (textBox80.Text != "")
                if (code == "")
                {
                    MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox80.Text = "";
                    textBox80.BackColor = Color.Firebrick;
                    textBox80.Select();
                    textBox80.ScrollToCaret();
                }
            }
        }

        private void textBox81_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox81, 43, 1, 43);
            if (textBox81.Text == "")
            {
                textBox158.Text = "";
            }
            else
            {
                if (textBox81.Text == "8" || textBox81.Text == "33" || textBox81.Text == "43")
                {
                    textBox158.Text = "";
                }
                else
                {
                    textBox158.Text = CodeAllContent("pricini_fire", "code_pricini", "name_pricini", textBox81.Text);
                }
                
            }
        }

        private void textBox5_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox5);
        }

        private void textBox6_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox6);
        }

        private void textBox8_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox8);
            if (textBox8.Text == "")
            {
                textBox18.BackColor = Color.White;
                textBox17.BackColor = Color.White;
                textBox10.BackColor = Color.White;
                textBox87.BackColor = Color.White;
                textBox92.BackColor = Color.White;
                textBox97.BackColor = Color.White;

                textBox20.BackColor = Color.White;
                textBox19.BackColor = Color.White;
                textBox83.BackColor = Color.White;
                textBox88.BackColor = Color.White;
                textBox93.BackColor = Color.White;
                textBox98.BackColor = Color.White;

                textBox22.BackColor = Color.White;
                textBox21.BackColor = Color.White;
                textBox84.BackColor = Color.White;
                textBox89.BackColor = Color.White;
                textBox94.BackColor = Color.White;
                textBox99.BackColor = Color.White;

                textBox24.BackColor = Color.White;
                textBox23.BackColor = Color.White;
                textBox85.BackColor = Color.White;
                textBox90.BackColor = Color.White;
                textBox95.BackColor = Color.White;
                textBox100.BackColor = Color.White;

                textBox26.BackColor = Color.White;
                textBox25.BackColor = Color.White;
                textBox86.BackColor = Color.White;
                textBox91.BackColor = Color.White;
                textBox96.BackColor = Color.White;
                textBox101.BackColor = Color.White;
            }
            if (textBox8.Text == "1")
            {
                textBox18.BackColor = Color.Aqua;
                textBox17.BackColor = Color.Aqua;
                textBox10.BackColor = Color.Aqua;
                textBox87.BackColor = Color.Aqua;
                textBox92.BackColor = Color.Aqua;
                textBox97.BackColor = Color.Aqua;
            }
           
            if (textBox8.Text == "2")
            {
                textBox18.BackColor = Color.Aqua;
                textBox17.BackColor = Color.Aqua;
                textBox10.BackColor = Color.Aqua;
                textBox87.BackColor = Color.Aqua;
                textBox92.BackColor = Color.Aqua;
                textBox97.BackColor = Color.Aqua;

                textBox20.BackColor = Color.Aqua;
                textBox19.BackColor = Color.Aqua;
                textBox83.BackColor = Color.Aqua;
                textBox88.BackColor = Color.Aqua;
                textBox93.BackColor = Color.Aqua;
                textBox98.BackColor = Color.Aqua;
            }
           
            if (textBox8.Text == "3")
            {
                textBox18.BackColor = Color.Aqua;
                textBox17.BackColor = Color.Aqua;
                textBox10.BackColor = Color.Aqua;
                textBox87.BackColor = Color.Aqua;
                textBox92.BackColor = Color.Aqua;
                textBox97.BackColor = Color.Aqua;

                textBox20.BackColor = Color.Aqua;
                textBox19.BackColor = Color.Aqua;
                textBox83.BackColor = Color.Aqua;
                textBox88.BackColor = Color.Aqua;
                textBox93.BackColor = Color.Aqua;
                textBox98.BackColor = Color.Aqua;

                textBox22.BackColor = Color.Aqua;
                textBox21.BackColor = Color.Aqua;
                textBox84.BackColor = Color.Aqua;
                textBox89.BackColor = Color.Aqua;
                textBox94.BackColor = Color.Aqua;
                textBox99.BackColor = Color.Aqua;
            }
            if (textBox8.Text == "4")
            {
                textBox18.BackColor = Color.Aqua;
                textBox17.BackColor = Color.Aqua;
                textBox10.BackColor = Color.Aqua;
                textBox87.BackColor = Color.Aqua;
                textBox92.BackColor = Color.Aqua;
                textBox97.BackColor = Color.Aqua;

                textBox20.BackColor = Color.Aqua;
                textBox19.BackColor = Color.Aqua;
                textBox83.BackColor = Color.Aqua;
                textBox88.BackColor = Color.Aqua;
                textBox93.BackColor = Color.Aqua;
                textBox98.BackColor = Color.Aqua;

                textBox22.BackColor = Color.Aqua;
                textBox21.BackColor = Color.Aqua;
                textBox84.BackColor = Color.Aqua;
                textBox89.BackColor = Color.Aqua;
                textBox94.BackColor = Color.Aqua;
                textBox99.BackColor = Color.Aqua;

                textBox24.BackColor = Color.Aqua;
                textBox23.BackColor = Color.Aqua;
                textBox85.BackColor = Color.Aqua;
                textBox90.BackColor = Color.Aqua;
                textBox95.BackColor = Color.Aqua;
                textBox100.BackColor = Color.Aqua;
            }
            if (textBox8.Text == "5")
            {
                textBox18.BackColor = Color.Aqua;
                textBox17.BackColor = Color.Aqua;
                textBox10.BackColor = Color.Aqua;
                textBox87.BackColor = Color.Aqua;
                textBox92.BackColor = Color.Aqua;
                textBox97.BackColor = Color.Aqua;

                textBox20.BackColor = Color.Aqua;
                textBox19.BackColor = Color.Aqua;
                textBox83.BackColor = Color.Aqua;
                textBox88.BackColor = Color.Aqua;
                textBox93.BackColor = Color.Aqua;
                textBox98.BackColor = Color.Aqua;

                textBox22.BackColor = Color.Aqua;
                textBox21.BackColor = Color.Aqua;
                textBox84.BackColor = Color.Aqua;
                textBox89.BackColor = Color.Aqua;
                textBox94.BackColor = Color.Aqua;
                textBox99.BackColor = Color.Aqua;

                textBox24.BackColor = Color.Aqua;
                textBox23.BackColor = Color.Aqua;
                textBox85.BackColor = Color.Aqua;
                textBox90.BackColor = Color.Aqua;
                textBox95.BackColor = Color.Aqua;
                textBox100.BackColor = Color.Aqua;

                textBox26.BackColor = Color.Aqua;
                textBox25.BackColor = Color.Aqua;
                textBox86.BackColor = Color.Aqua;
                textBox91.BackColor = Color.Aqua;
                textBox96.BackColor = Color.Aqua;
                textBox101.BackColor = Color.Aqua;
            }
        }

        private void textBox65_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox65);
        }

        private void textBox82_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox82);
        }


        private void textBox10_TextChanged_1(object sender, EventArgs e)
        {
            validFieldContent(textBox10, 3, 1, 3);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox10.Text = "";
            }
        }

        private void textBox83_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox83, 3, 1, 3);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox83.Text = "";
            }
        }

        private void textBox84_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox84, 3, 1, 3);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox84.Text = "";
            }
        }

        private void textBox85_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox85, 3, 1, 3);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox85.Text = "";
            }
        }

        private void textBox86_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox86, 3, 1, 3);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox86.Text = "";
            }

        }

        private void textBox87_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox87, 13, 1, 13);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox87.Text = "";
            }
        }
        public void validFieldContent(TextBox textbox, int code_count, int first, int last)
        {
            validateIntForm(textbox);
            bool state = false;
            int code;
            string message = "В це поле можно внести тільки коди від " + first + " до " + last;
            state = int.TryParse(textbox.Text, out int code1);
            if (state)
            {
                code = int.Parse(textbox.Text);
                if (code > code_count)
                {
                    MessageBox.Show(message, "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textbox.Text = "";
                }
            }
            if (textbox.Text == "0")
            {
                MessageBox.Show(message, "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textbox.Text = "";
            }
        }

        private void textBox88_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox88, 13, 1, 13);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox88.Text = "";
            }
        }

        private void textBox89_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox89, 13, 1, 13);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox89.Text = "";
            }
        }

        private void textBox90_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox90, 13, 1, 13);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox90.Text = "";
            }
        }

        private void textBox91_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox91, 13, 1, 13);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox91.Text = "";
            }
        }

        private void textBox92_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox92, 3, 1, 3);
            if (textBox92.Text == "1" || textBox92.Text == "2")
            {
                 if(textBox5.Text=="")
                textBox5.BackColor = Color.Aqua;
            }
            else
            {
                textBox5.BackColor = Color.White;
            }
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox92.Text = "";
            }
        }

        private void textBox93_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox93, 3, 1, 3);
           
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox93.Text = "";
            }
            
        }

        private void textBox94_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox94, 3, 1, 3);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox94.Text = "";
            }
        }

        private void textBox95_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox95, 3, 1, 3);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox95.Text = "";
            }
        }

        private void textBox96_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox96, 3, 1, 3);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox96.Text = "";
            }
        }

        private void textBox97_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox97, 17, 1, 17);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox97.Text = "";
            }
        }

        private void textBox98_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox98, 17, 1, 17);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox98.Text = "";
            }
        }

        private void textBox99_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox99, 17, 1, 17);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox99.Text = "";
            }
        }

        private void textBox100_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox100, 17, 1, 17);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox100.Text = "";
            }
        }

        private void textBox101_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox101, 17, 1, 17);
            if (textBox8.Text == "" && textBox65.Text == "" && textBox82.Text == "")
            {
                textBox101.Text = "";
            }
        }

        private void textBox12_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox12);
        }

        private void textBox11_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox11);
           
          
        }

        private void textBox66_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox66);
        }

        private void textBox9_TextChanged_1(object sender, EventArgs e)
        {
            validateFloatForm(textBox9);
        }

        private void textBox15_TextChanged_1(object sender, EventArgs e)
        {
            validateFloatForm(textBox15);
        }

        private void textBox7_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox7);
        }

        private void textBox14_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox14);
        }

        private void textBox13_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox13);
        }

        private void textBox16_TextChanged_1(object sender, EventArgs e)
        {
            validateIntForm(textBox16);
        }

        private void textBox102_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox102, 3, 1, 3);
            if (textBox102.Text == "1")
            {
                if (textBox128.Text == "")
                {
                    textBox128.BackColor = Color.Firebrick;
                }
                if (textBox129.Text == "")
                {
                    textBox129.BackColor = Color.Firebrick;
                }
                if (textBox58.Text == "")
                {
                    textBox58.BackColor = Color.Firebrick;
                }
                if (textBox49.Text == "")
                {
                    textBox49.BackColor = Color.Firebrick;
                }
                //textBox49.BackColor = Color.White;
                textBox137.BackColor = Color.White;

            }
            if (textBox102.Text == "2")
            {
                if (textBox128.Text == "")
                {
                    textBox128.BackColor = Color.Firebrick;
                }
                if (textBox129.Text == "")
                {
                    textBox129.BackColor = Color.Firebrick;
                }
                if (textBox58.Text == "")
                {
                    textBox58.BackColor = Color.Firebrick;
                }
                if (textBox49.Text == "")
                {
                    textBox49.BackColor = Color.Firebrick;
                }
                if (textBox137.Text == "")
                {
                    textBox137.BackColor = Color.Firebrick;
                }
                
            }
            if (textBox102.Text == "3")
            {
                textBox128.BackColor = Color.White;
                textBox129.BackColor = Color.White;
                textBox58.BackColor = Color.White;
                textBox49.BackColor = Color.White;
                textBox137.BackColor = Color.White;
            }
        }

        private void textBox103_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox103, 15, 1, 15);
        }

        private void textBox104_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox104, 15, 1, 15);
        }

        private void textBox105_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox105, 15, 1, 15);
        }

        private void textBox106_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox106, 15, 1, 15);
        }

        private void textBox107_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox107, 15, 1, 15);
        }

        private void textBox112_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox112, 13, 1, 13);
        }

        private void textBox111_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox111, 13, 1, 13);
        }

        private void textBox110_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox110, 13, 1, 13);
        }

        private void textBox109_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox109, 13, 1, 13);
        }

        private void textBox108_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox108, 13, 1, 13);
        }

        private void textBox113_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox113, 3, 1, 3);
            if (textBox113.Text == "1")
            {
                panel1.Visible = true;
            }
            else
            {
                panel1.Visible = false;
            }
        }

        private void textBox118_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox118, 5, 1, 5);
        }

        private void textBox117_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox117, 5, 1, 5);
        }

        private void textBox116_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox116, 5, 1, 5);
        }

        private void textBox115_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox115, 5, 1, 5);
        }

        private void textBox114_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox114, 5, 1, 5);
        }

        private void textBox123_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox123, 13, 1, 13);
        }

        private void textBox122_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox122, 13, 1, 13);
        }

        private void textBox121_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox121, 13, 1, 13);
        }

        private void textBox120_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox120, 13, 1, 13);
        }

        private void textBox119_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox119, 13, 1, 13);
        }

        private void textBox128_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox128, 5, 1, 5);
            bool val50 = int.TryParse(textBox102.Text, out int res50);
            if (res50 == 3)
            {
                if (textBox128.Text == "1")
                {
                   // MessageBox.Show("Ви не можете записати код 1 в це поле");
                    textBox128.Text = "";
                }
            }
        }

        private void textBox127_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox127, 5, 1, 5);
            bool val50 = int.TryParse(textBox102.Text, out int res50);
            if (res50 == 3)
            {
                if (textBox127.Text == "1")
                {
                   // MessageBox.Show("Ви не можете записати код 1 в це поле");
                    textBox127.Text = "";
                }
            }
        }

        private void textBox126_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox126, 5, 1, 5);
            bool val50 = int.TryParse(textBox102.Text, out int res50);
            if (res50 == 3)
            {
                if (textBox126.Text == "1")
                {
                   // MessageBox.Show("Ви не можете записати код 1 в це поле");
                    textBox126.Text = "";
                }
            }
        }

        private void textBox125_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox125, 5, 1, 5);
            bool val50 = int.TryParse(textBox102.Text, out int res50);
            if (res50 == 3)
            {
                if (textBox125.Text == "1")
                {
                   // MessageBox.Show("Ви не можете записати код 1 в це поле");
                    textBox125.Text = "";
                }
            }
        }

        private void textBox124_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox124, 5, 1, 5);
            bool val50 = int.TryParse(textBox102.Text, out int res50);
            if (res50 == 3)
            {
                if (textBox124.Text == "1")
                {
                   // MessageBox.Show("Ви не можете записати код 1 в це поле");
                    textBox124.Text = "";
                }
            }
        }

        private void textBox129_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox129, 36, 1, 36);
        }

        private void textBox130_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox130, 36, 1, 36);
        }

        private void textBox131_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox131, 36, 1, 36);
        }

        private void textBox132_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox132, 36, 1, 36);
        }

        private void textBox133_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox133, 36, 1, 36);
        }

        private void textBox134_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox134, 11, 1, 11);
        }

        private void textBox135_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox135, 11, 1, 11);
        }

        private void textBox136_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox136, 11, 1, 11);
        }

        private void textBox137_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox137, 13, 1, 13);
        }

        private void textBox138_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox138, 13, 1, 13);
        }

        private void textBox139_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox139, 13, 1, 13);
        }

        private void textBox140_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox140, 11, 1, 11);
        }

        private void textBox141_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox141, 11, 1, 11);
        }

        private void textBox142_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox142, 11, 1, 11);
        }

        private void textBox143_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox143, 7, 1, 7);
        }

        private void textBox144_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox144, 7, 1, 7);
        }

        private void textBox145_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox145, 7, 1, 7);
        }

        private void textBox146_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox146, 2, 1, 2);
            if (textBox146.Text != "")
            {
                dateTimePicker9.Enabled = true;
                maskedTextBox5.Enabled = true;
            }
            else
            {
                dateTimePicker9.Enabled = false;
                maskedTextBox5.Enabled = false;
            }
        }

        private void textBox147_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox147, 5, 1, 5);
        }

        private void textBox148_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox148, 3, 1, 3);
        }

        private void textBox149_TextChanged(object sender, EventArgs e)
        {
            validateIntForm(textBox149);
           // validFieldContent(textBox149, 6, 3, 6);
            bool res = int.TryParse(textBox149.Text, out int result);
            if (res)
            {
                if(result>3 && result < 7)
                {
                    textBox159.Text = CodeAllContent("zahodi_pojeji", "code_zahodi", "code_kku", result);
                }
                else
                {
                    textBox149.Text = "";
                    MessageBox.Show("В це поле можно ввести тільки цифри від 4 до 6");
                }
            }
        }

        private void comboBox84_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox151.Text = CodeAllContent("region", "name_region", "code_region", comboBox84.SelectedItem);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            
            string code_raion = "";
            string name_raion = "";
             try
             {
            if (textBox152.Text != "" || textBox153.Text != "")
                {
                    if (state_list)
                    {
                        code_raion = textBox152.Text;
                        name_raion = textBox153.Text;
                   
                   // connection();
                        string com = "Update 'current_raion' set  code_raion='" + code_raion + "',name_raion='" + name_raion +  "' where id_raion='" + id_raion + "'";
                   
                    // com = "Update 'current_raion' set  code_raion="+ int.Parse(words[0])+ ",name_raion='" + words[1]+"' where id_raion='"+ id + "'";
                    cmd_db = new SQLiteCommand(com, con_db);
                        cmd_db.ExecuteNonQuery();
                     //   Disconnect();

                        listBox1.Items.Insert(listBox1.SelectedIndex, textBox152.Text + ";" + textBox153.Text);
                         listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                        state_list = false;
                    }
                    else
                    {
                        code_raion = textBox152.Text;
                        name_raion = textBox153.Text;
                    //    connection();
                       string  com = "INSERT INTO 'current_raion' (code_raion,'name_raion') VALUES (" +code_raion + "," + "'" + name_raion + "'" + ")";
                        cmd_db = new SQLiteCommand(com, con_db);
                        cmd_db.ExecuteNonQuery();
                   //     Disconnect();

                        listBox1.Items.Add(textBox152.Text + ";" + textBox153.Text);
                    }
                    
                    textBox152.Text = "";
                    textBox153.Text = "";
               }
                

              
            }
            catch (SQLiteException)
            {
                MessageBox.Show("Всі коди мають бути унікальними");
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            state_list = true;
            try
            {
                string str = listBox1.SelectedItem.ToString();
                String[] words = str.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length > 1)
                {
                    textBox152.Text = words[0];
                    id_raion = CodeAllContent("current_raion", "code_raion", "id_raion", textBox152.Text);
                 
                    textBox153.Text = words[1];
                    //listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                }
                else
                {
                    textBox152.Text = words[0];
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string code_np = "";
            string name_np = "";
            try
            {
                if (textBox155.Text != "" || textBox154.Text != "")
                {
                    if (state_list2)
                    {
                        code_np = textBox155.Text;
                        name_np = textBox154.Text;
                    //    connection();
                        string com = "Update 'current_np' set  code_np='" + code_np + "',name_np='" + name_np + "' where id_np='" + id_np + "'";

                        // com = "Update 'current_raion' set  code_raion="+ int.Parse(words[0])+ ",name_raion='" + words[1]+"' where id_raion='"+ id + "'";
                        cmd_db = new SQLiteCommand(com, con_db);
                        cmd_db.ExecuteNonQuery();
                    //    Disconnect();

                        listBox2.Items.Insert(listBox2.SelectedIndex, textBox155.Text + ";" + textBox154.Text);
                        listBox2.Items.RemoveAt(listBox2.SelectedIndex);
                        state_list2 = false;
                    }
                    else
                    {
                        code_np = textBox155.Text;
                        name_np = textBox154.Text;
                    //    connection();
                        string com = "INSERT INTO 'current_np' (code_np,'name_np') VALUES (" +code_np + "," + "'" + name_np + "'" + ")";
                        cmd_db = new SQLiteCommand(com, con_db);
                        cmd_db.ExecuteNonQuery();
                     //   Disconnect();
                        listBox2.Items.Add(textBox155.Text + ";" + textBox154.Text);
                    }
                    textBox155.Text = "";
                    textBox154.Text = "";
                }
            }
            catch (SQLiteException)
            {
                MessageBox.Show("Всі коди мають бути унікальними");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            state_list2 = true;
            try
            {
                string str = listBox2.SelectedItem.ToString();
                String[] words = str.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length > 1)
                {
                    
                    textBox155.Text = words[0];
                    id_np = CodeAllContent("current_np", "code_np", "id_np", textBox155.Text);
                   
                    textBox154.Text = words[1];
                    //listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                }
                else
                {
                    textBox155.Text = words[0];
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button1.Visible = true;
            string com = "";
            if (textBox151.Text=="" || listBox1.Items.Count==0 || listBox2.Items.Count == 0)
            {
                MessageBox.Show( "Заповніть всі поля.", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                // connection();
                try
                {
                    if (con_db.State == ConnectionState.Open)
                    {
                        com = "DELETE FROM 'current_region'";
                        cmd_db = new SQLiteCommand(com, con_db);
                        cmd_db.ExecuteNonQuery();

                        com = "INSERT INTO 'current_region' ('id_current','code_current_region') VALUES (1," + textBox151.Text + ")";
                        cmd_db = new SQLiteCommand(com, con_db);
                        cmd_db.ExecuteNonQuery();
                        /* for (int i = 0; i < listBox1.Items.Count; i++)
                         {
                             string s = listBox1.Items[i].ToString();
                             String[] words = s.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                             if (words.Length > 1)
                             {
                                 if(words[0]!= CodeAllContent("current_raion", "code_raion", "code_raion", words[0]))
                                 {
                                     connection();
                                     com = "INSERT INTO 'current_raion' (code_raion,'name_raion') VALUES (" + words[0] + "," + "'" + words[1] + "'" + ")";
                                     cmd_db = new SQLiteCommand(com, con_db);
                                     cmd_db.ExecuteNonQuery();
                                     Disconnect();
                                 }

                             }
                             else
                             {
                                 MessageBox.Show("Введить назву района");
                             }
                         }*/


                        /* for (int j = 0; j < listBox2.Items.Count; j++)
                         {
                             string s = listBox2.Items[j].ToString();
                             String[] words = s.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                             if (words.Length > 1)
                             {
                                 if(words[0] != CodeAllContent("current_np", "code_np", "code_np", words[0]))
                                 {
                                     connection();
                                     com = "INSERT INTO 'current_np' (code_np,'name_np') VALUES (" + words[0] + "," + "'" + words[1] + "'" + ")";
                                     cmd_db = new SQLiteCommand(com, con_db);
                                     cmd_db.ExecuteNonQuery();
                                     Disconnect();
                                 }


                             }
                             else
                             {
                                 MessageBox.Show("Введить назву населеного пункту");
                             }
                         }*/
                    }
                    MessageBox.Show("Ви внесли дані в базу");
                    ReadDb();
                    textBox69.Text = CodeAllContent("current_region", "id_current", "code_current_region", "1");
                    panel2.Visible = false;
                    настройкаToolStripMenuItem.Enabled = true;
                }
                catch (SQLiteException ex)
                {
                    if (ex.Message != "")
                    {
                        MessageBox.Show("Коди в таблиці не повинні повторюватись");
                    }
                  //  Console.WriteLine("test " + ex.Message);
                }
                //Disconnect();
            }
           
        }

        private void textBox151_TextChanged(object sender, EventArgs e)
        {
            validFieldContent(textBox151, 27, 1, 27);
        }

        private void comboBox83_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = CodeAllContent("current_raion", "name_raion", "code_raion", comboBox83.SelectedItem);
            textBox1.BackColor = Color.White;
            textBox1.ForeColor = Color.Black;
        }

        private void comboBox85_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox71.Text= CodeAllContent("current_np", "name_np", "code_np", comboBox85.SelectedItem);
            textBox71.BackColor = Color.White;
            textBox71.ForeColor = Color.Black;
        }

        private void textBox71_Leave(object sender, EventArgs e)
        {
            if (!comboBox85.Focused)
            {
                string code = CodeAllContent("current_np", "code_np", "code_np", textBox71.Text);
                // if (textBox71.Text != "")
                if (code == "")
                {
                    MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox71.Text = "";
                    textBox71.BackColor = Color.Firebrick;
                    textBox71.Select();
                    textBox71.ScrollToCaret();
                }
            }
            
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
             DialogResult dialog= MessageBox.Show("Ви впевнені, що бажаєте видалити дані з цієї таблиці?","Увага!",MessageBoxButtons.YesNo,MessageBoxIcon.Warning);
             if (dialog==DialogResult.Yes)
             {
                 string com = "";
                 try
                 {
                    // connection();
                     com = "DELETE FROM 'current_raion'";
                     cmd_db = new SQLiteCommand(com, con_db);
                     cmd_db.ExecuteNonQuery();
                   //  Disconnect();
                     listBox1.Items.Clear();
                 }
                 catch (Exception ex)
                 {

                     Console.WriteLine(ex.Message);
                 }

             }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialog = MessageBox.Show("Ви впевнені, що бажаєте видалити дані з цієї таблиці?", "Увага!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialog == DialogResult.Yes)
            {
                string com = "";
                try
                {
                 //   connection();
                    com = "DELETE FROM 'current_np'";
                    cmd_db = new SQLiteCommand(com, con_db);
                    cmd_db.ExecuteNonQuery();
                 //   Disconnect();
                    listBox2.Items.Clear();
                }
                catch (Exception ex)
                {

                    Console.WriteLine(ex.Message);
                }

            }
        }

        private void кодРегіонуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel2.Visible = true;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            button7.Visible = true;
            ReadDb();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if(listBox1.Items.Count>0 && listBox2.Items.Count > 0)
            {
                panel2.Visible = false;
            }
            else
            {
                MessageBox.Show("Заповніть всі необхідні поля");
            }
            button1.Visible = true;

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Disconnect();
        }


        private void textBox70_Leave(object sender, EventArgs e)
        {
           
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
           
           /* string content_val = "";
            
            string select_code = "SELECT `number_cartka` FROM kartka_obliku WHERE code_raion=" + textBox1.Text + " AND number_cartka=" + textBox2.Text + " AND main_dop=" + textBox3.Text;
            //  connection();
            try
            {
                if (con_db.State == ConnectionState.Open)
                {
                    cmd_db = new SQLiteCommand(select_code, con_db);
                    rdr = cmd_db.ExecuteReader();

                    while (rdr.Read())
                    {
                        content_val = rdr[0].ToString();
                        // dict.Add(code, rdr[0].ToString());

                    }
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }*/
            if (textBox2.Text == "")
            {
                MessageBox.Show("Це поле обов'язкове для заповнення!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox2.Text = "";
                textBox2.BackColor = Color.Firebrick;
                textBox2.Select();
                textBox2.ScrollToCaret();

            }
           /* if (content_val != "")
            {
                textBox2.BackColor = Color.Firebrick;
                textBox2.ForeColor = Color.White;
                MessageBox.Show("Картка під таким номером вже існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox2.Select();
                textBox2.ScrollToCaret();
            }*/

        }

        private void textBox79_Leave(object sender, EventArgs e)
        {
           
        }

        private void створитиPogStatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
            panel4.Visible = false;
            panel2.Visible = false;
            panel5.Visible = false;
            //ReadDb();
            cmd_db = new SQLiteCommand("SELECT `number_cartka`,`main_dop`,`date_viniknenya` from kartka_obliku", con_db);
            rdr = cmd_db.ExecuteReader();
            checkedListBox1.Items.Clear();

            while (rdr.Read())
            {
                // region_items.Add(rdr[1].ToString());
                checkedListBox1.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            button1.Visible = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string select_code="";
            string[] rep;
            string path = Environment.CurrentDirectory + "/export/POG_STAT_" + textBox69.Text + ".txt";
            if (File.Exists(path))
            {
                File.Delete(path);

            }
            string str = "КОД_РЕГ|НАЗВА_РЕГ|КОД_РАЙОНУ|НАЗВА_РАОЙНУ|ТИП_НП|№_КАРТКИ|ОСН_ДОД|ДАТА_ПОЖ|КОД_НП|АДРЕСА|КОД_ОБ|НАЗВА_ ОБ|КОД_ВЛАС|НАЗВА_ВЛАС|КОД_РИЗ|КОД_ПІДК|НАЗВА_ПІДК|КІЛЬК_ПОВ|ПОВЕРХ|КОД_ВОГН|КОД_КАТЕГОР|КОД_МІСЦ|НАЗВА_МІСЦ|КОД_ВИРІБ|НАЗВА_ВИРІБ|КОД_ПРИЧ|НАЗВА_ПРИЧ|ВИЯВЛ_ЗАГ|ВИЯВЛ_ДІТ|ЗАГ_ВНАСЛ|ЗАГ_ДІТ|ЗАГ_ОС|ПІБ1_ЗАГ|ПІБ2_ЗАГ|ПІБ3_ЗАГ|ПІБ4_ЗАГ|ПІБ5_ЗАГ|ВІК1_ЗАГ|ВІК2_ЗАГ|ВІК3_ЗАГ|ВІК4_ЗАГ|ВІК5_ЗАГ|СТАТЬ1_ЗАГ|СТАТЬ2_ЗАГ|СТАТЬ3_ЗАГ|СТАТЬ4_ЗАГ|СТАТЬ5_ЗАГ|СОЦСТАТ1_ЗАГ|СОЦСТАТ2_ЗАГ|СОЦСТАТ3_ЗАГ|СОЦСТАТ4_ЗАГ|СОЦСТАТ5_ЗАГ|МОМ_НАСТ1|МОМ_НАСТ2|МОМ_НАСТ3|МОМ_НАСТ4|МОМ_НАСТ5|УМОВ_ВПЛИВ1|УМОВ_ВПЛИВ2|УМОВ_ВПЛИВ3|УМОВ_ВПЛИВ4|УМОВ_ВПЛИВ5|ТРАВМ_ПОЖ|ТРАВМ_ДІТ|ТРАВМ_ОС|ПРЯМ_ЗБИТ|ПОБ_ЗБИТ|ЗН_БУД|ПОШК_БУД|ЗН_ТЕХН|ПОШК_ТЕХН|ЗН_ЗЕРН|ЗН_ХЛІБ_КОР|ЗН_ХЛІБ_ВАЛК|ЗН_КОРМ|ЗН_ТОРФ|ПОШК_ТОРФ|ЗАГ_ТВАР|ЗАГ_ПТИЦ|ЗН_ТЕКСТ|ВР_ЛЮД|ВР_ДІТ|ВР_ТВАР|ВР_ПТИЦ|ВР_БУД|ВР_ТЕХН|ВР_ЗЕРН|ВР_ХЛІБ_КОР|ВР_ХЛІБ_ВАЛК|ВР_КОРМ|ВР_ТОРФ|ВР_ТЕКСТ|ВР_МАТЦІН|ДАТА_ПОВІД|ЧАС_ПОВІД|ЧАС_ПРИБ|ІНФ_ЛІКВ_ПОЖ|ДАТА_ЛОК|ЧАС_ЛОК|ДАТА_ЛІКВ|ЧАС_ЛІКВ|УМОВ_ПОШИР1|УМОВ_ПОШИР2|УМОВ_ПОШИР3|УМОВ_ПОШИР4|УМОВ_ПОШИР5|УМОВ_УСКЛ1|УМОВ_УСКЛ2|УМОВ_УСКЛ3|УМОВ_УСКЛ4|УМОВ_УСКЛ5|НАЯВ_СПЗ|КОД1_СПЗ|КОД2_СПЗ|КОД3_СПЗ|КОД4_СПЗ|КОД5_СПЗ|КОД1_ДІЇ_СПЗ|КОД2_ДІЇ_СПЗ|КОД3_ДІЇ_СПЗ|КОД4_ДІЇ_СПЗ|КОД5_ДІЇ_СПЗ|УЧАСН1|УЧАСН2|УЧАСН3|УЧАСН4|УЧАСН5|КІЛЬК_УЧ1|КІЛЬК_УЧ2|КІЛЬК_УЧ3|КІЛЬК_УЧ4|КІЛЬК_УЧ5|ТЕХН1|ТЕХН2|ТЕХН3|ТЕХН4|ТЕХН5|КІЛЬК_ТЕХН1|КІЛЬК_ТЕХН2|КІЛЬК_ТЕХН3|КІЛЬК_ТЕХН4|КІЛЬК_ТЕХН5|CТВ1|CТВ2|CТВ3|КІЛЬК_СТВ1|КІЛЬК_СТВ2|КІЛЬК_СТВ3|ВОГН_РЕЧ1|ВОГН_РЕЧ2|ВОГН_РЕЧ3|ПЕРВ_ЗАС1|ПЕРВ_ЗАС2|ПЕРВ_ЗАС3|ДЖЕР1|ДЖЕР2|ДЖЕР3|ВИК_ГДЗС|КІЛЬК_ЛАНОК|ЧАС_ГДЗС|ДАТА_ПЕРЕВ|КОД_ВИД_ПЕРЕВ|КОД_УМОВ_ДІЯЛН|КОД_ЗАХ1|КОД_ЗАХ2|№_СТ_ККУ|ДАТА_ЗАПОВН|ПІБ_ЗАПОВН|" + Environment.NewLine;
            File.AppendAllText(path, str);

            if (checkedListBox1.CheckedItems.Count == 0)
            {
                foreach (var item in checkedListBox1.Items)
                {
                    rep = item.ToString().Split(' ');
                    rep = rep[2].Split(',');
                    select_code = "SELECT * FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1];
                    SavePog_Stat(select_code, path);
                }
               //select_code = "SELECT * FROM kartka_obliku";
               // SavePog_Stat(select_code, path);
            }
            else
            {
                foreach (var item in checkedListBox1.CheckedItems)
                {
                    rep = item.ToString().Split(' ');
                    rep = rep[2].Split(',');
                    select_code = "SELECT * FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1];
                    SavePog_Stat(select_code, path);
                }
               
            }
            MessageBox.Show("Ви створили файл POG_STAT");
        }
        public void SavePog_Stat(string select_code, string path)
        {
             string content = "";

               string[] names;
               string[] vik;
               string[] stat;
               string[] status;
               string[] code_moment;
               string[] code_umovi;
               string[] code_umovi_posh;
               string[] code_umovi_uskl;
               string[] code_system;
               string[] code_dii;
               string[] code_uch;
               string[] kilk_uch;
               string[] code_fireauto;
               string[] kilk_auto;
               string[] code_firestvol;
               string[] kilk_firestvol;
               string[] vogn_rech;
               string[] code_perv;
               string[] code_djerela;
               string[] code_zahodi;




               cmd_db = new SQLiteCommand(select_code, con_db);
               rdr = cmd_db.ExecuteReader();
               foreach (DbDataRecord record in rdr)
               {
                   names = record["zag_names"].ToString().Split(',');
                   vik = record["zag_vik"].ToString().Split(',');
                   stat = record["zag_stat_code"].ToString().Split(',');
                   status = record["code_status"].ToString().Split(',');
                   code_moment = record["code_moment"].ToString().Split(',');
                   code_umovi = record["code_umovi"].ToString().Split(',');
                   code_umovi_posh = record["code_umovi_posh"].ToString().Split(',');
                   code_umovi_uskl = record["code_umovi_uskl"].ToString().Split(',');
                   code_system = record["code_system"].ToString().Split(',');
                   code_dii = record["code_dii"].ToString().Split(',');
                   code_uch = record["code_uch"].ToString().Split(',');
                   kilk_uch = record["kilk_uch"].ToString().Split(',');
                   code_fireauto = record["code_fireauto"].ToString().Split(',');
                   kilk_auto = record["kilk_auto"].ToString().Split(',');
                   code_firestvol = record["code_firestvol"].ToString().Split(',');
                   kilk_firestvol = record["kilk_firestvol"].ToString().Split(',');
                   vogn_rech = record["code_rechovini"].ToString().Split(',');
                   code_perv = record["code_pervini"].ToString().Split(',');
                   code_djerela = record["code_djerela"].ToString().Split(',');
                   code_zahodi = record["code_zahodi"].ToString().Split(',');

                   content += record["code_region"] + "|"+record["name_region"]+"|"+ record["code_raion"] + "|"+record["name_raion"]+"|"+record["code_region_item"] + "|" + record["number_cartka"] + "|" + record["main_dop"] + "|" + record["date_viniknenya"] + "|" + record["code_adress"] + "|" + record["name_adress"] + "|" + record["fire_code"] + "|" + record["fire_item"] + "|" + record["code_forma"] + "|" + record["name_forma"] + "|" + record["code_riziku"] + "|" + record["code_object"] + "|" + record["name_object"] + "|" + record["poverhovist"] + "|" + record["code_poverh"] + "|" + record["code_stoikist"] + "|" + record["code_category"] + "|" + record["code_place"] + "|" + record["name_place"] + "|" + record["code_virib"] + "|" + record["item_virib"] + "|" + record["code_pricini"] + "|" + record["name_pricini"] + "|" + record["viavleno"] + "|" + record["via_ditei"] + "|" + record["zag_vnaslidok"] + "|" + record["zag_ditei"] + "|" + record["zag_fire"] + "|" + names[0] + "|" + names[1] + "|" + names[2] + "|" + names[3] + "|" + names[4] + "|" + vik[0] + "|" + vik[1] + "|" + vik[2] + "|" + vik[3] + "|" + vik[4] + "|" + stat[0] + "|" + stat[1] + "|" + stat[2] + "|" + stat[3] + "|" + stat[4] + "|" + status[0] + "|" + status[1] + "|" + status[2] + "|" + status[3] + "|" + status[4] + "|" + code_moment[0] + "|" + code_moment[1] + "|" + code_moment[2] + "|" + code_moment[3] + "|" + code_moment[4] + "|" + code_umovi[0] + "|" + code_umovi[1] + "|" + code_umovi[2] + "|" + code_umovi[3] + "|" + code_umovi[4] + "|" + record["travm"] + "|" + record["travm_ditei"] + "|" + record["travm_fire"] + "|" + record["pramiy"] + "|" + record["pobichniy"] + "|" + record["zn_bud"] + "|" + record["posh_bud"] + "|" + record["zn_tehnika"] + "|" + record["posh_tehnika"] + "|" + record["zn_zerno"] + "|" + record["zn_koreni"] + "|" + record["zn_valki"] + "|" + record["zn_korm"] + "|" + record["zn_torf"] + "|" + record["posh_torf"] + "|" + record["zag_tvarin"] + "|" + record["zag_ptici"] + "|" + record["dop_info"] + "|" + record["vr_ludei"] + "|" + record["vr_ditei"] + "|" + record["vr_tvarin"] + "|" + record["vr_ptici"] + "|" + record["vr_bud"] + "|" + record["vr_tehnika"] + "|" + record["vr_zerno"] + "|" + record["vr_koreni"] + "|" + record["vr_valki"] + "|" + record["vr_korm"] + "|" + record["vr_torf"] + "|" + record["vr_dop"] + "|" + record["vr_mat"] + "|" + record["data_pov"] + "|" + record["time_pov"] + "|" + record["time_pributa"] + "|" + record["code_fire_likvid"] + "|" + record["data_lokal"] + "|" + record["time_lokal"] + "|" + record["data_likvid"] + "|" + record["time_likvid"] + "|" + code_umovi_posh[0] + "|" + code_umovi_posh[1] + "|" + code_umovi_posh[2] + "|" + code_umovi_posh[3] + "|" + code_umovi_posh[4] + "|" + code_umovi_uskl[0] + "|" + code_umovi_uskl[1] + "|" + code_umovi_uskl[2] + "|" + code_umovi_uskl[3] + "|" + code_umovi_uskl[4] + "|" + record["code_spz"] + "|" + code_system[0] + "|" + code_system[1] + "|" + code_system[2] + "|" + code_system[3] + "|" + code_system[4] + "|" + code_dii[0] + "|" + code_dii[1] + "|" + code_dii[2] + "|" + code_dii[3] + "|" + code_dii[4] + "|" + code_uch[0] + "|" + code_uch[1] + "|" + code_uch[2] + "|" + code_uch[3] + "|" + code_uch[4] + "|" + kilk_uch[0] + "|" + kilk_uch[1] + "|" + kilk_uch[2] + "|" + kilk_uch[3] + "|" + kilk_uch[4] + "|" + code_fireauto[0] + "|" + code_fireauto[1] + "|" + code_fireauto[2] + "|" + code_fireauto[3] + "|" + code_fireauto[4] + "|" + kilk_auto[0] + "|" + kilk_auto[1] + "|" + kilk_auto[2] + "|" + kilk_auto[3] + "|" + kilk_auto[4] + "|" + code_firestvol[0] + "|" + code_firestvol[1] + "|" + code_firestvol[2] + "|" + kilk_firestvol[0] + "|" + kilk_firestvol[1] + "|" + kilk_firestvol[2] + "|" + vogn_rech[0] + "|" + vogn_rech[1] + "|" + vogn_rech[2] + "|" + code_perv[0] + "|" + code_perv[1] + "|" + code_perv[2] + "|" + code_djerela[0] + "|" + code_djerela[1] + "|" + code_djerela[2] + "|" + record["vikr_gds"] + "|" + record["kilk_gds"] + "|" + record["time_gds"] + "|" + record["data_perevirki"] + "|" + record["code_perevirka"] + "|" + record["code_dialnist"] + "|" + code_zahodi[0] + "|" + code_zahodi[1] + "|" + record["number_kku"] + "|"+ record["data_zapovnenya"] + "|" + record["pid_osibi"] + Environment.NewLine;

                   content.Trim().Replace(" ", "");
               }
               File.AppendAllText(path, content);
           
        }

        private void button10_Click(object sender, EventArgs e)
        {

            //Console.WriteLine(dateTimePicker11.Value.Date);
            CheckedListBox list = new CheckedListBox();
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                list.Items.Add(checkedListBox1.Items[i]);
            }
            checkedListBox1.Items.Clear();
            string[] date_string;
           
            
            for (int i = 0; i < list.Items.Count; i++)
            {
                date_string = list.Items[i].ToString().Split('(');
                date_string[1] = date_string[1].Remove(date_string[1].Length - 1);
                DateTime date = DateTime.Parse(date_string[1]);
                if(date>=dateTimePicker11.Value && date <= dateTimePicker12.Value)
                {
                    checkedListBox1.Items.Add(list.Items[i]);
                    
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //ReadDb();
            cmd_db = new SQLiteCommand("SELECT `number_cartka`,`main_dop`,`date_viniknenya` from kartka_obliku", con_db);
            rdr = cmd_db.ExecuteReader();
            checkedListBox1.Items.Clear();

            while (rdr.Read())
            {
                // region_items.Add(rdr[1].ToString());
                checkedListBox1.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
            }
        }

        private void textBox81_Leave(object sender, EventArgs e)
        {
            
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
           
        }

        private void textBox102_Leave(object sender, EventArgs e)
        {
            
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            textBox3.BackColor = Color.White;
            textBox3.ForeColor = Color.Black;
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
           
        }

        private void створитиФорму701ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel4.Visible = true;
            panel3.Visible = false;
            panel2.Visible = false;
            panel5.Visible = false;
            cmd_db = new SQLiteCommand("SELECT `number_cartka`,`main_dop`,`date_viniknenya` from kartka_obliku", con_db);
            rdr = cmd_db.ExecuteReader();
            checkedListBox2.Items.Clear();

            while (rdr.Read())
            {
                // region_items.Add(rdr[1].ToString());
                checkedListBox2.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            int count_fire = 0;
            float pramiy = 0;
            float pobichniy = 0;
            int via_ditei = 0;
            int viavleno = 0;
            int city_vil = 0;
            int pidpr = 0;
            float city_zb = 0;
            float city_pobichniy = 0;
            int viavl_city = 0;
            int viavl_city_ditei = 0;
            float pidpr_zb = 0;
            float pidpr_pob_zb = 0;
            int pidpr_viavl = 0;
            int pidpr_viavl_ditei = 0;
            string[] rep;
            string select_code = "";
            string code_region = "";
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `name_region` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1];
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();
                while (rdr.Read())
                {
                    code_region = rdr[0].ToString();
                }
            }
               

            while (rdr.Read())
            {
                if (rdr[0].ToString() == "0")
                    count_fire += 1;
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `main_dop` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop="+rep[1];
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if(rdr[0].ToString()=="0")
                    count_fire += 1;     
                }
            }

            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `pramiy` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1];
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    //if (rdr[0].ToString() == "0")
                        pramiy += float.Parse(rdr[0].ToString());
                    Console.WriteLine(rdr[0].ToString());
                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `pobichniy` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1];
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() != " ")
                        pobichniy += float.Parse(rdr[0].ToString());
                   
                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `via_ditei` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1];
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() != " ")
                        via_ditei += int.Parse(rdr[0].ToString());
                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `viavleno` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1];
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() != " ")
                        viavleno += int.Parse(rdr[0].ToString());
                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `main_dop` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1]+" AND (code_region_item=1 OR code_region_item=2 OR code_region_item=4)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString()=="0")
                      city_vil += 1;
                    //Console.WriteLine(rdr[0]);
                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `pramiy` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1] + " AND (code_region_item=1 OR code_region_item=2 OR code_region_item=4)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    city_zb += float.Parse(rdr[0].ToString());
                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `pobichniy` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1] + " AND (code_region_item=1 OR code_region_item=2 OR code_region_item=4)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() != " ")
                        city_pobichniy += float.Parse(rdr[0].ToString());
                    
                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `viavleno` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1] + " AND (code_region_item=1 OR code_region_item=2 OR code_region_item=4)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() != " ")
                        viavl_city += int.Parse(rdr[0].ToString());

                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `via_ditei` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1] + " AND (code_region_item=1 OR code_region_item=2 OR code_region_item=4)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() != " ")
                        viavl_city_ditei += int.Parse(rdr[0].ToString());

                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `main_dop` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1] + " AND (code_object=11 OR code_object=12 OR code_object=13)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() == "0")
                        pidpr += 1;

                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `pramiy` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1] + " AND (code_object=11 OR code_object=12 OR code_object=13)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    pidpr_zb += float.Parse(rdr[0].ToString());

                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `pobichniy` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1] + " AND (code_object=11 OR code_object=12 OR code_object=13)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() != " ")
                        pidpr_pob_zb += float.Parse(rdr[0].ToString());

                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `viavleno` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1] + " AND (code_object=11 OR code_object=12 OR code_object=13)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() != " ")
                        pidpr_viavl += int.Parse(rdr[0].ToString());

                }
            }
            foreach (var item in checkedListBox2.Items)
            {
                rep = item.ToString().Split(' ');
                rep = rep[2].Split(',');
                select_code = "SELECT `via_ditei` FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1] + " AND (code_object=11 OR code_object=12 OR code_object=13)";
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    if (rdr[0].ToString() != " ")
                        pidpr_viavl_ditei += int.Parse(rdr[0].ToString());

                }
            }


            /* cmd_db = new SQLiteCommand("SELECT `pramiy` from kartka_obliku WHERE data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
             rdr = cmd_db.ExecuteReader();

             while (rdr.Read())
             {
                 pramiy += float.Parse(rdr[0].ToString());
             }*/
            /* cmd_db = new SQLiteCommand("SELECT `pobichniy` from kartka_obliku WHERE data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
             rdr = cmd_db.ExecuteReader();

             while (rdr.Read())
             {
                 if(rdr[0].ToString()!=" ")
                 pobichniy += float.Parse(rdr[0].ToString());
             }*/

            /*cmd_db = new SQLiteCommand("SELECT `via_ditei` from kartka_obliku WHERE data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
            rdr = cmd_db.ExecuteReader();

            while (rdr.Read())
            {
                if (rdr[0].ToString() != " ")
                    via_ditei += int.Parse(rdr[0].ToString());
            }*/

            /* cmd_db = new SQLiteCommand("SELECT `viavleno` from kartka_obliku WHERE data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
             rdr = cmd_db.ExecuteReader();

             while (rdr.Read())
             {
                  if (rdr[0].ToString() != " ")
                 viavleno += int.Parse(rdr[0].ToString());
             }*/

            /*cmd_db = new SQLiteCommand("SELECT COUNT(*) from kartka_obliku WHERE main_dop=0 AND (code_region_item=1 OR code_region_item=2 OR code_region_item=4) AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
            rdr = cmd_db.ExecuteReader();

            while (rdr.Read())
            {
                // if (rdr[0].ToString() != " ")
                city_vil = int.Parse(rdr[0].ToString());
            }*/



            /* cmd_db = new SQLiteCommand("SELECT `pramiy` from kartka_obliku WHERE (code_region_item=1 OR code_region_item=2 OR code_region_item=4) AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
             rdr = cmd_db.ExecuteReader();

             while (rdr.Read())
             {
                 // if (rdr[0].ToString() != " ")
                 city_zb += float.Parse(rdr[0].ToString());
             }*/

            /* cmd_db = new SQLiteCommand("SELECT `pobichniy` from kartka_obliku WHERE (code_region_item=1 OR code_region_item=2 OR code_region_item=4)  AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
             rdr = cmd_db.ExecuteReader();

             while (rdr.Read())
             {
                  if (rdr[0].ToString() != " ")
                 city_pobichniy += float.Parse(rdr[0].ToString());
             }*/

            /*  cmd_db = new SQLiteCommand("SELECT `viavleno` from kartka_obliku WHERE (code_region_item=1 OR code_region_item=2 OR code_region_item=4) AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
              rdr = cmd_db.ExecuteReader();

              while (rdr.Read())
              {
                   if (rdr[0].ToString() != " ")
                  viavl_city += int.Parse(rdr[0].ToString());
              }*/

            /* cmd_db = new SQLiteCommand("SELECT `via_ditei` from kartka_obliku WHERE (code_region_item=1 OR code_region_item=2 OR code_region_item=4) AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
             rdr = cmd_db.ExecuteReader();

             while (rdr.Read())
             {
                  if (rdr[0].ToString() != " ")
                 viavl_city_ditei += int.Parse(rdr[0].ToString());
             }*/

            /*cmd_db = new SQLiteCommand("SELECT COUNT(*) from kartka_obliku WHERE main_dop=0 AND (code_object=11 OR code_object=12 OR code_object=13) AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
            rdr = cmd_db.ExecuteReader();

            while (rdr.Read())
            {
                // if (rdr[0].ToString() != " ")
                pidpr = int.Parse(rdr[0].ToString());
            }*/

            /* cmd_db = new SQLiteCommand("SELECT `pramiy` from kartka_obliku WHERE (code_object=11 OR code_object=12 OR code_object=13) AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
             rdr = cmd_db.ExecuteReader();

             while (rdr.Read())
             {
                 // if (rdr[0].ToString() != " ")
                 pidpr_zb += float.Parse(rdr[0].ToString());
             }*/

            /* cmd_db = new SQLiteCommand("SELECT `pobichniy` from kartka_obliku WHERE (code_object=11 OR code_object=12 OR code_object=13) AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
             rdr = cmd_db.ExecuteReader();

             while (rdr.Read())
             {
                  if (rdr[0].ToString() != " ")
                 pidpr_pob_zb += float.Parse(rdr[0].ToString());
             }*/

            /* cmd_db = new SQLiteCommand("SELECT `viavleno` from kartka_obliku WHERE (code_object=11 OR code_object=12 OR code_object=13) AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
             rdr = cmd_db.ExecuteReader();

             while (rdr.Read())
             {
                  if (rdr[0].ToString() != " ")
                 pidpr_viavl += int.Parse(rdr[0].ToString());
             }*/

            /*cmd_db = new SQLiteCommand("SELECT `via_ditei` from kartka_obliku WHERE (code_object=11 OR code_object=12 OR code_object=13) AND data_zapovnenya BETWEEN '" + dateTimePicker13.Value.ToShortDateString() + "' AND '" + dateTimePicker14.Value.ToShortDateString() + "'", con_db);
            rdr = cmd_db.ExecuteReader();

            while (rdr.Read())
            {
                 if (rdr[0].ToString() != " ")
                pidpr_viavl_ditei += int.Parse(rdr[0].ToString());
            }*/

             string path = Environment.CurrentDirectory + "/export/form701.xlsx";
             if (File.Exists(path))
             {
                 File.Delete(path);
             }
             File.Copy(Environment.CurrentDirectory + "/MyDataBase/form701_template.xlsx", Environment.CurrentDirectory + "/export/form701.xlsx");


             var doc = new Excel.Application();
             doc.Visible = false;

             Excel.Workbook excelappworkbooks = doc.Workbooks.Open(path);

             var excelsheets = excelappworkbooks.Worksheets;

             var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            var excelcells0 = excelworksheet.get_Range("H7", "H7");
            excelcells0.Value2 = code_region;
            var excelcells = excelworksheet.get_Range("C18", "C18");
             excelcells.Value2 = count_fire;
            var excelcells1 = excelworksheet.get_Range("C19", "C19");
            excelcells1.Value2 = pramiy.ToString("#.##");
            var excelcells2 = excelworksheet.get_Range("C20", "C20");
            excelcells2.Value2 = pobichniy.ToString("#.##");
            var excelcells3 = excelworksheet.get_Range("C21", "C21");
            excelcells3.Value2 = viavleno;
            var excelcells4 = excelworksheet.get_Range("C22", "C22");
            excelcells4.Value2 = via_ditei;

            var excelcells5 = excelworksheet.get_Range("D18", "D18");
            excelcells5.Value2 = city_vil;
            var excelcells6 = excelworksheet.get_Range("D19", "D19");
            excelcells6.Value2 = city_zb.ToString("#.##");
            var excelcells7 = excelworksheet.get_Range("D20", "D20");
            excelcells7.Value2 = city_pobichniy.ToString("#.##");
            var excelcells8 = excelworksheet.get_Range("D21", "D21");
            excelcells8.Value2 = viavl_city;
            var excelcells9 = excelworksheet.get_Range("D22", "D22");
            excelcells9.Value2 = viavl_city_ditei;

            var excelcells10 = excelworksheet.get_Range("E18", "E18");
            excelcells10.Value2 = pidpr;
            var excelcells11 = excelworksheet.get_Range("E19", "E19");
            excelcells11.Value2 = pidpr_zb.ToString("#.##");
            var excelcells12 = excelworksheet.get_Range("E20", "E20");
            excelcells12.Value2 = pidpr_pob_zb.ToString("#.##");
            var excelcells13 = excelworksheet.get_Range("E21", "E21");
            excelcells13.Value2 = pidpr_viavl;
            var excelcells14 = excelworksheet.get_Range("E22", "E22");
            excelcells14.Value2 = pidpr_viavl_ditei;

            excelappworkbooks.Save();
             doc.Quit();
            MessageBox.Show("Форма 701 успішно створена!");
        }

        private void button13_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            button1.Visible = true;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            CheckedListBox list = new CheckedListBox();
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                list.Items.Add(checkedListBox2.Items[i]);
            }
            checkedListBox2.Items.Clear();
            string[] date_string;


            for (int i = 0; i < list.Items.Count; i++)
            {
                date_string = list.Items[i].ToString().Split('(');
                date_string[1] = date_string[1].Remove(date_string[1].Length - 1);
                DateTime date = DateTime.Parse(date_string[1]);
                if (date >= dateTimePicker13.Value && date <= dateTimePicker14.Value)
                {
                    checkedListBox2.Items.Add(list.Items[i]);
                }
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            cmd_db = new SQLiteCommand("SELECT `number_cartka`,`main_dop`,`date_viniknenya` from kartka_obliku", con_db);
            rdr = cmd_db.ExecuteReader();
            checkedListBox2.Items.Clear();

            while (rdr.Read())
            {
                // region_items.Add(rdr[1].ToString());
                checkedListBox2.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.PageUp)
            {
                foreach (var item in tb)
                {
                    if (item.Focused)
                    {
                        GetTable(item.Name, e);
                        
                    }
                }
                foreach (var item in tb_masked)
                {
                    if (item.Focused)
                    {
                        GetTable(item.Name, e);

                    }
                }


            }
            if (e.KeyCode == Keys.Right)
            {
                if (comboBox3.Focused)
                {
                    comboBox4.DroppedDown = true;
                    comboBox4.Focus();
                }
                if (comboBox12.Focused)
                {
                    comboBox13.DroppedDown = true;
                    comboBox13.Focus();
                }
            }
            if (e.KeyCode == Keys.Left)
            {
                if (comboBox4.Focused)
                {
                    comboBox3.DroppedDown = true;
                    comboBox3.Focus();
                }
                if (comboBox13.Focused)
                {
                    comboBox12.DroppedDown = true;
                    comboBox12.Focus();
                }

            }
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                SendKeys.Send("{TAB}");
            }

            if (e.KeyCode == Keys.F1)
            {
                //current = 0;
                
                foreach (var item in tb)
                { 
                    if (item.Focused)
                    {
                        if (item.TabIndex == 155 && maskedTextBox5.Enabled == false)
                        {
                            current = 154;
                        }
                        else if (item.TabIndex == 116 && textBox113.Text=="")
                        {
                            current = 106;
                        }
                        else
                        {
                            current = item.TabIndex;
                        }
                            
                    }  
                }
                foreach (var item in tb)
                {
                   
                        if (item.TabIndex == current - 1)
                        {
                            item.Focus();
                        }
                }
                /* foreach (var item in tb_date)
                 {
                    // if(item.Name== "dateTimePicker3" || item.Name == "dateTimePicker4" || item.Name == "dateTimePicker5" || item.Name == "dateTimePicker7")
                    // {
                         if (item.Focused)
                         {
                             current = item.TabIndex;
                         }
                    // }

                 }
                 foreach (var item in tb_date)
                 {
                    // if (item.Name == "dateTimePicker3" || item.Name == "dateTimePicker4" || item.Name == "dateTimePicker5" || item.Name == "dateTimePicker7")
                    // {
                         if (item.TabIndex == current - 1)
                         {
                             item.Focus();
                         }
                   //  }
                 }*/
              
                foreach (var item in tb_masked)
                {
                    if (item.Focused)
                    {
                        current = item.TabIndex;
                       
                    }
                }
                foreach (var item in tb_masked)
                {
                    if (item.TabIndex == current - 1)
                    {
                       // Console.WriteLine(current);
                        item.Focus();
                    }
                }
              //  Console.WriteLine(current.ToString());
            }
        }
        public void GetTable(string name, KeyEventArgs e)
        {
      
            TextBox[] names_table = {textBox1,textBox70,textBox71,textBox74,textBox75,textBox76,textBox68,textBox77,textBox78,textBox79,textBox81, textBox10, textBox83, textBox84, textBox85, textBox86, textBox87, textBox88, textBox89, textBox90,
                textBox91, textBox92, textBox93, textBox94, textBox95, textBox96, textBox97, textBox98, textBox99, textBox100, textBox101, textBox102, textBox103,textBox104,textBox105,textBox106,textBox107,textBox112,textBox111,textBox110,textBox109,textBox108,textBox113,
                textBox118,textBox117,textBox116,textBox115,textBox114,textBox123,textBox122,textBox121,textBox120,textBox119,textBox128,textBox127,textBox126,textBox125,textBox124,textBox129,textBox130,textBox131,textBox132,textBox133,textBox134,textBox135,textBox136,textBox137,textBox138,textBox139,textBox140,textBox141,textBox142,textBox143,textBox144,textBox145,
            textBox146,textBox147,textBox148,textBox149};
            ComboBox[] names_combo = { comboBox83, comboBox2, comboBox85, comboBox5, comboBox6, comboBox7, comboBox8, comboBox9, comboBox10, comboBox11, comboBox24, comboBox14, comboBox15, comboBox16, comboBox17, comboBox18, comboBox19, comboBox20,
            comboBox21,comboBox22,comboBox23,comboBox25,comboBox26,comboBox27,comboBox28,comboBox29,comboBox30,comboBox31,comboBox32,comboBox33,comboBox34,comboBox35,comboBox36,comboBox37,comboBox38,comboBox39,comboBox40,comboBox45,comboBox44,comboBox43,
            comboBox42,comboBox41,comboBox46,comboBox47,comboBox48,comboBox49,comboBox50,comboBox51,comboBox52,comboBox53,comboBox54,comboBox55,comboBox56,comboBox57,comboBox58,comboBox59,comboBox60,comboBox61,comboBox66,comboBox65,comboBox64,comboBox63,comboBox62,comboBox67,
            comboBox68,comboBox69,comboBox72,comboBox71,comboBox70,comboBox75,comboBox74,comboBox73,comboBox78,comboBox77,comboBox76,comboBox79,comboBox80,comboBox81,comboBox82};
            MaskedTextBox[] masked = { maskedTextBox1, maskedTextBox2, maskedTextBox3, maskedTextBox4, maskedTextBox5, maskedTextBox6 };
            DateTimePicker[] datetime = { dateTimePicker1, dateTimePicker2, dateTimePicker6, dateTimePicker8, dateTimePicker9, dateTimePicker10 };
            for(int i=0; i < names_table.Length; i++)
            {
                if (names_table[i].Name == name)
                {
                    names_combo[i].TabIndex = names_table[i].TabIndex;
                        if (names_combo[i].DroppedDown)
                        {
                            names_combo[i].DroppedDown = false;
                        }
                        else
                        {
                            names_combo[i].DroppedDown = true;
                            names_combo[i].Focus();
                        }
                    
                }
            }
            if (name == textBox72.Name)
            {
                if (!comboBox3.DroppedDown)
                {
                    comboBox3.DroppedDown = true;
                    comboBox3.Focus();
                }
            }
            if (name == textBox80.Name)
            {
                if (!comboBox12.DroppedDown)
                {
                    comboBox12.DroppedDown = true;
                    comboBox12.Focus();
                }
            }

            for (int i = 0; i < masked.Length; i++)
            {
                if (masked[i].Name == name)
                {
                   
                    datetime[i].Focus();
                    SendKeys.SendWait("%{DOWN}");
                   // Console.WriteLine(datetime[i].Focused);
                }
            }
        }
    
        private void dateTimePicker1_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата виникнення пожежі";
            bool state = false;
            int code;
            state = int.TryParse(textBox3.Text, out int res);

            string content_val = "";

            string select_code = "SELECT `number_cartka` FROM kartka_obliku WHERE code_raion=" + textBox1.Text + " AND number_cartka=" + textBox2.Text + " AND main_dop=" + textBox3.Text;
            //  connection();
            try
            {
                //if (con_db.State == ConnectionState.Open)
                //{
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    content_val = rdr[0].ToString();
                    // dict.Add(code, rdr[0].ToString());
                }
                // }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            if (state)
            {
                if (content_val != "" && flag_edit==false)
                {
                    MessageBox.Show("Картка під таким номером вже існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox2.BackColor = Color.Firebrick;
                    textBox2.ForeColor = Color.White;
                    textBox2.Select();
                    textBox2.ScrollToCaret();
                }
                else
                {
                    code = int.Parse(textBox3.Text);
                    if (code > 0 && code <= 9)
                    {
                        textBox3.ForeColor = Color.Firebrick;
                    }
                    if (code == 0)
                    {
                        textBox3.ForeColor = Color.Black;
                    }
                    if (code > 9)
                    {
                        MessageBox.Show("В це поле можно внести тільки коди від 0 до 9", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        textBox3.Select();
                        textBox3.ScrollToCaret();
                    }
                }

            }
            else
            {
                MessageBox.Show("В це поле можно внести тільки цифри", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox3.Text = "0";
                textBox2.Select();
                textBox2.ScrollToCaret();
            }
        }

        private void textBox73_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Об’єкт пожежі";
            textBox73.BackColor = Color.White;
            string code = CodeAllContent("fire_objects", "fire_code", "fire_item", textBox72.Text);
            if (code == "")
            {

                MessageBox.Show("Такого кода не існує", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox73.Text = "";
                textBox72.Text = "";
                textBox72.Select();
                textBox72.ScrollToCaret();
            }
        }

        private void textBox156_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Місце виникнення пожежі";
            textBox156.BackColor = Color.White;
            string code = CodeAllContent("place_fire", "code_place", "code_place", textBox79.Text);
            // if (textBox71.Text != "")
            if (code == "")
            {
                MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox79.Text = "";
                textBox79.Select();
                textBox79.ScrollToCaret();
            }
        }

        private void textBox157_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Виріб-ініціатор пожежі";
            textBox157.BackColor = Color.White;
            string code = CodeAllContent("virib_iniciator", "code_virib", "item_virib", textBox80.Text);
            //if (textBox80.Text != "")
            if (code == "")
            {
                MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox80.Text = "";
                textBox80.BackColor = Color.Firebrick;
                textBox80.Select();
                textBox80.ScrollToCaret();
            }
        }

        private void textBox158_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Причина пожежі";
            textBox158.BackColor = Color.White;
            string code = CodeAllContent("pricini_fire", "code_pricini", "code_pricini", textBox81.Text);
            //if (textBox81.Text != "")
            if (code == "")
            {
                MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox81.Text = "";
                textBox81.BackColor = Color.Firebrick;
                textBox81.Select();
                textBox81.ScrollToCaret();
            }
        }

        private void dateTimePicker3_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "час повідомлення про пожежу";
            if (dateTimePicker2.Value.ToString() == "01.01.1900 0:00:00")
            {
                MessageBox.Show("Заповніть дату повідомлення про пожежу");
                dateTimePicker2.Select();
            }
            //Console.WriteLine(dateTimePicker3.Value.ToShortTimeString());
        }

        private void dateTimePicker6_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата локалізації пожежі";

           /* string code = CodeAllContent("info_fire", "code_fire_likvid", "code_fire_likvid", textBox102.Text);
            //if (textBox81.Text != "")
            if (textBox102.Text == "" && textBox3.Text=="0")
            {
                MessageBox.Show("Це поле обов'язкове для заповнення", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox102.Select();
                textBox102.ScrollToCaret();
            }
            else if (code == "")
            {
                MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox102.Text = "";
                textBox102.BackColor = Color.Firebrick;
                textBox102.Select();
                textBox102.ScrollToCaret();
            }*/
           
        }

        private void dateTimePicker7_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Час ліквідації пожежі";
            if (dateTimePicker8.Value.ToString() == "01.01.1900 0:00:00")
            {
                MessageBox.Show("Заповніть дату ліквідації пожежі");
                dateTimePicker8.Select();
            }
        }

        private void dateTimePicker9_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата останньої перевірки";
        }

        private void textBox80_TextChanged(object sender, EventArgs e)
        {
            if (textBox80.Text == "")
            {
                textBox157.Text = "";
               
            }
            else
            {
                bool res = int.TryParse(textBox80.Text, out int code);
                if (code == 109 || code == 114 || code == 123 || code == 129 || code == 132 || code == 133 || code == 208 || code == 215 || code == 223 || code == 228 || code == 229 || code == 303
                || code == 409 || code == 513 || code == 610 || code == 705 || code == 804 || code == 806 || code == 904 || code == 1004 || code == 1103 || code == 1210 || code == 1226)
                {
                    textBox157.Text = "";
                }
                else
                {
                    textBox157.Text = CodeAllContent("virib_iniciator", "code_virib", "item_virib", textBox80.Text);
                }
               
            }
        }

        private void dateTimePicker2_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата повідомлення про пожежу";
        }

        private void dateTimePicker4_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Час прибуття 1 - го підрозділу";
        }

        private void dateTimePicker5_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Час локалізації пожежі ";
        }

        private void dateTimePicker8_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата ліквідації пожежі";
        }

        private void dateTimePicker10_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата заповнення Картки обліку пожежі";
        }

        private void редагуванняToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel5.Visible = true;
            button17.Text = "Перейти к редагуванню";
            button16.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            
            radioButton1.Checked = false;
            radioButton2.Checked = false;

            cmd_db = new SQLiteCommand("SELECT `number_cartka`,`main_dop`,`date_viniknenya` from kartka_obliku", con_db);
            rdr = cmd_db.ExecuteReader();
            listBox3.Items.Clear();

            while (rdr.Read())
            {
                // region_items.Add(rdr[1].ToString());
                listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
            }
            label108.Text = listBox3.Items.Count.ToString();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            panel5.Visible = false;
            button1.Visible = true;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if(button17.Text=="Перейти к редагуванню")
            {
                if (listBox3.SelectedIndex != -1)
                {
                    foreach (var item in tb)
                    {
                        item.Enabled = true;

                    }
                    foreach (var item in tb_date)
                    {
                        item.Enabled = true;

                    }

                    button1.Visible = true;
                    flag_edit = true;
                    panel5.Visible = false;
                }
                else
                {
                    return;
                }
                   
            }
            if(button17.Text== "Переглянути картку")
            {
                if (listBox3.SelectedIndex != -1)
                {
                    button16.Visible = false;
                    foreach (var item in tb)
                    {
                        item.Enabled = false;

                    }
                    foreach (var item in tb_date)
                    {
                        item.Enabled = false;

                    }
                    panel5.Visible = false;
                }
                else
                {
                    return;
                }
                   
            }
            if (button17.Text == "Видалити картку")
            {
                
                try
                {
                    if (listBox3.SelectedIndex!=-1)
                    {
                        string res = listBox3.SelectedItem.ToString();
                        string[] rep;
                        rep = res.Split(' ');
                        rep = rep[2].Split('(');

                        DialogResult res_del = MessageBox.Show("Ви впевнені що хочете видалити цю картку", "Увага!", MessageBoxButtons.YesNo);
                        if (res_del == DialogResult.Yes)
                        {
                            cmd_db = new SQLiteCommand("DELETE from kartka_obliku WHERE `id_kartki`=" + rep[1], con_db);
                            rdr = cmd_db.ExecuteReader();
                            MessageBox.Show("Картка видалена з бази даних");
                            cmd_db = new SQLiteCommand("SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku", con_db);
                            rdr = cmd_db.ExecuteReader();
                            listBox3.Items.Clear();

                            while (rdr.Read())
                            {
                                // region_items.Add(rdr[1].ToString());
                                listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        DialogResult res_del = MessageBox.Show("Ви впевнені що хочете видалити всі картки", "Увага!", MessageBoxButtons.YesNo);
                        if (res_del == DialogResult.Yes)
                        {
                            cmd_db = new SQLiteCommand("DELETE from kartka_obliku", con_db);
                            rdr = cmd_db.ExecuteReader();
                            MessageBox.Show("Всі картки видалені з бази даних");
                            listBox3.Items.Clear();
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {

                    Console.WriteLine(ex.Message);
                }
                label108.Text = listBox3.Items.Count.ToString();

            }

            
            foreach (var item in tb)
            {
                if (item.Name != "textBox69")
                {
                    item.Clear();
                }
                if (item.Name == "textBox3")
                {
                    item.Text = "0";
                }

            }
            foreach (var item in tb_date)
            {
                if (item.Name == "dateTimePicker10")
                {
                    item.Value = DateTime.Today;
                }
                else
                {
                    item.Value = DateTime.Parse("01.01.1900");
                }

            }
            foreach (var item in tb_masked)
            {
                if (item.Name != maskedTextBox6.Name)
                {
                    item.Text = "";
                }
            }
            try
            {
                string res = listBox3.SelectedItem.ToString();
                string[] rep;
                rep = res.Split(' ');
                rep = rep[2].Split(',');
                string select_code = "SELECT * FROM kartka_obliku WHERE number_cartka=" + rep[0] + " AND main_dop=" + rep[1];

                EditForm edit = new EditForm();
                edit.GetFieldfromDb(select_code,this);
                foreach (var item in tb)
                {

                    item.BackColor = Color.White;
                    
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
            
        }

        private void переглядКартокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel5.Visible = true;
            button17.Text = "Переглянути картку";
            button1.Visible = false;
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            // button16.Visible = false;
            cmd_db = new SQLiteCommand("SELECT `number_cartka`,`main_dop`,`date_viniknenya` from kartka_obliku", con_db);
            rdr = cmd_db.ExecuteReader();
            listBox3.Items.Clear();

            while (rdr.Read())
            {
                // region_items.Add(rdr[1].ToString());
                listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
            }
            label108.Text = listBox3.Items.Count.ToString();
        }

        private void открітьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flag_edit = false;
            string filename="";
            string path = Environment.CurrentDirectory + "/export/POG_STAT_" + textBox69.Text + ".txt";
            string[] str_db=null;
            string content_val="";
            Dictionary<string, string> dict_db = new Dictionary<string, string>();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            bool flag_res = false;
            bool flag_good = false;
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            DialogResult res = DialogResult.No;
            bool IsMessageShow = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                 if ((openFileDialog1.OpenFile()) != null)
                {
                    filename=openFileDialog1.FileName;
                }
            }
            if (!File.Exists(path))
            {
                string str = "КОД_РЕГ|НАЗВА_РЕГ|КОД_РАЙОНУ|НАЗВА_РАОЙНУ|ТИП_НП|№_КАРТКИ|ОСН_ДОД|ДАТА_ПОЖ|КОД_НП|АДРЕСА|КОД_ОБ|НАЗВА_ ОБ|КОД_ВЛАС|НАЗВА_ВЛАС|КОД_РИЗ|КОД_ПІДК|НАЗВА_ПІДК|КІЛЬК_ПОВ|ПОВЕРХ|КОД_ВОГН|КОД_КАТЕГОР|КОД_МІСЦ|НАЗВА_МІСЦ|КОД_ВИРІБ|НАЗВА_ВИРІБ|КОД_ПРИЧ|НАЗВА_ПРИЧ|ВИЯВЛ_ЗАГ|ВИЯВЛ_ДІТ|ЗАГ_ВНАСЛ|ЗАГ_ДІТ|ЗАГ_ОС|ПІБ1_ЗАГ|ПІБ2_ЗАГ|ПІБ3_ЗАГ|ПІБ4_ЗАГ|ПІБ5_ЗАГ|ВІК1_ЗАГ|ВІК2_ЗАГ|ВІК3_ЗАГ|ВІК4_ЗАГ|ВІК5_ЗАГ|СТАТЬ1_ЗАГ|СТАТЬ2_ЗАГ|СТАТЬ3_ЗАГ|СТАТЬ4_ЗАГ|СТАТЬ5_ЗАГ|СОЦСТАТ1_ЗАГ|СОЦСТАТ2_ЗАГ|СОЦСТАТ3_ЗАГ|СОЦСТАТ4_ЗАГ|СОЦСТАТ5_ЗАГ|МОМ_НАСТ1|МОМ_НАСТ2|МОМ_НАСТ3|МОМ_НАСТ4|МОМ_НАСТ5|УМОВ_ВПЛИВ1|УМОВ_ВПЛИВ2|УМОВ_ВПЛИВ3|УМОВ_ВПЛИВ4|УМОВ_ВПЛИВ5|ТРАВМ_ПОЖ|ТРАВМ_ДІТ|ТРАВМ_ОС|ПРЯМ_ЗБИТ|ПОБ_ЗБИТ|ЗН_БУД|ПОШК_БУД|ЗН_ТЕХН|ПОШК_ТЕХН|ЗН_ЗЕРН|ЗН_ХЛІБ_КОР|ЗН_ХЛІБ_ВАЛК|ЗН_КОРМ|ЗН_ТОРФ|ПОШК_ТОРФ|ЗАГ_ТВАР|ЗАГ_ПТИЦ|ЗН_ТЕКСТ|ВР_ЛЮД|ВР_ДІТ|ВР_ТВАР|ВР_ПТИЦ|ВР_БУД|ВР_ТЕХН|ВР_ЗЕРН|ВР_ХЛІБ_КОР|ВР_ХЛІБ_ВАЛК|ВР_КОРМ|ВР_ТОРФ|ВР_ТЕКСТ|ВР_МАТЦІН|ДАТА_ПОВІД|ЧАС_ПОВІД|ЧАС_ПРИБ|ІНФ_ЛІКВ_ПОЖ|ДАТА_ЛОК|ЧАС_ЛОК|ДАТА_ЛІКВ|ЧАС_ЛІКВ|УМОВ_ПОШИР1|УМОВ_ПОШИР2|УМОВ_ПОШИР3|УМОВ_ПОШИР4|УМОВ_ПОШИР5|УМОВ_УСКЛ1|УМОВ_УСКЛ2|УМОВ_УСКЛ3|УМОВ_УСКЛ4|УМОВ_УСКЛ5|НАЯВ_СПЗ|КОД1_СПЗ|КОД2_СПЗ|КОД3_СПЗ|КОД4_СПЗ|КОД5_СПЗ|КОД1_ДІЇ_СПЗ|КОД2_ДІЇ_СПЗ|КОД3_ДІЇ_СПЗ|КОД4_ДІЇ_СПЗ|КОД5_ДІЇ_СПЗ|УЧАСН1|УЧАСН2|УЧАСН3|УЧАСН4|УЧАСН5|КІЛЬК_УЧ1|КІЛЬК_УЧ2|КІЛЬК_УЧ3|КІЛЬК_УЧ4|КІЛЬК_УЧ5|ТЕХН1|ТЕХН2|ТЕХН3|ТЕХН4|ТЕХН5|КІЛЬК_ТЕХН1|КІЛЬК_ТЕХН2|КІЛЬК_ТЕХН3|КІЛЬК_ТЕХН4|КІЛЬК_ТЕХН5|CТВ1|CТВ2|CТВ3|КІЛЬК_СТВ1|КІЛЬК_СТВ2|КІЛЬК_СТВ3|ВОГН_РЕЧ1|ВОГН_РЕЧ2|ВОГН_РЕЧ3|ПЕРВ_ЗАС1|ПЕРВ_ЗАС2|ПЕРВ_ЗАС3|ДЖЕР1|ДЖЕР2|ДЖЕР3|ВИК_ГДЗС|КІЛЬК_ЛАНОК|ЧАС_ГДЗС|ДАТА_ПЕРЕВ|КОД_ВИД_ПЕРЕВ|КОД_УМОВ_ДІЯЛН|КОД_ЗАХ1|КОД_ЗАХ2|№_СТ_ККУ|ДАТА_ЗАПОВН|ПІБ_ЗАПОВН|" + Environment.NewLine;
                File.AppendAllText(path, str);

                string[] loadfile = File.ReadAllLines(filename);

                for (int i = 1; i < loadfile.Length; i++)
                {
                    File.AppendAllText(path, loadfile[i] + Environment.NewLine);
                }

            }
            else
            {
                try
                {
                    string[] loadfile = File.ReadAllLines(filename);

                    for (int i = 1; i < loadfile.Length; i++)
                    {
                        //File.AppendAllText(path, loadfile[i]+Environment.NewLine);
                        str_db = loadfile[i].Split('|');
                        for (int j = 0; j < str_db.Length; j++)
                        {
                            dict_db["code_region"] = str_db[0];
                            dict_db["name_region"] = str_db[1];
                            dict_db["code_raion"] = str_db[2];
                            dict_db["name_raion"] = str_db[3];
                            dict_db["code_region_item"] = str_db[4];
                            dict_db["number_cartka"] = str_db[5];
                            dict_db["main_dop"] = str_db[6];
                            dict_db["date_viniknenya"] = str_db[7];

                            dict_db["code_adress"] = str_db[8];
                            dict_db["name_adress"] = str_db[9];
                            dict_db["fire_code"] = str_db[10];
                            dict_db["fire_item"] = str_db[11];
                            dict_db["code_forma"] = str_db[12];
                            dict_db["name_forma"] = str_db[13];
                            dict_db["code_riziku"] = str_db[14];
                            dict_db["name_riziku"] = "";
                            dict_db["code_object"] = str_db[15];
                            dict_db["name_object"] = str_db[16];
                            dict_db["poverhovist"] = str_db[17];
                            dict_db["code_poverh"] = str_db[18];
                            dict_db["name_poverh"] = "";
                            dict_db["code_stoikist"] = str_db[19];
                            dict_db["name_stoikist"] = "";
                            dict_db["code_category"] = str_db[20];
                            dict_db["name_category"] = "";
                            dict_db["code_place"] = str_db[21];
                            dict_db["name_place"] = str_db[22];
                            dict_db["code_virib"] = str_db[23];
                            dict_db["item_virib"] = str_db[24];
                            dict_db["code_pricini"] = str_db[25];
                            dict_db["name_pricini"] = str_db[26];

                            dict_db["viavleno"] = str_db[27];
                            dict_db["via_ditei"] = str_db[28];
                            dict_db["zag_vnaslidok"] = str_db[29];
                            dict_db["zag_ditei"] = str_db[30];
                            dict_db["zag_fire"] = str_db[31];
                            dict_db["zag_names"] = str_db[32] + "," + str_db[33] + "," + str_db[34] + "," + str_db[35] + "," + str_db[36];
                            dict_db["zag_vik"] = str_db[37] + "," + str_db[38] + "," + str_db[39] + "," + str_db[40] + "," + str_db[41];
                            dict_db["zag_stat_code"] = str_db[42] + "," + str_db[43] + "," + str_db[44] + "," + str_db[45] + "," + str_db[46];
                            dict_db["zag_stat_name"] = "";
                            dict_db["code_status"] = str_db[47] + "," + str_db[48] + "," + str_db[49] + "," + str_db[50] + "," + str_db[51];
                            dict_db["name_status"] = "";
                            dict_db["code_moment"] = str_db[52] + "," + str_db[53] + "," + str_db[54] + "," + str_db[55] + "," + str_db[56];
                            dict_db["moment"] = "";
                            dict_db["code_umovi"] = str_db[57] + "," + str_db[58] + "," + str_db[59] + "," + str_db[60] + "," + str_db[61];
                            dict_db["name_umovi"] = "";
                            dict_db["travm"] = str_db[62];
                            dict_db["travm_ditei"] = str_db[63];
                            dict_db["travm_fire"] = str_db[64];
                            dict_db["pramiy"] = str_db[65];
                            dict_db["pobichniy"] = str_db[66];
                            dict_db["zn_bud"] = str_db[67];
                            dict_db["posh_bud"] = str_db[68];
                            dict_db["zn_tehnika"] = str_db[69];
                            dict_db["posh_tehnika"] = str_db[70];
                            dict_db["zn_zerno"] = str_db[71];
                            dict_db["zn_koreni"] = str_db[72];
                            dict_db["zn_valki"] = str_db[73];
                            dict_db["zn_korm"] = str_db[74];
                            dict_db["zn_torf"] = str_db[75];
                            dict_db["posh_torf"] = str_db[76];
                            dict_db["zag_tvarin"] = str_db[77];
                            dict_db["zag_ptici"] = str_db[78];
                            dict_db["dop_info"] = str_db[79];

                            dict_db["vr_ludei"] = str_db[80];
                            dict_db["vr_ditei"] = str_db[81];
                            dict_db["vr_tvarin"] = str_db[82];
                            dict_db["vr_ptici"] = str_db[83];
                            dict_db["vr_bud"] = str_db[84];
                            dict_db["vr_tehnika"] = str_db[85];
                            dict_db["vr_zerno"] = str_db[86];
                            dict_db["vr_koreni"] = str_db[87];
                            dict_db["vr_valki"] = str_db[88];
                            dict_db["vr_korm"] = str_db[89];
                            dict_db["vr_torf"] = str_db[90];
                            dict_db["vr_dop"] = str_db[91];
                            dict_db["vr_mat"] = str_db[92];

                            dict_db["data_pov"] = str_db[93];
                            dict_db["time_pov"] = str_db[94];
                            dict_db["time_pributa"] = str_db[95];
                            dict_db["code_fire_likvid"] = str_db[96];
                            dict_db["name_fire_likvid"] = "";
                            dict_db["data_lokal"] = str_db[97];
                            dict_db["time_lokal"] = str_db[98];
                            dict_db["data_likvid"] = str_db[99];
                            dict_db["time_likvid"] = str_db[100];
                            dict_db["code_umovi_posh"] = str_db[101] + "," + str_db[102] + "," + str_db[103] + "," + str_db[104] + "," + str_db[105];
                            dict_db["name_umovi_posh"] = "";
                            dict_db["code_umovi_uskl"] = str_db[106] + "," + str_db[107] + "," + str_db[108] + "," + str_db[109] + "," + str_db[110];
                            dict_db["name_umovi_uskl"] = "";
                            dict_db["code_spz"] = str_db[111];
                            dict_db["name_spz"] = "";
                            dict_db["code_system"] = str_db[112] + "," + str_db[113] + "," + str_db[114] + "," + str_db[115] + "," + str_db[116];
                            dict_db["name_system"] = "";
                            dict_db["code_dii"] = str_db[117] + "," + str_db[118] + "," + str_db[119] + "," + str_db[120] + "," + str_db[121];
                            dict_db["name_dii"] = "";

                            dict_db["code_uch"] = str_db[122] + "," + str_db[123] + "," + str_db[124] + "," + str_db[125] + "," + str_db[126];
                            dict_db["name_uch"] = "";
                            dict_db["kilk_uch"] = str_db[127] + "," + str_db[128] + "," + str_db[129] + "," + str_db[130] + "," + str_db[131];
                            dict_db["code_fireauto"] = str_db[132] + "," + str_db[133] + "," + str_db[134] + "," + str_db[135] + "," + str_db[136];
                            dict_db["name_fireauto"] = "";
                            dict_db["kilk_auto"] = str_db[137] + "," + str_db[138] + "," + str_db[139] + "," + str_db[140] + "," + str_db[141];
                            dict_db["code_firestvol"] = str_db[142] + "," + str_db[143] + "," + str_db[144];
                            dict_db["name_firestvol"] = "";
                            dict_db["kilk_firestvol"] = str_db[145] + "," + str_db[146] + "," + str_db[147];
                            dict_db["code_rechovini"] = str_db[148] + "," + str_db[149] + "," + str_db[150];
                            dict_db["name_rechovini"] = "";
                            dict_db["code_pervini"] = str_db[151] + "," + str_db[152] + "," + str_db[153];
                            dict_db["name_pervini"] = "";
                            dict_db["code_djerela"] = str_db[154] + "," + str_db[155] + "," + str_db[156];
                            dict_db["name_djerela"] = "";
                            dict_db["vikr_gds"] = str_db[157];
                            dict_db["kilk_gds"] = str_db[158];
                            dict_db["time_gds"] = str_db[159];

                            dict_db["data_perevirki"] = str_db[160];
                            dict_db["code_perevirka"] = str_db[161];
                            dict_db["name_perevirka"] = "";
                            dict_db["code_dialnist"] = str_db[162];
                            dict_db["name_dialnist"] = "";
                            dict_db["code_zahodi"] = str_db[163] + "," + str_db[164];
                            dict_db["name_zahodi"] = "";
                            dict_db["number_kku"] = str_db[165];
                            dict_db["data_zapovnenya"] = str_db[166];
                            dict_db["pid_osibi"] = str_db[167];
                        }
                        string select_code = "SELECT `number_cartka` FROM kartka_obliku WHERE code_raion=" + dict_db["code_raion"] + " AND number_cartka=" + dict_db["number_cartka"] + " AND main_dop=" + dict_db["main_dop"];
                        cmd_db = new SQLiteCommand(select_code, con_db);
                        rdr = cmd_db.ExecuteReader();

                        while (rdr.Read())
                        {
                            content_val = rdr[0].ToString();
                        // dict.Add(code, rdr[0].ToString());

                        }
                  
                        if (content_val == dict_db["number_cartka"].ToString())
                        {
                            if (IsMessageShow)
                            {
                                res = MessageBox.Show("Картки з такими номерами вже існують. Ви впевнені що хочете оновити ці картки", "Увага!", MessageBoxButtons.YesNo);
                                IsMessageShow = false;
                            }
                            if (res == DialogResult.Yes)
                            {
                                flag_edit = true;
                                SaveInDb save = new SaveInDb(dict_db, flag_edit);
                            flag_good = true;
                            }
                        }
                        else
                        {
                            flag_edit = false;
                            SaveInDb save = new SaveInDb(dict_db, flag_edit);
                        flag_good = true;
                        }
                       
                   
                    // dict_db.Clear();
                    }
                   
                    if (flag_good)
                    {
                        MessageBox.Show("Файл успішно завантажений");
                       // flag_good = false;
                    }
                    

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
           
        }

        private void maskedTextBox1_Leave(object sender, EventArgs e)
        {
            try
            {
                bool val = DateTime.TryParse(maskedTextBox1.Text, out DateTime res);
                if (val)
                {
                    dateTimePicker1.Value = res;
                }
                else
                {
                    dateTimePicker1.Value = DateTime.Parse("01.01.1900");
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
        }

        private void maskedTextBox2_Leave(object sender, EventArgs e)
        {
            //dateTimePicker2.Value = DateTime.Parse(maskedTextBox2.Text);
            try
            {
                bool val = DateTime.TryParse(maskedTextBox2.Text, out DateTime res);
                if (val)
                {
                    dateTimePicker2.Value = res;
                }
                else
                {
                    dateTimePicker2.Value = DateTime.Parse("01.01.1900");
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
        }

        private void maskedTextBox3_Leave(object sender, EventArgs e)
        {
            // dateTimePicker6.Value = DateTime.Parse(maskedTextBox3.Text);
            try
            {
                bool val = DateTime.TryParse(maskedTextBox3.Text, out DateTime res);
                if (val)
                {
                    dateTimePicker6.Value = res;
                }
                else
                {
                    dateTimePicker6.Value = DateTime.Parse("01.01.1900");
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
        }

        private void maskedTextBox4_Leave(object sender, EventArgs e)
        {
            // dateTimePicker8.Value = DateTime.Parse(maskedTextBox4.Text);
            try
            {
                bool val = DateTime.TryParse(maskedTextBox4.Text, out DateTime res);
                if (val)
                {
                    dateTimePicker8.Value = res;
                }
                else
                {
                    dateTimePicker8.Value = DateTime.Parse("01.01.1900");
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
        }

        private void maskedTextBox5_Leave(object sender, EventArgs e)
        {
            // dateTimePicker9.Value = DateTime.Parse(maskedTextBox5.Text);
            try
            {
                bool val = DateTime.TryParse(maskedTextBox5.Text, out DateTime res);
                if (val)
                {
                    dateTimePicker9.Value = res;
                }
                else
                {
                    dateTimePicker9.Value = DateTime.Parse("01.01.1900");
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
        }

        private void maskedTextBox6_Leave(object sender, EventArgs e)
        {
            //  dateTimePicker10.Value = DateTime.Parse(maskedTextBox6.Text);
            try
            {
                bool val = DateTime.TryParse(maskedTextBox6.Text, out DateTime res);
                if (val)
                {
                    dateTimePicker10.Value = res;
                }
                else
                {
                    dateTimePicker10.Value = DateTime.Parse("01.01.1900");
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            maskedTextBox2.Text = dateTimePicker2.Value.ToString();
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            maskedTextBox3.Text = dateTimePicker6.Value.ToString();
        }

        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            maskedTextBox4.Text = dateTimePicker8.Value.ToString();
        }

        private void dateTimePicker9_ValueChanged(object sender, EventArgs e)
        {
            maskedTextBox5.Text = dateTimePicker9.Value.ToString();
        }

        private void dateTimePicker10_ValueChanged(object sender, EventArgs e)
        {
            maskedTextBox6.Text = dateTimePicker10.Value.ToString();
        }

        private void maskedTextBox7_Leave(object sender, EventArgs e)
        {
            bool val =DateTime.TryParse(maskedTextBox7.Text,out DateTime res);
            if (val)
            {
                dateTimePicker3.Value = res;
            }
            else
            {
                dateTimePicker3.Value = DateTime.Parse("00:00");
            }
        }

        private void maskedTextBox8_Leave(object sender, EventArgs e)
        {
            bool val = DateTime.TryParse(maskedTextBox8.Text, out DateTime res);
            if (val)
            {
                dateTimePicker4.Value = res;
            }
            else
            {
                dateTimePicker4.Value = DateTime.Parse("00:00");
            }
        }

        private void maskedTextBox9_Leave(object sender, EventArgs e)
        {
            bool val = DateTime.TryParse(maskedTextBox9.Text, out DateTime res);
            if (val)
            {
                dateTimePicker5.Value = res;
            }
            else
            {
                dateTimePicker5.Value = DateTime.Parse("00:00");
            }
        }

        private void maskedTextBox10_Leave(object sender, EventArgs e)
        {
            bool val = DateTime.TryParse(maskedTextBox10.Text, out DateTime res);
            if (val)
            {
                dateTimePicker7.Value = res;
            }
            else
            {
                dateTimePicker7.Value = DateTime.Parse("00:00");
            }
        }

        private void maskedTextBox1_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата виникнення пожежі";
            bool state = false;
            int code;
            state = int.TryParse(textBox3.Text, out int res);

            string content_val = "";

            string select_code = "SELECT `number_cartka` FROM kartka_obliku WHERE code_raion=" + textBox1.Text + " AND number_cartka=" + textBox2.Text + " AND main_dop=" + textBox3.Text;
            //  connection();
            try
            {
                //if (con_db.State == ConnectionState.Open)
                //{
                cmd_db = new SQLiteCommand(select_code, con_db);
                rdr = cmd_db.ExecuteReader();

                while (rdr.Read())
                {
                    content_val = rdr[0].ToString();
                    // dict.Add(code, rdr[0].ToString());
                }
                // }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            if (state)
            {
                if (content_val != "" && flag_edit == false)
                {
                    MessageBox.Show("Картка під таким номером вже існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox2.BackColor = Color.Firebrick;
                    textBox2.ForeColor = Color.White;
                    textBox2.Select();
                    textBox2.ScrollToCaret();
                }
                else
                {
                    code = int.Parse(textBox3.Text);
                    if (code > 0 && code <= 9)
                    {
                        textBox3.ForeColor = Color.Firebrick;
                    }
                    if (code == 0)
                    {
                        textBox3.ForeColor = Color.Black;
                    }
                    if (code > 9)
                    {
                        MessageBox.Show("В це поле можно внести тільки коди від 0 до 9", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        textBox3.Select();
                        textBox3.ScrollToCaret();
                    }
                }

            }
            else
            {
                MessageBox.Show("В це поле можно внести тільки цифри", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox3.Text = "0";
                textBox2.Select();
                textBox2.ScrollToCaret();
            }
        }

        private void maskedTextBox2_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата повідомлення про пожежу";
        }

        private void maskedTextBox7_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "час повідомлення про пожежу";
            if (dateTimePicker2.Value.ToString() == "01.01.1900 0:00:00")
            {
                MessageBox.Show("Заповніть дату повідомлення про пожежу");
                dateTimePicker2.Select();
            }
        }

        private void maskedTextBox8_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Час прибуття 1 - го підрозділу";
        }

        private void maskedTextBox3_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата локалізації пожежі";
            string code = CodeAllContent("info_fire", "code_fire_likvid", "code_fire_likvid", textBox102.Text);
            //if (textBox81.Text != "")
            if (textBox102.Text == "" && textBox3.Text == "0")
            {
                MessageBox.Show("Це поле обов'язкове для заповнення", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox102.Select();
                textBox102.ScrollToCaret();
            }
            else if (code == "")
            {
                MessageBox.Show("Такого кода не існує!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox102.Text = "";
                textBox102.BackColor = Color.Firebrick;
                textBox102.Select();
                textBox102.ScrollToCaret();
            }
        }

        private void maskedTextBox9_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Час локалізації пожежі ";
        }

        private void maskedTextBox4_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата ліквідації пожежі";
        }

        private void maskedTextBox10_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Час ліквідації пожежі";
            if (dateTimePicker8.Value.ToString() == "01.01.1900 0:00:00")
            {
                MessageBox.Show("Заповніть дату ліквідації пожежі");
                dateTimePicker8.Select();
            }
        }

        private void maskedTextBox5_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата останньої перевірки";
        }

        private void maskedTextBox6_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Дата заповнення Картки обліку пожежі";
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            //Console.WriteLine(dateTimePicker3.Value.ToShortTimeString());
            if (dateTimePicker3.Value.ToShortTimeString() != "0:00")
            {
                maskedTextBox7.Text = dateTimePicker3.Value.ToShortTimeString();
            }
            else
            {
                maskedTextBox7.Text = "";
            }
           
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker4.Value.ToShortTimeString() != "0:00")
            {
                maskedTextBox8.Text = dateTimePicker4.Value.ToShortTimeString();
            }
            else
            {
                maskedTextBox8.Text = "";
            }
           // maskedTextBox8.Text = dateTimePicker4.Value.ToShortTimeString();
        }

        private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker5.Value.ToShortTimeString() != "0:00")
            {
                maskedTextBox9.Text = dateTimePicker5.Value.ToShortTimeString();
            }
            else
            {
                maskedTextBox9.Text = "";
            }
            //maskedTextBox9.Text = dateTimePicker5.Value.ToShortTimeString();
        }

        private void dateTimePicker7_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker7.Value.ToShortTimeString() != "0:00")
            {
                maskedTextBox10.Text = dateTimePicker7.Value.ToShortTimeString();
            }
            else
            {
                maskedTextBox10.Text = "";
            }
            //maskedTextBox10.Text = dateTimePicker7.Value.ToShortTimeString();
        }

        private void видаленняКартокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel5.Visible = true;
            button16.Visible = true;
            button17.Text = "Видалити картку";
            radioButton1.Checked = false;
            radioButton2.Checked = false;

            cmd_db = new SQLiteCommand("SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku", con_db);
            rdr = cmd_db.ExecuteReader();
            listBox3.Items.Clear();

            while (rdr.Read())
            {
                // region_items.Add(rdr[1].ToString());
                listBox3.Items.Add("номер картки " +"("+rdr[3].ToString()+" ) "+ rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
            }
            label108.Text = listBox3.Items.Count.ToString();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            string zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku";

   
            if (maskedTextBox11.Text != "  .  .")
            {
                maskedTextBox12.Text = "";
                maskedTextBox13.Text = "";
                textBox160.Text = "";
                textBox161.Text = "";
                textBox162.Text = "";
                textBox165.Text = "";
                textBox163.Text = "";
                textBox164.Text = "";
                textBox166.Text = "";

                zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku WHERE date_viniknenya='" + maskedTextBox11.Text+"'";
                cmd_db = new SQLiteCommand(zapros, con_db);
                rdr = cmd_db.ExecuteReader();


                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    if(button17.Text!= "Видалити картку")
                    {
                        listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    else
                    {
                        listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    
                }
            }
            if (maskedTextBox12.Text != "  .  ." && maskedTextBox13.Text != "  .  .")
            {
                maskedTextBox11.Text = "";
                textBox160.Text = "";
                textBox161.Text = "";
                textBox162.Text = "";
                textBox165.Text = "";
                textBox163.Text = "";
                textBox164.Text = "";
                textBox166.Text = "";

                cmd_db = new SQLiteCommand(zapros, con_db);
                rdr = cmd_db.ExecuteReader();
                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    if(DateTime.Parse(rdr[2].ToString())>=DateTime.Parse(maskedTextBox12.Text) && DateTime.Parse(rdr[2].ToString()) <= DateTime.Parse(maskedTextBox13.Text))
                    {
                        if (button17.Text != "Видалити картку")
                        {
                            listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                        }
                        else
                        {
                            listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                        }
                    }
                    
                }
            }
            if (textBox160.Text != "")
            {
                maskedTextBox12.Text = "";
                maskedTextBox13.Text = "";
                maskedTextBox11.Text = "";
                textBox161.Text = "";
                textBox162.Text = "";
                textBox165.Text = "";
                textBox163.Text = "";
                textBox164.Text = "";
                textBox166.Text = "";

                zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku WHERE number_cartka='" + textBox160.Text + "'";
                cmd_db = new SQLiteCommand(zapros, con_db);
                rdr = cmd_db.ExecuteReader();


                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    if (button17.Text != "Видалити картку")
                    {
                        listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    else
                    {
                        listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                }
            }
            if(textBox161.Text!="" && textBox162.Text != "")
            {
                maskedTextBox12.Text = "";
                maskedTextBox13.Text = "";
                maskedTextBox11.Text = "";
                textBox160.Text = "";
                textBox165.Text = "";
                textBox163.Text = "";
                textBox164.Text = "";
                textBox166.Text = "";

                zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku WHERE number_cartka >='" + textBox161.Text + "' AND number_cartka <='" + textBox162.Text +"'";
                cmd_db = new SQLiteCommand(zapros, con_db);
                rdr = cmd_db.ExecuteReader();


                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    if (button17.Text != "Видалити картку")
                    {
                        listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    else
                    {
                        listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                }
            }
            if (textBox165.Text != "")
            {
                maskedTextBox12.Text = "";
                maskedTextBox13.Text = "";
                maskedTextBox11.Text = "";
                textBox161.Text = "";
                textBox162.Text = "";
                textBox160.Text = "";
                textBox163.Text = "";
                textBox164.Text = "";
                textBox166.Text = "";

                zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku WHERE main_dop='" + textBox165.Text + "'";
                cmd_db = new SQLiteCommand(zapros, con_db);
                rdr = cmd_db.ExecuteReader();


                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    if (button17.Text != "Видалити картку")
                    {
                        listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    else
                    {
                        listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                }
            }
            if (textBox164.Text != "" && textBox163.Text != "")
            {
                maskedTextBox12.Text = "";
                maskedTextBox13.Text = "";
                maskedTextBox11.Text = "";
                textBox161.Text = "";
                textBox162.Text = "";
                textBox160.Text = "";
                textBox165.Text = "";
                textBox166.Text = "";

                zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku WHERE main_dop >='" + textBox164.Text + "' AND main_dop <='" + textBox163.Text +"'";
                cmd_db = new SQLiteCommand(zapros, con_db);
                rdr = cmd_db.ExecuteReader();


                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    if (button17.Text != "Видалити картку")
                    {
                        listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    else
                    {
                        listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                }
            }
            if (textBox166.Text != "")
            {
                maskedTextBox12.Text = "";
                maskedTextBox13.Text = "";
                maskedTextBox11.Text = "";
                textBox161.Text = "";
                textBox162.Text = "";
                textBox160.Text = "";
                textBox165.Text = "";
                textBox163.Text = "";
                textBox164.Text = "";

                zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku WHERE code_raion ='" + textBox166.Text + "'";
                cmd_db = new SQLiteCommand(zapros, con_db);
                rdr = cmd_db.ExecuteReader();


                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    if (button17.Text != "Видалити картку")
                    {
                        listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    else
                    {
                        listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                }
            }






            if (maskedTextBox11.Text == "  .  ." && maskedTextBox12.Text == "  .  ." && maskedTextBox13.Text == "  .  ." && textBox160.Text=="" && textBox161.Text == "" && textBox162.Text == "" && textBox165.Text == "" && textBox164.Text == "" && textBox163.Text == "" && textBox166.Text == "")
            {
                radioButton1.Checked = false;
                radioButton2.Checked = false;

                zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku";
                cmd_db = new SQLiteCommand(zapros, con_db);
                   rdr = cmd_db.ExecuteReader();


                   while (rdr.Read())
                   {
                    // region_items.Add(rdr[1].ToString());
                    if (button17.Text != "Видалити картку")
                    {
                        listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    else
                    {
                        listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                }
            }

            label108.Text = listBox3.Items.Count.ToString();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            if (radioButton1.Checked)
            {
              string  zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku WHERE code_raion ='" + textBox166.Text + "' AND main_dop=0";

                cmd_db = new SQLiteCommand(zapros, con_db);
                rdr = cmd_db.ExecuteReader();


                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    if (button17.Text != "Видалити картку")
                    {
                        listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    else
                    {
                        listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                }
            }
            label108.Text = listBox3.Items.Count.ToString();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            if (radioButton2.Checked)
            {
                string zapros = "SELECT `number_cartka`,`main_dop`,`date_viniknenya`,`id_kartki` from kartka_obliku WHERE code_raion ='" + textBox166.Text + "' AND main_dop!=0";

                cmd_db = new SQLiteCommand(zapros, con_db);
                rdr = cmd_db.ExecuteReader();


                while (rdr.Read())
                {
                    // region_items.Add(rdr[1].ToString());
                    if (button17.Text != "Видалити картку")
                    {
                        listBox3.Items.Add("номер картки " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                    else
                    {
                        listBox3.Items.Add("номер картки " + "(" + rdr[3].ToString() + " ) " + rdr[0].ToString() + "," + rdr[1].ToString() + " дата виникнення " + "(" + rdr[2].ToString() + ")");
                    }
                }
            }
            label108.Text = listBox3.Items.Count.ToString();
        }

        private void textBox166_TextChanged(object sender, EventArgs e)
        {
            if (textBox166.Text == "")
            {
                radioButton1.Checked = false;
                radioButton2.Checked = false;
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            maskedTextBox12.Text = "";
            maskedTextBox13.Text = "";
            maskedTextBox11.Text = "";
            textBox161.Text = "";
            textBox162.Text = "";
            textBox160.Text = "";
            textBox165.Text = "";
            textBox163.Text = "";
            textBox164.Text = "";
            textBox166.Text = "";
            radioButton1.Checked = false;
            radioButton2.Checked = false;
        }
    }
}
       