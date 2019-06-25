using IniParser;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AppForm = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;

using Excel = Microsoft.Office.Interop.Excel;

namespace StrainMenuCreator.Ounces
{
    public partial class OunceForm : Form
    {
        public string TemplateFile { get; set; }

        public OunceForm()
        {
            InitializeComponent();
        }

        private static string oauth => File.Exists(@"Z:\Slack Bot\SlackBot_Auth.txt") ? File.ReadAllText(@"Z:\Slack Bot\SlackBot_Auth.txt") : "https://hooks.slack.com/services/T00000000/B00000000/XXXXXXXXXXXXXXXXXXXXXXXX";
        SlackClient Bot = new SlackClient(oauth);

        private List<String> Names = new List<String> { };
        private List<String> OuncePrice = new List<String> { };
        private List<String> HalfOunce = new List<String> { };

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!File.Exists("Ounces.ini"))
                RepairOunces();
            dataGridView1.DataSource = Data();
            RefreshGrid();
        }

        private void RepairOunces() {
            using (StreamWriter sw = File.CreateText("Ounces.ini"))
            {
                sw.WriteLine("[Settings]");
                sw.WriteLine("Total = 1");
                sw.WriteLine("");
                sw.WriteLine("[Name]");
                sw.WriteLine("1 = NA");
                sw.WriteLine("");
                sw.WriteLine("[O_Price]");
                sw.WriteLine("1 = NA");
                sw.WriteLine("");
                sw.WriteLine("[H_Price]");
                sw.WriteLine("1 = NA");
            }
        }

        private void RefreshGrid()
        {
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[0].ToString());

            string[] Names = Values.ToArray();
            Values.Clear();
            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

            var parser = new FileIniDataParser();
            var data = parser.ReadFile(AppForm.StartupPath + @"\Settings.ini");
            menuStart.Value = decimal.Parse(data["Settings"]["menuStart"]);

            data = parser.ReadFile(AppForm.StartupPath + @"\Ounces.ini");
            numericUpDown1.Value = Int32.Parse(data["Settings"]["Total"]);
        }

        private DataTable Data()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Ounce");
            dt.Columns.Add("Half");

            var parser = new FileIniDataParser();
            var data = parser.ReadFile(AppForm.StartupPath + @"\Ounces.ini");
            Console.WriteLine("List loaded... Reading Settings>Total");
            int Total = Int32.Parse(data["Settings"]["Total"]);
            Console.WriteLine("Total = " + Total.ToString());
            for (int i = 0; i < Total; i++)
            {
                DataRow dr = dt.NewRow();
                Console.WriteLine("Reading " + i + "/" + (i + 1));

                dr["Name"] = data["Name"][(i + 1).ToString()];
                dr["Ounce"] = data["O_Price"][(i + 1).ToString()];
                dr["Half"] = data["H_Price"][(i + 1).ToString()];
                dt.Rows.Add(dr);
            }

            return dt;
        }

        private const int ColumnBase = 26;
        private const int DigitMax = 7; // ceil(log26(Int32.Max))
        private const string Digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static string GetLetter(int index)
        {
            if (index <= 0)
            {
                throw new IndexOutOfRangeException("index must be a positive number");
            }

            if (index <= ColumnBase)
            {
                return Digits[index - 1].ToString();
            }

            var sb = new StringBuilder().Append(' ', DigitMax);
            var current = index;
            var offset = DigitMax;
            while (current > 0)
            {
                sb[--offset] = Digits[--current % ColumnBase];
                current /= ColumnBase;
            }
            return sb.ToString(offset, DigitMax - offset);
        }

        public string Merge(int Num_Letter, int Num)
        {
            return (GetLetter(Num_Letter) + Num.ToString());
        }

        public int GetNum(string Range)
        {
            char c = char.Parse(Range.Substring(0, 1).ToLower());
            return char.ToUpper(c) - 63;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Text = "Saving...";
            Cursor.Current = Cursors.WaitCursor;
            GenPremade(AppForm.StartupPath + @"\Ounces.ini");
            EditTemplate(AppForm.StartupPath + @"\Template_36.xlsx");
            EditTemplate(AppForm.StartupPath + @"\Template_40.xlsx");
            EditTemplate(AppForm.StartupPath + @"\Template_44.xlsx");
            button1.Text = "Save";
            Cursor.Current = Cursors.Default;
            if (!ShouldCreate)
                MessageBox.Show("Done! Just create your menu like normal from the previous menu, in order to see your changes.");
            else
            {
                Bot.PostMessage("Ounces have been edited and written to the excel files. Please create a flower menu as you normally would to see the desired changes.", username: "botmenu", channel: "#menu_updates");
                this.Close();
            }
        }

        public void GenPremade(string Premade)
        {
            Console.WriteLine("Generating Premade...");
            using (var sw = File.CreateText(Premade)) { }

            var parser = new FileIniDataParser();
            var data = parser.ReadFile(Premade);

            #region Create lists from DataGridView

            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[0].ToString());
            }

            string[] Names = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[1].ToString());
            }

            string[] Costs = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[2].ToString());
            }

            string[] Types = Values.ToArray();
            Values.Clear();

            #endregion Create lists from DataGridView

            int i = 1;
            int Total = Names.Length;
            foreach (var name in Names)
            {
                if (name != "/r" && name != "/n" && name != "")
                {
                    data["Name"][i.ToString()] = name;
                    data["O_Price"][i.ToString()] = Costs[i - 1];
                    data["H_Price"][i.ToString()] = Types[i - 1];
                    i++;
                }
                else
                {
                    Total--;
                }
            }
            data["Settings"]["Total"] = Total.ToString();
            parser.WriteFile(Premade, data);
        }

        private void EditTemplate(string templateFile)
        {
            try
            {
                int menuOffset = (Int32)menuStart.Value;
                Console.WriteLine("Starting Excel edit...");
                Excel.Application excel = new Excel.Application();
                Excel.Workbook wkb = excel.Workbooks.Open(templateFile);
                Excel.Worksheet sheet = wkb.Worksheets[1] as Excel.Worksheet;
                Console.WriteLine("Opened Excel file, starting edit...");
                Range row = sheet.Rows.Cells[2, 1];

                Console.WriteLine("Clearing old ounces");
                for (int c = GetNum("N") - 1; c < GetNum("N") + 3; c++)
                {
                    for (int r = menuOffset; r < numericUpDown1.Value + menuOffset; r++)
                    {
                        Console.WriteLine("c{0} : r{1}", c, r);
                        row = sheet.Rows.Cells[r, c];
                        row.Value = "";
                    }
                }

                //.Range["A1:L33"]
                int Letter = GetNum("N") - 1;
                int Num = menuOffset;

                if (checkBox1.Checked)
                {
                    row = sheet.Rows.Cells[4, Letter];
                    row.Value = "Name";
                    row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    row.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    row.Font.Size = 30;
                    row.Font.Bold = true;
                    row.Font.Italic = true;
                    row.Font.Color = Color.White;

                    row = sheet.Rows.Cells[4, Letter + 1];
                    row.Value = "1oz";
                    row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    row.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    row.Font.Size = 30;
                    row.Font.Bold = true;
                    row.Font.Italic = true;
                    row.Font.Color = Color.White;

                    row = sheet.Rows.Cells[4, Letter + 2];
                    row.Value = "1/2oz";
                    row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    row.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    row.Font.Size = 30;
                    row.Font.Bold = true;
                    row.Font.Italic = true;
                    row.Font.Color = Color.White;

                    row = sheet.Rows.Cells[3, Letter];
                    row.Value = "";
                    row = sheet.Rows.Cells[3, Letter + 1];
                    row.Value = "";
                    row = sheet.Rows.Cells[3, Letter + 2];
                    row.Value = "";
                }

                #region Create lists from DataGridView

                List<string> Values = new List<string>();

                DataTable dt = (DataTable)dataGridView1.DataSource;
                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[0].ToString());
                }

                string[] Names = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[1].ToString());
                }

                string[] Costs = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[2].ToString());
                }

                string[] Types = Values.ToArray();
                Values.Clear();

                #endregion Create lists from DataGridView

                int i = 0;
                foreach (var name in Names)
                {
                    if (name != "" && name != "/n" && name != "/r")
                    {
                        //Flower Name
                        row = sheet.Rows.Cells[Num, Letter];
                        row.Value = name;
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        row.Font.Size = 28;
                        row.Font.Bold = true;
                        row.Font.Color = Color.Black;

                        //Ounce Cost
                        row = sheet.Rows.Cells[Num, Letter + 1];
                        decimal Cost_ = 0;
                        //NA check
                        if (Costs[i].ToUpper() != "NA")
                        {
                            if (Costs[i].Substring(0, 1) == "0")
                            {
                                Cost_ = Int32.Parse(Costs[i].Substring(1));
                            }
                            else
                            {
                                Cost_ = Int32.Parse(Costs[i]);
                            }

                            if (Cost_ > 99)
                            {
                                row.NumberFormat = "$###.00";
                            }
                            else
                            {
                                row.NumberFormat = "$##.00";
                            }

                            row.Value = Cost_;
                        }
                        else
                        {
                            row.Value = "NA";
                        }
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        row.Font.Size = 28;
                        row.Font.Bold = true;
                        row.Font.Color = Color.Black;

                        //Half Cost
                        row = sheet.Rows.Cells[Num, Letter + 2];
                        Cost_ = 0;
                        //NA check
                        if (Types[i].ToUpper() != "NA")
                        {
                            if (Types[i].Substring(0, 1) == "0")
                            {
                                Cost_ = Int32.Parse(Types[i].Substring(1));
                            }
                            else
                            {
                                Cost_ = Int32.Parse(Types[i]);
                            }

                            if (Cost_ > 99)
                            {
                                row.NumberFormat = "$###.00";
                            }
                            else
                            {
                                row.NumberFormat = "$##.00";
                            }

                            row.Value = Cost_;
                        }
                        else
                        {
                            row.Value = "NA";
                        }
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        row.Font.Size = 28;
                        row.Font.Bold = true;
                        row.Font.Color = Color.Black;
                    }
                    i++;
                    Num++;
                }

                Console.WriteLine("Done editing saving...");
                excel.Application.ActiveWorkbook.Save();
                object objFalse = false;
                wkb.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
                Console.WriteLine("Done");
            }
            catch (Exception e) { MessageBox.Show(e.ToString()); }
        }

        private void UpdateDeleteBox()
        {
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[0].ToString());
            }

            string[] Names = Values.ToArray();
            Values.Clear();

            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        }

        //Delete
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                #region Create lists from DataGridView

                List<string> Values = new List<string>();
                DataTable dt = (DataTable)dataGridView1.DataSource;
                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[0].ToString());
                }

                string[] Names = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[1].ToString());
                }

                string[] Types = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[2].ToString());
                }

                string[] Costs = Values.ToArray();
                Values.Clear();

                #endregion Create lists from DataGridView

                int i = Array.IndexOf(Names, Delete_Box.Text);

                var list = new List<string>(Names);
                list.RemoveAt(i);
                Names = list.ToArray();

                list = new List<string>(Costs);
                list.RemoveAt(i);
                Costs = list.ToArray();

                list = new List<string>(Types);
                list.RemoveAt(i);
                Types = list.ToArray();

                dt.Rows.RemoveAt(i);
                dataGridView1.DataSource = dt;
                Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error removing the selected strain!" + System.Environment.NewLine + ex.ToString());
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void menuStart_ValueChanged(object sender, EventArgs e)
        {
            var parser = new FileIniDataParser();
            var data = parser.ReadFile(AppForm.StartupPath + @"\Settings.ini");
            data["Settings"]["menuStart"] = menuStart.Value.ToString();
        }


        public bool ShouldCreate = false;
        private void OunceForm_Shown(object sender, EventArgs e)
        {
            if (ShouldCreate)
                button1_Click(sender, e);
        }
    }
}