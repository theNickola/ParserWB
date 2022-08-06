using System;
using System.IO;
using System.Net;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ParserWB
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label1.Text = "please wait ...";
            button1.Enabled = false;
            Excel.Application excelApp = new Excel.Application();
            try
            {
                const string testFile = "Keys.txt";
                StreamReader sr = File.OpenText(testFile);
                string[] mas = sr.ReadToEnd().Split();

                HttpWebRequest req = null;
                HttpWebResponse resp;

                excelApp = new Excel.Application();
                excelApp.Workbooks.Add();
                excelApp.ActiveSheet.Delete();
                excelApp.ActiveSheet.Delete();
                var oneActive = true;
                Excel._Worksheet workSheet = excelApp.ActiveSheet;



                for (int j = mas.Length-1 ; j >= 0; j--)
                {
                    if (mas[j] != "")
                    {
                        req = (HttpWebRequest)WebRequest.Create("https://search.wb.ru/exactmatch/ru/common/v4/search?appType=1&couponsGeo=12,3,18,15,21&curr=rub&dest=-1029256,-102269,-2162196,-1257786&emp=0&lang=ru&locale=ru&pricemarginCoeff=1.0&query={" + mas[j] + "}&reg=0&regions=68,64,83,4,38,80,33,70,82,86,75,30,69,22,66,31,40,1,48,71&resultset=catalog&sort=popular&spp=0&suppressSpellcheck=false");
                        resp = (HttpWebResponse)req.GetResponse();
                        Stream istrm = resp.GetResponseStream();

                        StreamReader sr2 = new StreamReader(istrm);
                        var json = sr2.ReadToEnd();

                        if(!oneActive)
                            workSheet = (Excel.Worksheet)excelApp.Worksheets.Add();
                        oneActive = false;

                        workSheet.Name = mas[j];
                        workSheet.Cells[1, 1] = "Title";
                        workSheet.Cells[1, 2] = "Brand";
                        workSheet.Cells[1, 3] = "Id";
                        workSheet.Cells[1, 4] = "Feedbacks";
                        workSheet.Cells[1, 5] = "Price";


                        if (!string.IsNullOrEmpty(json))
                        {
                            ProductWB result = Newtonsoft.Json.JsonConvert.DeserializeObject<ProductWB>(json);
                            for (int i = 0; i < result.data.products.Length; i++)
                            {
                                workSheet.Cells[i + 2, 1] = result.data.products[i].brand + " / " + get_title(result.data.products[i].id);
                                workSheet.Cells[i + 2, 2] = result.data.products[i].brand;
                                workSheet.Cells[i + 2, 3] = result.data.products[i].id;
                                workSheet.Cells[i + 2, 4] = result.data.products[i].feedbacks;
                                workSheet.Cells[i + 2, 5] = result.data.products[i].priceU / 100;
                            }
                        }
                        
                    }
                }
                excelApp.ActiveSheet.Delete();
                workSheet.SaveAs(string.Format(@"{0}\ResultParsing.xlsx", Environment.CurrentDirectory));
                excelApp.Quit();
            }
            catch(Exception ex)
            {
                excelApp.Quit();
                MessageBox.Show(ex.Message);
            }
            label1.Text = "press the \"BEGIN\" button to start the process";
            button1.Enabled = true;
        }

        private string get_title(int id)
        {
            string json = null;
            int first_numb=0, last_numb=0;
        
            try
            {
                HttpWebRequest req = null;
                HttpWebResponse resp;
                req = (HttpWebRequest)WebRequest.Create($"https://wbx-content-v2.wbstatic.net/ru/{id}.json");
                resp = (HttpWebResponse)req.GetResponse();
                Stream istrm = resp.GetResponseStream();
                StreamReader sr2 = new StreamReader(istrm);
                json = sr2.ReadToEnd();
                string first = "\"imt_name\":\"";
                string last = "\",\"subj_name\":\"";
                first_numb = json.IndexOf(first) + first.Length;
                last_numb = json.IndexOf(last) - first_numb;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return json.Substring(first_numb, last_numb);
        }
    }
}
