using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using MetroFramework.Components;
using MetroFramework.Forms;
using V83;

namespace Diplom1
{
    public partial class Form1 : MetroForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        string connectionString = @"Data Source = D09D\SQL_EXPRESS; Initial Catalog = ""Компетенции и Ко""; Integrated Security = True";

        private Stopwatch TimeStart()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            return stopWatch;
        }
        private void TimeStop(Stopwatch stwp, string Inf)
        {
            stwp.Stop();
            TimeSpan ts = stwp.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            ProgressInfo.AppendText(Inf + elapsedTime + Environment.NewLine);
        }

        private bool TextBoxEmpty()
        {
            if (FIOBox.TextLength == 0 || CodeBBox.TextLength == 0 || SemNumBox.TextLength == 0 ||
                CodePBox.TextLength == 0 || YearBox.TextLength == 0)
            {
                MessageBox.Show("Нужно заполнить все поля!", "Неправильное заполнение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            else
                return true;
        }

        private bool TextBoxFormat()
        {
            Regex r;
            r = new Regex(@"^[\d]+$");
            if (!r.IsMatch(YearBox.Text.Trim()))
            {
                MessageBox.Show("Неверный формат поля 'Год начала обучения'", "Неправильное заполнение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            if (!(Int32.Parse(YearBox.Text.Trim()) > 1990) || !(Int32.Parse(YearBox.Text.Trim()) < 2020))
            {
                MessageBox.Show("Неверный формат поля 'Год начала обучения'", "Неправильное заполнение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            if (!r.IsMatch(SemNumBox.Text.Trim()))
            {
                MessageBox.Show("Неверный формат поля 'Номер семестра'", "Неправильное заполнение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            r = new Regex(@"^[a-zA-Zа-яА-Я\s]+$");
            if (!r.IsMatch(FIOBox.Text.Trim()))
            {
                MessageBox.Show("Неверный формат поля 'ФИО студента'", "Неправильное заполнение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            r = new Regex(@"\d+\.\d+\.\d+");
            if (!r.IsMatch(CodeBBox.Text.Trim()))
            {
                MessageBox.Show("Неверный формат поля 'Базовая специальность'", "Неправильное заполнение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            if (!r.IsMatch(CodePBox.Text.Trim()))
            {
                MessageBox.Show("Неверный формат поля 'Специальность перевода'", "Неправильное заполнение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        private bool WebStart(string ProfileB, string ProfileP, string Year)
        {
            Stopwatch stwc = new Stopwatch();
            stwc.Start();

            switch (Year)
            {
                case "2015":
                    Year = "Учебный год 2015/2016";
                    break;
                case "2016":
                    Year = "Учебный год 2016/2017";
                    break;
                case "2017":
                    Year = "Учебный год 2017/2018";
                    break;
                case "2018":
                    Year = "Учебный год 2018/2019";
                    break;
                case "2019":
                    Year = "Учебный год 2019/2020";
                    break;
            }

            HtmlWeb web = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument DocWeb = web.Load(WebSiteBox.Text);

            TimeStop(stwc, "Подключение к веб-ресурсу было выполнено за ");
            stwc = new Stopwatch();
            stwc.Start();

            string BAZ_xPath = "//tr[td='" + ProfileB + "' and td='" + Year + "']";
            string PEREVOD_xPath = "//tr[td='" + ProfileP + "' and td='" + Year + "']";

            string Baz = "Базовая специальность";
            string Perevod = "Специальность перевода";

            if (!WebSearch(BAZ_xPath, WebBox1,DocWeb,Baz) || !WebSearch(PEREVOD_xPath, WebBox2, DocWeb, Perevod))
            {
                return false;
            }

            bool WebSearch(string xPath, System.Windows.Forms.TextBox textbox, HtmlAgilityPack.HtmlDocument docum, string Spec)
            {
                var nodes = docum.DocumentNode.SelectNodes(xPath);
                if (nodes == null)
                {
                    MessageBox.Show(Spec + " не найдена", "Поиск в Web");
                    return false;
                }
                var HTMLList = from table in nodes.Cast<HtmlNode>()
                                from row in table.SelectNodes("td").Cast<HtmlNode>()
                                select new { Cell_Text = row.InnerText };

                //Выводим текст
                foreach (var cell in HTMLList)
                {
                    textbox.AppendText(cell.Cell_Text + Environment.NewLine);
                }
                return true;
            }

            TimeStop(stwc, "Поиск данных в веб-ресурсе был выполнен за ");
            return true;
        }
            
        private bool OneCStart(string profB, string profP,string fio, ref string formBAZ,
            ref string formPER, ref string[] DiscListB, ref string[] DiscListP)
        {

            dynamic connect1C = Get1CConnection();

            try
            {
                string baz = "базовая";
                string per = "переводная";

                if (!OneCSearch(connect1C, profB, baz, fio, ref formBAZ, ref DiscListB) || !OneCSearch(connect1C, profP, per, fio, ref formPER, ref DiscListP))
                    return false;
                else
                    return true;
            }

            finally
            {
                Marshal.ReleaseComObject(connect1C);
            }

            bool OneCSearch(dynamic conn, string prof, string vidspec, string fioStud, ref string formObuch, ref string[] DiscList)
            {
                dynamic spec = null;
                spec = conn.Документы.УчебныйПлан.НайтиПоРеквизиту("Направление", prof);
                if (spec == null)
                {
                    MessageBox.Show(vidspec + "специальность не найдена", "Поиск в 1C");
                    return false;
                }

                formObuch = spec.Форма;
                dynamic tabCh = spec.Дисциплины;

                int size = spec.Дисциплины.Количество();

                DiscList = new string[size];
                string[] DiscKod = new string[size];
                string[] DiscSem = new string[size];
                string[] trudZE = new string[size];
                string[] trudCH = new string[size];

                int i = 0;
                foreach (dynamic discip in tabCh)
                {
                    DiscList[i] = discip.Дисциплина;
                    DiscKod[i] = discip.Код;
                    DiscSem[i] = discip.Семестр;
                    trudZE[i] = discip.ТрудоемкостьЗЕ;
                    trudCH[i] = discip.ТрудоемкостьЧАС;
                    i++;
                }

                SQLfillDiscip(connectionString, DiscList, DiscKod, DiscSem, vidspec, fioStud, trudZE, trudCH);
                return true;
            }

            dynamic Get1CConnection()
            {
                COMConnector comConnector = new COMConnector();
                dynamic connect = comConnector.Connect("File=\"C:\\InfoBase\"");
                return connect;
            }
        }

        private bool ExcelStart(string[] ProfileB, string[] ProfileP, string fiio)
        
        {
            //Excel-parser
            Excel.Application ObjExcel = null;
            Workbooks wrkbks = null;
            Workbook ObjWorkBook = null;
            Worksheet sheet = null;

            try
            {
                string pathToFile = ExcelBox.Text;
                ExcelConnect(pathToFile, ref ObjExcel, ref wrkbks, ref ObjWorkBook);

                //Ищем компетенции для базового и переводного профиля
                string baz = "базовая";
                string per = "переводная";
                if (ExcelSearch(ProfileB, baz, ObjWorkBook, ref sheet, 1, 2) == false)
                {
                    MessageBox.Show("Дисциплины базовой специальности не найдены", "Поиск в Excel");
                    return false;
                }
                if (ExcelSearch(ProfileP, per, ObjWorkBook, ref sheet, 1, 2) == false)
                {
                    MessageBox.Show("Дисциплины специальности перевода не найдены", "Поиск в Excel");
                    return false;
                }
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(ObjWorkBook);
                Marshal.ReleaseComObject(wrkbks);
                Marshal.ReleaseComObject(ObjExcel);
            }

            bool ExcelSearch(string[] Disc,string vidSpec, Workbook wrk, ref Worksheet sht, int numb1, int numb2)
            {
                string CompTemp = null;
                List<string> NameList = new List<string>();

                sht = (Worksheet)wrk.Sheets[numb1];

                Stopwatch stwc = new Stopwatch();
                stwc.Start();

                for (int i = 0; i < Disc.Length; i++)
                {
                    for (int j = 1; j <= 150; j++)
                    {
                        if (sht.Cells[j, 3].Value as string == Disc[i])
                        {
                            CompTemp = sht.Cells[j, 83].Value as string;
                            if (CompTemp.Contains(";"))
                            {
                                string[] tempList = CompTemp.Split(';');
                                for (int k = 0; k < tempList.Length; k++)
                                    NameList.Add(tempList[k]);
                            }
                            else
                            {
                                NameList.Add(CompTemp);
                            }
                            break;
                        }
                    }
                    if (CompTemp == null)
                        return false;                    
                }

                sht = (Worksheet)ObjWorkBook.Sheets[numb2];

                string[] NaimenComp = NameList.ToArray<string>();
                string[] CompText = new string[NaimenComp.Length];

                int flag = 0;

                for (int i = 0; i < NaimenComp.Length; i++)
                {
                    for (int j = 1; j <= 250; j++)
                    {
                        if (sht.Cells[j, 2].Value as string == NaimenComp[i].Trim())
                        {
                            CompText[i] = sht.Cells[j, 4].Value as string;
                            flag++;
                            break;
                        }
                    }
                    if (flag == NaimenComp.Length)
                        break;
                }

                SQLfillComp(connectionString, NaimenComp, CompText, vidSpec, fiio);

                TimeStop(stwc, "Поиск компетенций в Excel был выполнен за ");
                return true;
            }

            void ExcelConnect(string path, ref Excel.Application exc, ref Workbooks workbooks, ref Workbook workbook)
            {
                Stopwatch stwc = new Stopwatch();
                stwc.Start();

                //Подключение к Эксель
                exc = new Excel.Application();
                workbooks = exc.Workbooks;
                workbook = workbooks.Open(path, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                TimeStop(stwc, "Подключение к Excel было выполнено за ");
            }
            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.competencTableAdapter.Fill(this.компетенции_и_КоDataSet.Competenc);
            this.discipTableAdapter.Fill(this.компетенции_и_КоDataSet.Discip);
            this.infoTableAdapter.Fill(this.компетенции_и_КоDataSet.Info);
            this.WindowState = FormWindowState.Maximized;

        }
        private bool SQLfillInfo(string connectSQL, string ProfileB, string ProfileP,
            string FIO, string Year, string SemNum, string ProgFormB, string ProgFormP)
        {
            SqlConnection cn;
            string strSQL;
            using (cn = new SqlConnection(connectSQL))
            {
                try
                {
                    cn.Open();
                    strSQL = "INSERT INTO Info ([Базовая специальность], [Специальность перевода], [ФИО студента]," +
                        " [Год начала обучения], [Номер семестра], [Форма обучения(баз)], [Форма обучения(перевод)]) " +
                        "Values ('" + ProfileB + "', '" + ProfileP + "', '" + FIO + "', '" + Year + "', '" + SemNum + "', '" + ProgFormB + "', '" + ProgFormP + "')";
                    SqlCommand cmd = new SqlCommand(strSQL, cn);
                    if (cmd.ExecuteNonQuery() == 1)
                        ProgressInfo.AppendText("Запись в основную таблцу успешно добавлена" + Environment.NewLine);
                    else
                        ProgressInfo.AppendText("Запись в основную таблицу не была добавлена" + Environment.NewLine);
                    cn.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    return false;
                }
            }
            this.infoTableAdapter.Fill(this.компетенции_и_КоDataSet.Info);
            return true;
        }

        private void SQLfillDiscip(string connectSQL, string[] DiscNaimen, string[] DiscCode, string[] DiscSemestr, string VidSpec, string fio, string[] TrudZE, string[] TrudChas)
        {
            SqlConnection cn;
            string strSQL;
            SqlCommand cmd = null;
            using (cn = new SqlConnection(connectSQL))
            {
                try
                {
                    cn.Open();
                    for (int i = 0; i < DiscNaimen.Length; i++)
                    {
                        strSQL = "INSERT INTO Discip (Код, Наименование, [Вид специальности], [ФИО студента], Семестр, [Трудоемкость(ЗЕ)],[Трудоемкость(час)])" +
                            "Values ('" + DiscCode[i].Trim() + "', '" + DiscNaimen[i].Trim() + "', '" + VidSpec + "', '" + fio + "', '" + DiscSemestr[i] + "', '" + TrudZE[i] + "', '"+ TrudChas[i] +"')";
                        cmd = new SqlCommand(strSQL, cn);
                        cmd.ExecuteNonQuery();
                    }
                    ProgressInfo.AppendText("Запись о дисциплинах успешно добавлена(" + VidSpec + " специальность)" + Environment.NewLine);
                    cn.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    ProgressInfo.AppendText("Запись о дисциплинах не была добавлена(" + VidSpec + " специальность)" + Environment.NewLine);
                }
            }
            this.discipTableAdapter.Fill(this.компетенции_и_КоDataSet.Discip);
        }

        private void SQLfillComp(string connectSQL, string[] Naimenovanie, string[] Soderzhanie, string vidSpec, string fio)
        {
            SqlConnection cn;
            string strSQL;
            SqlCommand cmd = null;
            using (cn = new SqlConnection(connectSQL))
            {
                try
                {
                    cn.Open();
                    for (int i = 0; i < Naimenovanie.Length; i++)
                    {
                        strSQL = "INSERT INTO Competenc (Наименование, [Вид специальности], [ФИО студента], Содержание)" +
                            "Values ('" + Naimenovanie[i].Trim() + "', '" + vidSpec + "', '" + fio + "', '" +  Soderzhanie[i].Trim() + "')";
                        cmd = new SqlCommand(strSQL, cn);
                        cmd.ExecuteNonQuery();
                    }
                    if (cmd.ExecuteNonQuery() == 1)
                        ProgressInfo.AppendText("Запись о компетенциях успешно добавлена(" + vidSpec + " специальность)" + Environment.NewLine);
                    else
                        ProgressInfo.AppendText("Запись о компетенциях не была добавлена(" + vidSpec + " специальность)" + Environment.NewLine);

                    cn.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            this.competencTableAdapter.Fill(this.компетенции_и_КоDataSet.Competenc);
        }

        private void FindButton_Click(object sender, EventArgs e)
        {
            ProgressInfo.Clear();
            TickPic.Visible = false;
            CrossPic.Visible = false;
            if (TextBoxEmpty() && TextBoxFormat())
            {
                ProgressInfo.AppendText("Поиск начат..." + Environment.NewLine);
                string ProfileNameB, ProfileNameP;
                string Code, CodePerevod;
                string Year;
                string FIO;
                string SemNum;
                string ProgFormB = null, ProgFormP = null;
                string[] DiscListB = null, DiscListP = null;

                FIO = FIOBox.Text;
                Code = CodeBBox.Text;
                CodePerevod = CodePBox.Text;
                SemNum = SemNumBox.Text;
                Year = YearBox.Text.Trim();

                if (WebStart(Code, CodePerevod, Year))
                {
                    ProfileNameB = WebBox1.Lines[1];
                    ProfileNameP = WebBox2.Lines[1];
                    if (OneCStart(ProfileNameB, ProfileNameP, FIO, ref ProgFormB, ref ProgFormP,ref DiscListB, ref DiscListP))
                    {
                        if (ExcelStart(DiscListB, DiscListP, FIO))
                        {
                            if (SQLfillInfo(connectionString, ProfileNameB, ProfileNameP, FIO, Year, SemNum, ProgFormB, ProgFormP))
                                TickPic.Visible = true;
                            else
                                CrossPic.Visible = true;
                        }
                        else
                            CrossPic.Visible = true;
                    }
                    else
                        CrossPic.Visible = true;
                }
                else
                    CrossPic.Visible = true;
            }
            else
                CrossPic.Visible = true;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();
            newForm.ShowDialog();
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (ResourseCheck.Checked == true)
                SourseGroup.Visible = true;
            else
                SourseGroup.Visible = false;
        }

        private void infoBindingNavigatorSaveItem_Click_3(object sender, EventArgs e)
        {
            this.Validate();
            this.infoBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.компетенции_и_КоDataSet);
        }

        private void toolStripButton13_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.infoBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.компетенции_и_КоDataSet);
        }

        private void toolStripButton20_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.infoBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.компетенции_и_КоDataSet);
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.discipBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.компетенции_и_КоDataSet);

        }
    }
}
