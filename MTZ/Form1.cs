//programmed by Flexi => BETA
using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Data;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Text;
using System.Globalization;
using System.IO;
using MTZ;
using System.Runtime.InteropServices;


namespace MTZ
{
    public partial class Form1 : Form
    {
        string excelFilePath = "";
        static string fileAlert = "Es muss eine Datei angegeben werden!";
        //Einstempel Spanne
        private TimeSpan start = new TimeSpan(5, 30, 0);
        private TimeSpan end = new TimeSpan(10, 30, 0);

        //Ausstempel Spanne
        private TimeSpan start1 = new TimeSpan(10, 31, 0);
        private TimeSpan end1 = new TimeSpan(19, 0, 0);
        public Form1()
        {
            InitializeComponent();
            label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            InitializeComboBox();
            panel1.AllowDrop = true;
            panel1.DragDrop += Form1_DragDrop;
            panel1.DragEnter += Form1_DragEnter;
            panel1.BorderStyle = BorderStyle.FixedSingle;
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
        }

        //Hover Funktion für Panel
        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        //Drag and drop Funktion für Panel
        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                label4.Text = (file);
                excelFilePath = file;
            }
        }

        //div. Methoden welche bei Klick von Button aufgerufen werden
        private void button1_Click(object sender, EventArgs e)
        {
            if (excelFilePath == string.Empty)
            {
                MessageBox.Show(fileAlert);
            }
            else
            {
                string monat = "";
                if(comboBox1.SelectedItem != null)
                monat = comboBox1.SelectedItem.ToString();
                try
                {
                    int errorCatcher = DateTime.ParseExact(monat, "MMMM", CultureInfo.CurrentCulture).Month;

                }
                catch{
                    MessageBox.Show("Geben Sie einen gültigen Monat an.");
                    return;
                }
                List<MSG> msg = new List<MSG>();
                Microsoft.Office.Interop.Outlook.Application application = new Microsoft.Office.Interop.Outlook.Application();
                Accounts acc = application.Session.Accounts;
                Folder root = application.Session.DefaultStore.GetRootFolder() as Folder;
                ExchangeUser exchangeUser = application.Session.CurrentUser.AddressEntry.GetExchangeUser();
                foreach (Folder item in application.Session.Folders)
                {

                    foreach (Folder folder in item.Folders)
                    {
                        if (folder.Name == textBox2.Text && item.Name == exchangeUser.PrimarySmtpAddress)
                        {
                            root = folder;
                            break;
                        }
                    }
                }
                Folder f = EnumerateFolders(root, textBox2.Text);

                string name = exchangeUser.LastName + ", " + exchangeUser.FirstName;

                string year = textBox1.Text;
                if (year == string.Empty) year = DateTime.Now.Year.ToString();
                msg = IterateMessages(f, monat, year, name);

                DoppelListe dl = new DoppelListe();
                ExcelData ed = new ExcelData();

                if (msg != null)
                {
                    dl = FilterData(msg, exchangeUser.LastName, exchangeUser.FirstName, monat, year);
                    listBox1.DataSource = dl.zeitenListe;
                    ed = openExcel();
                    if (ed.goOn)
                    {
                        addDataToExcel(dl.zeitenListe, monat, dl.datumsListe, ed.ws);
                    }
                    //speichert und schließt excel workbook
                    ed.wb.Save();
                    ed.wb.Close();
                    ed.app.Quit();
                    //beendet excel prozess
                    Marshal.ReleaseComObject(ed.wb);
                    Marshal.ReleaseComObject(ed.app);
                    MessageBox.Show("Die Datei wurde erfolgreich bearbeitet.");
                }
            }
        }

        public ExcelData openExcel()
        {
            string spreadsheetLocation = Path.Combine(Directory.GetCurrentDirectory(), excelFilePath);
            try
            {
                var exApp = new Microsoft.Office.Interop.Excel.Application();
                var exWbk = exApp.Workbooks.Open(spreadsheetLocation);
                var exWks = (Microsoft.Office.Interop.Excel.Worksheet)exWbk.Sheets[2];

                return new ExcelData(true, exApp, exWbk, exWks);
            }
            catch(System.Exception e) 
            {
                MessageBox.Show("Die angegebene Datei konnte nicht geöffnet werden. Überprüfen Sie bitte ob schreibrechte vorliegen und ob der Pfad noch aktuell ist: " + spreadsheetLocation
                     + "\r\n" + e);
                return new ExcelData(false,null,null,null);
            }
        }

        //leert die angegebene Excel Datei
        public Worksheet clearExcel(Worksheet ws)
        {
            for (int row = 8; row < 50; row++)
            {
                ws.Cells[row, "E"] = "";
                ws.Cells[row, "F"] = "";
            }
            return ws;
        }

        //füllt Excel Datei mit Daten
        public void addDataToExcel(List<string> timeList, string SelectedMonth, List<string> dates, Worksheet myExcelWorkSheet)
        {
            myExcelWorkSheet = clearExcel(myExcelWorkSheet);

            int monat = DateTime.ParseExact(SelectedMonth, "MMMM", CultureInfo.CurrentCulture).Month;
            int DaysInSelectedMonth = DateTime.DaysInMonth(2018, monat);
            dates.Reverse();
            int i = 8 + DaysInSelectedMonth;

            //Geht ab der 8ten Excel Reihe durch Dokument
            for (int rowNumber = 8; rowNumber < i; rowNumber++)
            {
                int color = (int)(myExcelWorkSheet.Cells[rowNumber, "A"] as Range).Font.Color;
                DateTime cellValue = (DateTime)(myExcelWorkSheet.Cells[rowNumber, "A"] as Range).Value;

                foreach (string item in dates)
                {
                    DateTime dt = DateTime.Parse(item);

                    //Wenn Farbe des Datums Rot oder Blau ist überspringen
                    if (color == 255 || color == 16711680)
                    {
                        break;
                    }
                    //Überprüfe auf Datum und gucke ob dieses in der Liste vorhanden ist, wenn ja fülle Reihe mit zeiten aus
                    else if (cellValue.ToString("dd.MM.yy") == dt.ToString("dd.MM.yy"))
                    {
                        int index = dates.FindIndex(a => a == item);
                        index = index * 2;
                        //ABFRAGE OB ZEIT INNERHALB DER EINSTEMPELZEITEN LIEGT, WENN NICHT N.A. EINFÜGEN
                        if ((DateTime.Parse(timeList.ElementAt(index)).TimeOfDay < this.end) && (DateTime.Parse(timeList.ElementAt(index)).TimeOfDay > this.start))
                        myExcelWorkSheet.Cells[rowNumber, "E"] = timeList.ElementAt(index);
                        else
                        myExcelWorkSheet.Cells[rowNumber, "E"] = "N.V.";
                        //ABFRAGE OB ZEIT INNERHALB DER AUSSTEMPELZEITEN LIEGT, WENN NICHT N.A. EINFÜGEN
                        if ((DateTime.Parse(timeList.ElementAt(index+1)).TimeOfDay < this.end1) && (DateTime.Parse(timeList.ElementAt(index+1)).TimeOfDay > this.start1))
                        myExcelWorkSheet.Cells[rowNumber, "F"] = timeList.ElementAt(index + 1);
                        else
                        myExcelWorkSheet.Cells[rowNumber, "F"] = "N.V.";
                        break;
                    }
                    //Wenn Datum nicht in der Liste vorhanden ist fülle auf mit BS
                    else
                    {
                        myExcelWorkSheet.Cells[rowNumber, "E"] = "BS";
                    }
                }
            }

        }

        //geht durch Ordner von Outlook und gibt den Posteingang zurück, wenn kein spezieller Ordner angegeben wurde
        static Folder EnumerateFolders(Folder folder, string foldername)
        {
            if (foldername == string.Empty) foldername = "Posteingang";

            Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {

                foreach (Folder childFolder in childFolders)
                {
                    if (childFolder.FolderPath.Contains(foldername))
                    {
                        return childFolder;
                    }
                }
            }
            else
            {
                return folder;
            }
            
            return null;
        }

        //geht durch alle Nachrichten im Posteingang und sucht sich Mails mit Stempelzeiten und Monat in Subject
        static List<MSG> IterateMessages(Folder folder, string monat, string year, string name)
        {
            List<MSG> mails = new List<MSG>();
            string yearWithLastTwoDigits = year.Remove(0, 2);
            var fi = folder.Items;
            bool notFound = false;
            if (fi != null)
            {
                foreach (var item in fi)
                {
                    try
                    {
                        MailItem mi = (MailItem)item;

                        //wenn der Titel der Mail Stempelzeiten, den Monat und das Jahr beinhaltet 
                        if (mi.Subject.Contains("Stempelzeiten") && mi.Subject.Contains(monat) && (mi.Subject.Contains(year) || mi.Subject.Contains(yearWithLastTwoDigits)))
                        {
                            Console.WriteLine("TEST2" + mi.SentOn.Year.ToString());
                            MSG msg = new MSG(mi.Body, mi.Subject, mi.SenderName);
                            mails.Add(msg);
                            notFound = false;
                        }
                        if (mi.Subject.Contains("Stempelzeiten") && mi.Subject.Contains(monat))
                        {
                            Console.WriteLine("TEST1" + mi.SentOn.Year.ToString());
                            MSG msg2 = new MSG(filterMultipleMonths(mi.Subject, mi.Body, monat, yearWithLastTwoDigits), mi.Subject, mi.SenderName);
                            mails.Add(msg2);
                            notFound = false;
                        }
                    }
                    catch
                    {
                        notFound = true;
                    }
                }
                if(notFound)
                    MessageBox.Show("Mit den von Ihnen angegebenen Daten konnten keine Stempelzeiten gefunden werden. Bitte überprüfen Sie Ihre Eingaben.");
                else                
                    return mails;
            }
            return null;
        }

        static string filterMultipleMonths(string mailSubject, string mailBody,string monat, string year)
        {
            int selectedMonth = DateTime.ParseExact(monat, "MMMM", CultureInfo.CurrentCulture).Month;
            int oMonth = selectedMonth + 1;
            string previousMonth = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(selectedMonth - 1);
            string occuringMonth = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(selectedMonth + 1);
            if (mailSubject.Contains(previousMonth))
            {
            int index = 0;
                for (int i = 1; i < 32; i++)
                {
                    try
                    {
                        string a = i + "." + selectedMonth + "." + year;
                        index = mailBody.IndexOf(a);
                        break;
                    }
                    catch { }
                }
                string caseB = mailBody.Substring(index);
                return caseB;
            }
            else if (mailSubject.Contains(occuringMonth))
            {
                int index = 0;
                for (int i = 1; i < 32; i++)
                {
                    try
                    {
                        string a = i + "." + oMonth + "." + year;
                        index = mailBody.IndexOf(a);
                        break;
                    }
                    catch { }
                }
                if (index > 0)
                {
                    string caseA = mailBody.Substring(0, index);
                    return caseA;
                }

            }            
            return "";
        }

        //initialisert Dropdown mit Monaten
        private void InitializeComboBox()
        {
            comboBox1.TabIndex = 0;

            string[] installs = new string[] {"Januar", "Februar",
            "März", "April", "Mai",
            "Juni", "Juli",
            "August", "September", "Oktober",
            "November", "Dezember"};
            comboBox1.AutoCompleteCustomSource.AddRange(installs);
            comboBox1.Items.AddRange(installs);
            Controls.Add(comboBox1);
        }



        //filtert Daten je nach Absender, unterschiediche Schriftweisen, Screenshots etc. 
        private DoppelListe FilterData(List<MSG> msg, string lname, string fname, string monat, string year)
        {
            List<string> zeitenListe = new List<string>();
            List<string> datumsListe = new List<string>();
            List<Stempelzeit> stempelzeiten = new List<Stempelzeit>();
            foreach (MSG item in msg)
            {
                string mailBody = item.messages;
                string tableWithoutNames = "";
                bool twoletters = false;
                int i = 0;
                int selectedMonth = DateTime.ParseExact(monat, "MMMM", CultureInfo.CurrentCulture).Month;
                int selectedYear = Int16.Parse(year.Remove(0, 2));
                //geht solange durch die schleife bis ein datum im mail body gefunden wurde, gibt index von erst gefundenem Datum zurück
                for (int o = 1; o < 32; o++)
                {
                    string date = "";
                    date = o + "." + selectedMonth + ".18";

                    bool a = Regex.IsMatch(mailBody, @"(^|\s)" + date + @"(\s|$)");
                    if (a)
                    {
                        if (o > 9) twoletters = true;
                        else twoletters = false;
                        i = mailBody.IndexOf(date);
                        break;
                    }
                }
                //wenn erst gefundenes datum eine zweistelligen Tag beinhaltet verringer index um 1
                if (twoletters)
                    i--;

                if (i >= 0) mailBody = mailBody.Substring(i);
                string name = lname + ", " + fname;
                if (name.Length > 15)
                {
                    name = name.Substring(0, 15);
                }

                int j = mailBody.LastIndexOf(name);
                if (j > 0)
                {
                    tableWithoutNames = mailBody.Substring(0, j);
                }
                tableWithoutNames = tableWithoutNames.Replace(name, "");

                //loopt durch jedes mögliche Datum
                for (int o = 32; o > 0; o--)
                {
                    string datum = "" + o + "." + selectedMonth + "." + selectedYear + "";
                    if (tableWithoutNames.Contains(datum))
                    {
                            datumsListe.Add(datum);
                        tableWithoutNames = Regex.Replace(tableWithoutNames, datum, "");
                    }
                }
                zeitenListe.Add(tableWithoutNames);
            }

            DoppelListe dl = new DoppelListe(zeitenListe, datumsListe);
            dl.zeitenListe = separateTimes(dl.zeitenListe);
            dl.zeitenListe = removeDoubles(dl.zeitenListe);
            return dl;
        }

        //entfernt jeglichen whitespace aus Zeiten
        private List<string> separateTimes(List<string> Times)
        {
            string[] buffer = new string[50];
            List<string> listWithoutWhitespace = new List<string>();

            foreach (string item in Times)
            {
                string ss = item.Replace("\r\n", "~");
                ss = Encoding.ASCII.GetString(Encoding.ASCII.GetBytes(ss));
                ss = ss.Replace("?", "~");
                try
                {
                    buffer = ss.Split(new string[] { "~~~~~~" }, StringSplitOptions.None);
                }
                catch
                {

                }
                for (int t = 0; t < buffer.Length; t++)
                {
                    buffer[t] = buffer[t].Replace("~", "");

                }
                listWithoutWhitespace.AddRange(buffer);
            }
            return listWithoutWhitespace;
        }

        //entfernt Zeiten welche durch doppeltes Ein- oder Ausstempeln entstehen
        private List<string> removeDoubles(List<string> list)
        {
            List<string> rawTimeData = new List<string>();

            DateTime lastDateTime = new DateTime();
            string lastItem = "";
            foreach (string item in list)
            {
                try{
                    DateTime dt = DateTime.Parse(item);
                //Repräsentiert Zeiten zum Einstempeln, mit etwas Toleranz
                if ((dt.TimeOfDay < this.end) && (dt.TimeOfDay > this.start) && (lastDateTime.TimeOfDay < this.end) && (lastDateTime.TimeOfDay > this.start))
                {
                    continue;
                }
                    //Repräsentiert Zeiten zum Ausstempeln mit etwas Toleranz
                else if ((dt.TimeOfDay < this.end1) && (dt.TimeOfDay > this.start1) && (lastDateTime.TimeOfDay < this.end1) && (lastDateTime.TimeOfDay > this.start1))
                {
                    rawTimeData.RemoveAt(rawTimeData.Count-1);
                    rawTimeData.Add(item);
                    continue;
                }

                if (DateTime.Compare(dt, lastDateTime) == 0)
                {
                    continue;
                }
                else
                {
                    rawTimeData.Add(item);
                }
                lastDateTime = dt;
                lastItem = item;
            }
                catch
                {

                }
            }
            return rawTimeData;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (excelFilePath != string.Empty)
                excelFilePath = "";
            label4.Text = "";
        }
    }
}
