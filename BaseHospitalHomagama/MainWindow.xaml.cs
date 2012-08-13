using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.Data.OleDb;
using System.Diagnostics;
using System.Windows.Threading;
using System.Threading;
using System.Data;

namespace BaseHospitalHomagama
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Label label;
        AutoCompleteBox text;
        List<Label> speciLables = new List<Label>();
        List<AutoCompleteBox> speciTexts = new List<AutoCompleteBox>();
        specimenSuggestions spSg;
        List<String> suggestions;
        Database database = new Database("database.mdb");

        private DispatcherTimer timer1 =new DispatcherTimer();


        public MainWindow()
        {
            InitializeComponent();
            textClinicalDetails.SpellCheck.IsEnabled = true;
            textMacroscopy.SpellCheck.IsEnabled = true;
            textMicroscopy.SpellCheck.IsEnabled = true;
            textConclusion.SpellCheck.IsEnabled = true;
            speciLables.Add(label7);
            speciTexts.Add(textSpecimen);
            comboBoxTitle.SelectedIndex = 0;
            comboBoxGender.SelectedIndex = 0;

            datePicker1.SelectedDate = DateTime.Today;
            datePicker2.SelectedDate = DateTime.Today;
            
            spSg = new specimenSuggestions();

            timer1.Interval = new TimeSpan(0, 0, 2);
            timer1.Tick += new EventHandler(timer1_Elapsed);

            if (!File.Exists("specimenList.cse"))
            {
                spSg.specimen = new String[] {"Uterine Curettings","Cervical Polyps","Product of Conception",
            "Product of ERPC","Endometrial Sampling","Uterus and Bilateral Ovaries","Ovarian Cyst",
            "Thyroid Gland","Appendix","Breast Lump","Sebaceous Cyst","Ganglion"};
                spSg.store();
            }
            else
            {
                spSg = specimenSuggestions.retrieve();
            }
            textSpecimen.ItemsSource = spSg.specimen;

            //test 
            //DataContext = new List<Record>
            //{
            //    new Record ("FDFDH","D","EFW","FE","CXV",34,"EFW", new string[]{"df","ew"},"cxv","setds","cbv","32"),
            //     new Record ("FDFDH","D","EFW","FE","hgfn",34,"EFhgfhgfhfghf", new string[]{"df","ew"},"cxv","sfdfgdfgdfg","cbv","32")
            //};
            

        }

        void timer1_Elapsed(object sender, EventArgs e)
        {
            labelError.Visibility = System.Windows.Visibility.Hidden;
            labelError.Foreground = Brushes.Red;
            timer1.Stop();
        }

        private void buttonAddSpecimen_Click(object sender, RoutedEventArgs e)
        {
            if (speciTexts.Count == 1)
            {
                label7.Content = "Specimen 1 :";
                buttonRemoveSpecimen.Visibility = System.Windows.Visibility.Visible;
            }
            speciLables.Add(label = new Label());
            label.Width = label7.Width;
            label.Height = label7.Height;
            speciTexts.Add(text = new AutoCompleteBox());
            text.ItemsSource = spSg.specimen;
            label.Content = "Specimen " + speciLables.Count + " :";
            label.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Right;
            text.Height = 26;
            text.Width = textSpecimen.Width;
            stackPanel2.Children.Add(label);
            stackPanel1.Children.Add(text);
            label.Margin = new Thickness(0,4,0,0);
            text.Margin = new Thickness(0, 4, 0, 0);
            grid2.Margin = new Thickness(grid2.Margin.Left, grid2.Margin.Top + 30, grid2.Margin.Right, grid2.Margin.Bottom);
            buttonAddSpecimen.Margin = new Thickness(buttonAddSpecimen.Margin.Left, buttonAddSpecimen.Margin.Top + 30, buttonAddSpecimen.Margin.Right, buttonAddSpecimen.Margin.Bottom);
            buttonRemoveSpecimen.Margin = new Thickness(buttonRemoveSpecimen.Margin.Left, buttonRemoveSpecimen.Margin.Top + 30, buttonRemoveSpecimen.Margin.Right, buttonRemoveSpecimen.Margin.Bottom);
        }

        private Boolean retrieveReport(String reportNo)
        {
            clear();
            database.connectToDatabase();
            if (database.hasEntry(reportNo))
            {
                Record tempRecord = database.getReport(reportNo);
                textReportNo.Text = reportNo;
                textPatientName.Text = tempRecord.name;
                textWardNo.Text = tempRecord.ward;
                textBhtNo.Text = tempRecord.bht;
                comboBoxTitle.SelectedIndex = -1;
                comboBoxTitle.Text = tempRecord.title;
                textPatientName.Text = tempRecord.name;
                textAge.Text = tempRecord.age.ToString();
                comboBoxGender.Text = tempRecord.gender;
                textSpecimen.Text = tempRecord.specimen[0];
                for (int i = 1; i < tempRecord.specimen.Length; i++)
                {
                    buttonAddSpecimen_Click(this, (RoutedEventArgs)RoutedEventArgs.Empty);
                    speciTexts.Last().Text = tempRecord.specimen[i];
                }

                textMacroscopy.Text = tempRecord.macroscopy;
                textMicroscopy.Text = tempRecord.microscopy;
                textConclusion.Text = tempRecord.conclusion;
                database.closeConnection();
                return true;

            }
            else
            {
                database.closeConnection();
                return false;
            }


        }

        private void print()
        {
            object FileName = AppDomain.CurrentDomain.BaseDirectory + "\\BASE HOSPITAL HOMAGAMA.docx";//
            object saveAs = "b.docx";
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Options.set_DefaultFilePath(0, AppDomain.CurrentDomain.BaseDirectory);
            Microsoft.Office.Interop.Word.Document aDoc = null;
            object readOnly = true;
            object isVisible = false;
            wordApp.Visible = false;
            aDoc = wordApp.Documents.Open(ref FileName, ref missing, ref readOnly, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible,
                 ref missing, ref missing, ref missing, ref missing);
            aDoc.Activate();

            FindindReplace(wordApp, "<name>", comboBoxTitle.Text + " " + textPatientName.Text);
            FindindReplace(wordApp, "<age>", textAge.Text);
            FindindReplace(wordApp, "<gender>", comboBoxGender.Text);
            FindindReplace(wordApp, "<rep>", textReportNo.Text);
            FindindReplace(wordApp, "<ward>", textWardNo.Text);
            FindindReplace(wordApp, "<bht>", textBhtNo.Text);
            for (int i = 0; i < speciLables.Count; i++)
            {
                if (i < speciLables.Count - 1)
                    FindindReplace(wordApp, "<specimen>", "Specimen " + (i + 1) + "\t: " + speciTexts.ElementAt(i).Text + "\n<specimen>");
                else
                {
                    if (speciLables.Count == 1)
                        FindindReplace(wordApp, "<specimen>", "Specimen\t: " + speciTexts.ElementAt(i).Text);
                    else
                        FindindReplace(wordApp, "<specimen>", "Specimen " + (i + 1) + "\t: " + speciTexts.ElementAt(i).Text);
                }
            }
            FindindReplace(wordApp, "<macro>", replaceNewLines(textMacroscopy.Text));
            FindindReplace(wordApp, "<micro>", replaceNewLines(textMicroscopy.Text));
            FindindReplace(wordApp, "<con>", replaceNewLines(textConclusion.Text));
            FindindReplace(wordApp, "<date>", datePicker1.DisplayDate.Date.Day + " / " + datePicker1.DisplayDate.Date.Month + " / " + datePicker1.DisplayDate.Date.Date.Year);


            aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            wordApp.Quit();

            ProcessStartInfo info = new ProcessStartInfo(saveAs.ToString());

            info.Verb = "Print";

            info.CreateNoWindow = true;

            info.WindowStyle = ProcessWindowStyle.Hidden;

            Process.Start(info);

        }
        private String replaceNewLines(String str)
        {
            String ret="";
            for (int i = 0; i < str.Length; i++)
            {
                if ((char)str.ElementAt(i) == '\n')
                {
                    ret += "\v";
                }
                else if ((char)str.ElementAt(i) == '\r')
                {
                    //do nothing
                }
                else
                {
                    ret += (char)str.ElementAt(i);
                }
            }
            return ret;
        }

        private void FindindReplace(Microsoft.Office.Interop.Word.Application WordApp, object findText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object nmatchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object missing = System.Reflection.Missing.Value;
            //object replace = 2;
            object wrap = 1;
            //WordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards,
            //    ref matchSoundsLike, ref nmatchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText,
            //    ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            WordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards,
                ref matchSoundsLike, ref nmatchAllWordForms, ref forward, ref wrap, ref format, ref missing,
                ref missing, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            WordApp.Selection.Text = (string)replaceWithText;
        }

        private void buttonRemoveSpecimen_Click(object sender, RoutedEventArgs e)
        {
            methodForButtonRemoveSpecimen_Click();
        }
        private void methodForButtonRemoveSpecimen_Click()
        {
            text = speciTexts.Last(); speciTexts.Remove(text); stackPanel1.Children.Remove(text);
            label = speciLables.Last(); speciLables.Remove(label); stackPanel2.Children.Remove(label);
            buttonRemoveSpecimen.Margin = new Thickness(buttonRemoveSpecimen.Margin.Left, buttonRemoveSpecimen.Margin.Top - 30, buttonRemoveSpecimen.Margin.Right, buttonRemoveSpecimen.Margin.Bottom);
            buttonAddSpecimen.Margin = new Thickness(buttonAddSpecimen.Margin.Left, buttonAddSpecimen.Margin.Top - 30, buttonAddSpecimen.Margin.Right, buttonAddSpecimen.Margin.Bottom);
            if (speciLables.Count == 1)
            {
                label7.Content = "Specimen :";
                buttonRemoveSpecimen.Visibility = System.Windows.Visibility.Hidden;
            }
            grid2.Margin = new Thickness(grid2.Margin.Left, grid2.Margin.Top - 30, grid2.Margin.Right, grid2.Margin.Bottom);
            
        }

        private void buttonPrint_Click(object sender, RoutedEventArgs e)
        {
            labelError.Foreground = Brushes.CadetBlue;
            labelError.Content = "Saving report in the database....";
            labelError.Visibility = System.Windows.Visibility.Visible;
            labelError.UpdateLayout();
            if (!save())
            {
                return;
            }
            labelError.Foreground = Brushes.CadetBlue;
            labelError.Visibility = System.Windows.Visibility.Visible;
            labelError.Content = "Transferring report to the printer....";

            print();

            labelError.Content = "Report has been transferred to the printer....";

            timer1.Start();

        }

        private void comboBoxTitle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            methodForComboBoxTitle_SelectionChanged();
        }

        private void methodForComboBoxTitle_SelectionChanged()
        {
            switch (comboBoxTitle.SelectedIndex)
            {
                case 0:
                case 4:// Mr. and Master
                    {
                        comboBoxTitle.IsEditable = false;
                        comboBoxGender.IsEnabled = false;
                        comboBoxGender.SelectedIndex = 0;
                        break;
                    }
                case 1:
                case 2:
                case 5: // Mrs, Ms., Miss
                    {
                        comboBoxTitle.IsEditable = false;
                        comboBoxGender.IsEnabled = false;
                        comboBoxGender.SelectedIndex = 1;
                        break;
                    }
                case 3:
                case 6: // Rev, Baby
                    {
                        comboBoxTitle.IsEditable = false;
                        comboBoxGender.IsEnabled = true;
                        comboBoxGender.SelectedIndex = 0;
                        break;
                    }
                default: // other
                    {
                        comboBoxTitle.SelectedIndex = -1;
                        comboBoxTitle.IsEditable = true;
                        comboBoxGender.IsEnabled = true;
                        comboBoxGender.SelectedIndex = 0;
                        break;
                    }
            }
        }

        private void buttonSave_Click(object sender, RoutedEventArgs e)
        {
            grid1.IsEnabled = false;
            if (save())
            {
                labelError.Foreground = Brushes.CadetBlue;
                labelError.Content = "Report saved....";
                labelError.Visibility = System.Windows.Visibility.Visible;
                timer1.Start();
            }
            grid1.IsEnabled = true;
        }

        private void buttonEdit_Click(object sender, EventArgs e)
        {
            writeEnable();
            buttonSave.Content = "Save";
            buttonSave.Click -= buttonEdit_Click;
            buttonSave.Click +=new RoutedEventHandler(buttonSave_Click);
        }

        private Boolean save()
        {
            labelError.Foreground = Brushes.Red;
            labelError.Visibility = System.Windows.Visibility.Visible;
            if (textReportNo.Text == "")
            {
                labelError.Content = "Report number not entered.";
                return false;
            }
            if (textWardNo.Text == "")
            {
                labelError.Content = "Ward number not entered.";
                return false;
            }
            if (textBhtNo.Text == "")
            {
                labelError.Content = "BHT number not entered.";
                return false;
            }
            if (textPatientName.Text == "")
            {
                labelError.Content = "Patient name not entered.";
                return false;
            }
            int i;
            if (!int.TryParse(textAge.Text, out i))
            {
                labelError.Content = "Entered age is not valid.";
                return false;
            }

            for (i = 0; i < speciLables.Count; i++)
            {
                if (speciTexts.ElementAt(i).Text == "")
                {
                    if (speciLables.Count == 1)
                        labelError.Content = "Specimen is not entered.";
                    else
                        labelError.Content = "Specimen " + (i + 1) + " is not entered.";
                    return false;
                }
            }

            if (textMacroscopy.Text == "")
            {
                labelError.Content = "Macroscopy is empty.";
                return false;
            }
            if (textMicroscopy.Text == "")
            {
                labelError.Content = "Microscopy is empty.";
                return false;
            }
            if (textConclusion.Text == "")
            {
                labelError.Content = "Conclusion is empty.";
                return false;
            }

            labelError.Visibility = System.Windows.Visibility.Hidden;

            database.connectToDatabase();

            if (database.hasEntry(textReportNo.Text))
                database.deleteEntry(textReportNo.Text);

            database.store(new Record(textReportNo.Text, textWardNo.Text, textBhtNo.Text, comboBoxTitle.Text, textPatientName.Text, Int32.Parse(textAge.Text), comboBoxGender.Text, textBoxListToStringArray(speciTexts), textMacroscopy.Text, textMicroscopy.Text, textConclusion.Text, (datePicker1.DisplayDate.Date.Day + " / " + datePicker1.DisplayDate.Date.Month + " / " + datePicker1.DisplayDate.Date.Year)));

            database.closeConnection();
            updateSuggestions();
            return true;

        }

        private string[] textBoxListToStringArray(List<AutoCompleteBox> list)
        {
            String[] str = new String[list.Count];
            for (int i = 0; i < str.Length; i++)
            {
                str[i] = list.ElementAt(i).Text;
            }
            return str;
        }

        private void buttonClear_Click(object sender, RoutedEventArgs e)
        {
            clear();
        }

        private void clear()
        {
            writeEnable();
            textReportNo.Text = null;
            textWardNo.Text = null;
            textBhtNo.Text = null;
            textPatientName.Text = null;
            textAge.Text = null;
            comboBoxTitle.SelectedIndex = 0;
            comboBoxGender.SelectedIndex = 0;
            if ((string)buttonSave.Content == "Edit")
            {
                buttonSave.Content = "Save";
                buttonSave.Click -= buttonEdit_Click;
                buttonSave.Click += new RoutedEventHandler(buttonSave_Click);
            }
            //Removing the additional text boxes and lables if any
            if (speciLables.Count > 1)
            {
                int count = speciLables.Count;
                while (count > 1)
                {
                    methodForButtonRemoveSpecimen_Click();
                    count--;
                }

            }
            textSpecimen.Text = null;
            textClinicalDetails.Text = null;
            textMacroscopy.Text = null;
            textMicroscopy.Text = null;
            textConclusion.Text = null;

            datePicker1.SelectedDate = DateTime.Today;
            datePicker2.SelectedDate = DateTime.Today;
        }

        private void exitMenuItem_Click(object sender, RoutedEventArgs e)
        {

            Application.Current.Shutdown();
        }

        private void getPreviousReports_Click(object sender, RoutedEventArgs e)
        {
            if (textReportNo.Text != "")//simple search
            {
                if (retrieveReport(textReportNo.Text))
                {
                    writeDisable();
                    buttonSave.Content = "Edit";
                    buttonSave.Click -= buttonSave_Click;
                    buttonSave.Click += new RoutedEventHandler(buttonEdit_Click);
                }
                else
                {
                    MessageBox.Show("The reference number you entered did not match any report.\nPlease check the reference number and try again.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            else  // advanced search
            {
                grid3.Visibility = System.Windows.Visibility.Visible;
                textPreview.Text = "";
                if (textPatientName.Text != "")
                {
                    database.connectToDatabase();
                    DataContext = database.getRefNoSetByPartOfName(textPatientName.Text);
                    database.closeConnection();
                }
            }
        }

        private void writeDisable()
        {
            textReportNo.IsReadOnly = true; textWardNo.IsReadOnly = true; textBhtNo.IsReadOnly = true;
            comboBoxTitle.IsEnabled = false; textPatientName.IsReadOnly = true; textAge.IsReadOnly = true;
            comboBoxGender.IsEnabled = false;
            for (int i = 0; i < speciTexts.Count; i++)
            {
                speciTexts.ElementAt(i).IsEnabled = false;
            }
            buttonAddSpecimen.IsEnabled = false;
            buttonRemoveSpecimen.IsEnabled = false;
            textClinicalDetails.IsReadOnly = true; textMacroscopy.IsReadOnly = true; textMicroscopy.IsReadOnly = true; textConclusion.IsReadOnly = true;
            comboBoxSeverity.IsEnabled = false;
            datePicker1.IsEnabled = false;
            datePicker2.IsEnabled = false;
        }

        private void writeEnable()
        {
            textReportNo.IsReadOnly = false; textWardNo.IsReadOnly = false; textBhtNo.IsReadOnly = false;
            comboBoxTitle.IsEnabled = true; textPatientName.IsReadOnly = false; textAge.IsReadOnly = false;
            methodForComboBoxTitle_SelectionChanged();
            for (int i = 0; i < speciTexts.Count; i++)
            {
                speciTexts.ElementAt(i).IsEnabled = true;
            }
            buttonAddSpecimen.IsEnabled = true;
            buttonRemoveSpecimen.IsEnabled = true;
            textClinicalDetails.IsReadOnly = false; textMacroscopy.IsReadOnly = false; textMicroscopy.IsReadOnly = false; textConclusion.IsReadOnly = false;
            comboBoxSeverity.IsEnabled = true;
            datePicker1.IsEnabled = true;
            datePicker2.IsEnabled = true;
        }


        private void updateSuggestions()
        {
            Boolean added = false;
            for (int i = 0; i < speciTexts.Count; i++)
            {
                if (addToSuggestions(speciTexts.ElementAt(i).Text))
                    added = true;
            }
            if (added)
            {
                textSpecimen.ItemsSource=spSg.specimen;
                spSg.store();
            }
        }

        private Boolean addToSuggestions(String str)
        {
            if (!spSg.specimen.Contains(str))
            {
                suggestions = spSg.specimen.ToList();
                suggestions.Add(str);
                spSg.specimen = suggestions.ToArray();
                return true;
            }
            return false;
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(!((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).complete)
            {
            database.connectToDatabase();
            database.getTheRest((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item);
            database.closeConnection();
            }
            textPreview.Text="";
            textPreview.Text += "Macroscopy :\n"+((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).macroscopy+"\n\n";
            textPreview.Text += "Microscopy :\n" + ((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).microscopy + "\n\n";
            textPreview.Text += "Conclusion :\n" + ((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).conclusion ;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }



    }

    [Serializable()]
    public class specimenSuggestions : ISerializable
    {
        public String[] specimen;

        public specimenSuggestions()
        {
            specimen = null;
        }

        public specimenSuggestions(SerializationInfo info, StreamingContext ctxt)
        {
            specimen = (String[])info.GetValue("specimen", typeof(String[]));
        }

        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            info.AddValue("specimen", specimen);
        }

        public void store()
        {
            Stream stream = File.Open("specimenList.cse", FileMode.Create);
            BinaryFormatter bformatter = new BinaryFormatter();

            Console.WriteLine("Writing serialized data into file");
            bformatter.Serialize(stream, this);
            stream.Close();
        }

        public static specimenSuggestions retrieve()
        {
            Stream stream = File.Open("specimenList.cse", FileMode.Open);
            BinaryFormatter bformatter = new BinaryFormatter();

            Console.WriteLine("Reading serialized data from file");
            specimenSuggestions spSg = (specimenSuggestions)bformatter.Deserialize(stream);
            stream.Close();
            return spSg;
        }

    }

    public class Record
    {
        public Boolean complete = false;
        public String name { set; get; }
        public String gender { set; get; }
        public String date { set; get; }
        public String macroscopy;
        public String microscopy;
        public String conclusion;
        public String reportNo { set; get; }
        public String bht { set; get; }
        public String ward { set; get; }
        public String title;
        public String[] specimen;
        public int age { set; get; }


        public Record(String reportNo, String ward, String bht, String title, String name, int age, String gender, String[] specimen, String macroscopy, String microscopy,
                                                        String conclusion, String date)
        {

            this.name = name;
            this.age = age;
            this.gender = gender;
            this.date = date;
            this.macroscopy = macroscopy;
            this.microscopy = microscopy;
            this.conclusion = conclusion;
            this.reportNo = reportNo;
            this.specimen = specimen;
            this.ward = ward;
            this.bht = bht;
            this.title = title;
        }
        public static String[] StringToArray(String line)
        {
            return line.Split('∆');
        }

        public static String ArrayToString(String[] list)
        {
            String line = "";
            for (int i = 0; i < list.Length; i++)
            {
                line += list[i];
                if (i < (list.Length - 1))
                    line += "∆";
            }
            return line;
        }
    }


    class Database
    {
        OleDbConnection MyConn;
        String databasePath;

        public Database(String path)
        {
            databasePath = path;
        }

        public void connectToDatabase()
        {
            string ConnStr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databasePath + ";Persist Security Info=False;";
            MyConn = new OleDbConnection(ConnStr);
            MyConn.Open();
        }

        public void closeConnection()
        {
            MyConn.Close();
        }

        public Record getReport(String refNo)
        {
            string StrCmd = "SELECT * FROM Table1 WHERE reportNo = '" + refNo + "'";
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
            if (ObjReader == null)
            {
                Console.WriteLine();////////////////////////////////////error in connecting to the database
                return null;
            }
            else
            {

                String name, gender, specimen, date, macroscopy, microscopy, conclusion, reportNo, bht, ward, title;
                int age;
                ObjReader.Read();
                name = ObjReader["patientName"].ToString();
                gender = ObjReader["gender"].ToString();
                specimen = ObjReader["specimen"].ToString();
                ward = ObjReader["ward"].ToString();
                bht = ObjReader["bht"].ToString();
                title = ObjReader["title"].ToString();
                date = ObjReader["testDate"].ToString();
                macroscopy = ObjReader["macroscopy"].ToString();
                microscopy = ObjReader["microscopy"].ToString();
                conclusion = ObjReader["conclusion"].ToString();
                reportNo = ObjReader["reportNo"].ToString();
                age = Int32.Parse(ObjReader["age"].ToString());
                Record p = new Record(reportNo, ward, bht, title, name, age, gender, Record.StringToArray(specimen), macroscopy, microscopy, conclusion, date);
                p.complete = true;
                return p;
            }
        }

        public HashSet<String> getRefNoSetByDate(String date)
        {
            HashSet<String> refNoSet = new HashSet<String>();
            string StrCmd = "SELECT * FROM Table1 WHERE date = '" + date + "'";
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
            while (ObjReader.Read())
            {
                refNoSet.Add(ObjReader["referenceNo"].ToString());
            }
            return refNoSet;
        }

        public List<Record> getRefNoSetByPartOfName(String partOfTheName)
        {
            List<Record> refNoSet = new List<Record>();
            string StrCmd = "SELECT reportNo, patientName,gender, ward, bht, age, testDate FROM Table1";
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
             String name, gender,  date, reportNo, bht, ward;
                int age;
                
            while (ObjReader.Read())
            {
                if (ObjReader["patientName"].ToString().Contains(partOfTheName))
                {

                    name = ObjReader["patientName"].ToString();
                gender = ObjReader["gender"].ToString();
                ward = ObjReader["ward"].ToString();
                bht = ObjReader["bht"].ToString();
                date = ObjReader["testDate"].ToString();
                reportNo = ObjReader["reportNo"].ToString();
                age = Int32.Parse(ObjReader["age"].ToString());
                refNoSet.Add(new Record(reportNo, ward, bht, "", name, age, gender, new String[] { }, "", "", "", date));
                   
                }
            }
            return refNoSet;
        }

        public void getTheRest(Record record)
        {
            string StrCmd = "SELECT specimen, title, macroscopy, microscopy, conclusion FROM Table1  WHERE reportNo = '" + record.reportNo + "'" ;
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
            ObjReader.Read();
            record.specimen = Record.StringToArray(ObjReader["specimen"].ToString());
            record.title = ObjReader["title"].ToString();
            record.macroscopy = ObjReader["macroscopy"].ToString();
            record.microscopy = ObjReader["microscopy"].ToString();
            record.conclusion = ObjReader["conclusion"].ToString();
            record.complete = true;
        }

        public void store(Record record)
        {
            OleDbCommand Cmd = new OleDbCommand("INSERT INTO Table1 ( reportNo, ward,bht,title, patientName, age, gender, specimen,macroscopy, microscopy, conclusion,testDate ) VALUES ('" + record.reportNo + "'," + "'" + record.ward + "'," + "'" + record.bht + "'," + "'" + record.title + "'," + "'" + record.name + "',"
                + "'" + record.age + "'," + "'" + record.gender + "'," + "'" + Record.ArrayToString(record.specimen) + "'," + "'" + record.macroscopy + "'," + "'" + record.microscopy + "'," + "'" + record.conclusion + "'," + "'" + record.date + "')", MyConn); ;
            //OleDbCommand Cmd = new OleDbCommand("INSERT INTO Table1 ( name) VALUES ('" +record.microscopy + "')", MyConn); 

            Cmd.ExecuteNonQuery();
        }

        public bool hasEntry(String reportNo)
        {
            OleDbCommand cmdCheck = new OleDbCommand("SELECT COUNT(*) FROM Table1 WHERE reportNo = '" + reportNo + "'", MyConn);
            if (Convert.ToInt32(cmdCheck.ExecuteScalar()) == 0)
            {
                cmdCheck.ExecuteNonQuery();
                return false;
            }
            cmdCheck.ExecuteNonQuery();
            return true;
        }

        public void deleteEntry(String reportNo)
        {
            OleDbCommand cmdCheck = new OleDbCommand("DELETE FROM Table1 WHERE reportNo= '" + reportNo + "'", MyConn);//= "+user_id", MyConn);
            cmdCheck.ExecuteNonQuery();
        }
    }


}
