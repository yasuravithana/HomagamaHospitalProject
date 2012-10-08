
/*For the database and word template Password=CSE'10_CSR
 software password = homagama
 */



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
using System.Globalization;

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
        
        String draftsPath = @"Temp\Drafts List.dft";
        String templatePath = @"Data\Histopathology Report Template.docx";
        String printPath = @"Temp\Print.docx";
        public static String specimensPath = @"Temp\Specimens.cse";
        public static String databasePath = @"Data\Database.mdb";

        public static bool printReqD, printTestedD, canceled;

        Database database = new Database(MainWindow.databasePath);        

        Record draftWorkingOn = null;

        CultureInfo cultureInfo;//**
        TextInfo textInfo;//**
        int selectedList;           //0-all, 1-other
        int total, start, end; //variables used in numbering the search results 
        public static string topdate;
        public static string bottomdate;
        public static int topid;
        public static int bottomid;
        public static bool hasmore;
        public static int listsize=10;
        public static String template="";
        public static String searchPhrase;
        String[] templates;
        String templField;
        WindowTemplates wintemp;
        public static DispatcherTimer timer2 = new DispatcherTimer();

        private DispatcherTimer timer1 = new DispatcherTimer();
        
        public MainWindow()
        {
           
            InitializeComponent();
            textBoxPassword.Focus();

            textClinicalDetails.SpellCheck.IsEnabled = true;
            textMacroscopy.SpellCheck.IsEnabled = true;
            textMicroscopy.SpellCheck.IsEnabled = true;
            textConclusion.SpellCheck.IsEnabled = true;            
            speciLables.Add(label7);
            speciTexts.Add(textSpecimen);
            comboBoxTitle.SelectedIndex = 0;
            comboBoxGender.SelectedIndex = 0;
            comboBoxSeverity.SelectedIndex = 0;                                          //**
            buttonTemplate1.Visibility = System.Windows.Visibility.Hidden;
            buttonTemplate2.Visibility = System.Windows.Visibility.Hidden;
            buttonTemplate3.Visibility = System.Windows.Visibility.Hidden;
            buttonTemplate4.Visibility = System.Windows.Visibility.Hidden;

            datePicker1.SelectedDate = DateTime.Today;
            datePicker2.SelectedDate = DateTime.Today;
            cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;//**
            textInfo = cultureInfo.TextInfo;//**

            spSg = new specimenSuggestions();

            timer1.Interval = new TimeSpan(0, 0, 4);
            timer1.Tick += new EventHandler(timer1_Elapsed);
            timer2.Interval = new TimeSpan(0, 0, 1);
            timer2.Tick += new EventHandler(timer2_Elapsed);

            if (!File.Exists(MainWindow.specimensPath))
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
            methodForButtonAddSpecimen_Click();
        }
        private void methodForButtonAddSpecimen_Click()
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
            label.Margin = new Thickness(0, 4, 0, 0);
            text.Margin = new Thickness(0, 4, 0, 0);
            grid2.Margin = new Thickness(grid2.Margin.Left, grid2.Margin.Top + 30, grid2.Margin.Right, grid2.Margin.Bottom);
            buttonAddSpecimen.Margin = new Thickness(buttonAddSpecimen.Margin.Left, buttonAddSpecimen.Margin.Top + 30, buttonAddSpecimen.Margin.Right, buttonAddSpecimen.Margin.Bottom);
            buttonRemoveSpecimen.Margin = new Thickness(buttonRemoveSpecimen.Margin.Left, buttonRemoveSpecimen.Margin.Top + 30, buttonRemoveSpecimen.Margin.Right, buttonRemoveSpecimen.Margin.Bottom);

        }

        /*private Boolean retrieveReport(String reportNo)
        {
            database.connectToDatabase();
            if (database.hasEntry(reportNo, "reportNo"))
            {
                Record tempRecord = database.getReport(reportNo);                
                database.closeConnection();
                fillFields(tempRecord);
                return true;
            }
            else
            {
                database.closeConnection();
                return false;
            }
        }*/

        private void fillFields(Record record)
        {
            clear();
            textReportNo.Text = record.Reference_No;
            textPatientName.Text = record.Name;
            textWardNo.Text = record.Ward;
            textBhtNo.Text = record.BHT;
            comboBoxTitle.SelectedIndex = -1;
            comboBoxTitle.Text = record.title;
            if (comboBoxTitle.Text == "Baby")
            {
                textMonth.Visibility = System.Windows.Visibility.Visible;
                labelMonths.Visibility = System.Windows.Visibility.Visible;
                textMonth.Text = record.months.ToString();
            }
            textPatientName.Text = record.Name;
            textAge.Text = record.years.ToString();
            comboBoxGender.Text = record.Gender;
            textSpecimen.Text = record.specimenArray[0];
            comboBoxSeverity.Text = record.severity;        //**
            textClinicalDetails.Text = record.clinicalDetails;  //**
            for (int i = 1; i < record.specimenArray.Length; i++)
            {
                methodForButtonAddSpecimen_Click();
                speciTexts.Last().Text =record.specimenArray[i];
            }

            textMacroscopy.Text = record.macroscopy;
            textMicroscopy.Text = record.microscopy;
            textConclusion.Text = record.conclusion;

            String[] dateformat2 = record.TestDate.Split('/');                                                                                  //**
            datePicker2.SelectedDate = new DateTime(Int32.Parse(dateformat2[0]), Int32.Parse(dateformat2[1]), Int32.Parse(dateformat2[2]));     //**
            String[] dateformat1 = record.requestDate.Split('/');                                                                                  //**
            datePicker1.SelectedDate = new DateTime(Int32.Parse(dateformat1[0]), Int32.Parse(dateformat1[1]), Int32.Parse(dateformat1[2]));     //**
        }


        private void print()
        {
            object FileName = AppDomain.CurrentDomain.BaseDirectory + "\\" + templatePath;//
            object saveAs = printPath;
            object password = "CSE'10_CSR";
            object noPassword = "";
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Options.set_DefaultFilePath(0, AppDomain.CurrentDomain.BaseDirectory);
            Microsoft.Office.Interop.Word.Document aDoc = null;
            object readOnly = true;
            object isVisible = false;
            wordApp.Visible = false;
            aDoc = wordApp.Documents.Open(ref FileName, ref missing, ref readOnly, ref missing, ref password,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible,
                 ref missing, ref missing, ref missing, ref missing);
            aDoc.Activate();

            FindindReplace(wordApp, "<name>", comboBoxTitle.Text + " " + textPatientName.Text);
            if (comboBoxTitle.Text == "Baby")
                FindindReplace(wordApp, "<age>", textAge.Text + " years " + textMonth.Text + " months");
            else
                FindindReplace(wordApp, "<age>", textAge.Text + " years ");
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
            if (textClinicalDetails.Text=="")                                   //**
                FindindReplace(wordApp, "<clinical>", replaceNewLines(""));    //**
            else                                                              //**
                FindindReplace(wordApp, "<clinical>", replaceNewLines("\n\nClinical Details	: " + textClinicalDetails.Text));//**
            FindindReplace(wordApp, "<macro>", replaceNewLines(textMacroscopy.Text));
            FindindReplace(wordApp, "<micro>", replaceNewLines(textMicroscopy.Text));
            FindindReplace(wordApp, "<con>", replaceNewLines(textConclusion.Text));
            String dates = "";
            if (printReqD)
            {
                dates += "\nRequested on : " + datePicker1.SelectedDate.Value.Date.Day + " / " + datePicker1.SelectedDate.Value.Date.Month + " / " + datePicker1.SelectedDate.Value.Date.Date.Year + "\t";
            }
            if (printTestedD)
            {
                dates += "Tested on : " + datePicker2.SelectedDate.Value.Date.Day + " / " + datePicker2.SelectedDate.Value.Date.Month + " / " + datePicker2.SelectedDate.Value.Date.Date.Year;
            }
            FindindReplace(wordApp, "<dates>", dates);
           // FindindReplace(wordApp, "<date>", datePicker2.SelectedDate.Value.Date.Day + " / " + datePicker2.SelectedDate.Value.Date.Month + " / " + datePicker2.SelectedDate.Value.Date.Date.Year);
            //FindindReplace(wordApp, "<reqdate>", datePicker1.SelectedDate.Value.Date.Day + " / " + datePicker1.SelectedDate.Value.Date.Month + " / " + datePicker1.SelectedDate.Value.Date.Date.Year);//**
            FindindReplace(wordApp, "<printdate>", DateTime.Today.Date.Day + " / " + DateTime.Today.Date.Month + " / " + DateTime.Today.Date.Year);//**

            aDoc.SaveAs(ref saveAs, ref missing, ref missing, ref noPassword, ref missing, ref missing,
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
            String ret = "";
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

            
            Stream FileStream = null;
            if (grid3.Visibility == System.Windows.Visibility.Visible)
            {
                if ((Record)dataGrid1.SelectedItem == null)
                {
                    MessageBox.Show("Please select a report first.", "", MessageBoxButton.OK, MessageBoxImage.Stop);
                    return;
                }
            }

            if (grid3.Visibility == System.Windows.Visibility.Hidden)
            {
                labelError.Foreground = Brushes.CadetBlue;
                labelError.Content = "Saving report in the database....";
                labelError.Visibility = System.Windows.Visibility.Visible;
                if (!save())
                {
                    return;
                }
            }

            if (draftWorkingOn != null)
            {
                ((List<Record>)dataGridDraftsList.DataContext).Remove(draftWorkingOn);
                draftWorkingOn = null;
                DraftList list = new DraftList();
                list.list = (List<Record>)dataGridDraftsList.DataContext;
                FileStream = File.Create(draftsPath);
                BinaryFormatter serializer = new BinaryFormatter();
                serializer.Serialize(FileStream, list);
                FileStream.Close();
            }

            if (grid3.Visibility == System.Windows.Visibility.Hidden)
            {
                labelError.Foreground = Brushes.CadetBlue;
                labelError.Visibility = System.Windows.Visibility.Visible;
                labelError.Content = "Report Saved...";
            }

            Window1 printC = new Window1();
            printC.Owner = this;
            printC.Left = this.Left;
            printC.Top = this.Top;
            printC.ShowDialog();

            if (!canceled)
            {
                labelError.Foreground = Brushes.CadetBlue;
                labelError.Visibility = System.Windows.Visibility.Visible;
                labelError.Content = "Transferring report to the printer....";

                print();

                labelError.Content = "Report has been transferred to the printer....";
            }

            timer1.Start();

        }

        private void comboBoxTitle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            methodForComboBoxTitle_SelectionChanged();
        }

        private void methodForComboBoxTitle_SelectionChanged()
        {
            if(textAge.Text == "0")
                textAge.Text = "";
            textMonth.Text = "";
            textMonth.Visibility = System.Windows.Visibility.Hidden;
            labelMonths.Visibility = System.Windows.Visibility.Hidden;
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
                case 3://Rev 
                    {
                        comboBoxTitle.IsEditable = false;
                        comboBoxGender.IsEnabled = true;
                        comboBoxGender.SelectedIndex = 0;
                        break;
                    }
                case 6: // Baby
                    {
                        comboBoxTitle.IsEditable = false;
                        comboBoxGender.IsEnabled = true;
                        comboBoxGender.SelectedIndex = 0;
                        textAge.Text = "0";
                        textMonth.Text = "0";
                        textMonth.Visibility = System.Windows.Visibility.Visible;
                        labelMonths.Visibility = System.Windows.Visibility.Visible;
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
            Stream FileStream = null;

            grid1.IsEnabled = false;
            if (save())
            {
                labelError.Foreground = Brushes.CadetBlue;
                labelError.Content = "Report saved....";
                labelError.Visibility = System.Windows.Visibility.Visible;
                timer1.Start();
                
                
                if (draftWorkingOn != null)
                {
                    ((List<Record>)dataGridDraftsList.DataContext).Remove(draftWorkingOn);
                    draftWorkingOn = null;
                    DraftList list = new DraftList();
                    list.list = (List<Record>)dataGridDraftsList.DataContext;
                    FileStream = File.Create(draftsPath);
                    BinaryFormatter serializer = new BinaryFormatter();
                    serializer.Serialize(FileStream, list);
                    FileStream.Close();
                }


            }
            grid1.IsEnabled = true;
        }

        private void buttonEdit_Click(object sender, RoutedEventArgs e)
        {
            //writeEnable();
            textReportNo.IsEnabled = false;
            buttonDelete.Visibility = System.Windows.Visibility.Visible;
            buttonSave.Content = "Save";
            buttonSave.Click -= buttonEdit_Click;
            buttonSave.Click += new RoutedEventHandler(buttonSave_Click);
        }

        private Boolean save()
        {
            labelError.Foreground = Brushes.Red;
            labelError.Visibility = System.Windows.Visibility.Visible;
            if (textReportNo.Text == "")
            {
                labelError.Content = "Reference number not entered.";
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
                labelError.Content = "Entered age(year) is not valid.";
                return false;
            }
            if (comboBoxTitle.Text == "Baby")
            {
                if (! int.TryParse(textMonth.Text, out i))
                {
                    labelError.Content = "Entered age(month) is not valid.";
                    return false;
                }
                else if (Int32.Parse(textMonth.Text) > 11 || Int32.Parse(textMonth.Text) < 0)
                {
                    labelError.Content = "Entered age(month) should be between 0 & 11";
                    return false;
                }
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
                else      //**
                    speciTexts.ElementAt(i).Text=(textInfo.ToTitleCase(speciTexts.ElementAt(i).Text));//**
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
            int months = 0;
            if (comboBoxTitle.Text == "Baby")
                months = Int32.Parse(textMonth.Text);
            database.connectToDatabase();
            if (!database.hasEntry(textReportNo.Text,"reportNo"))
            {
                database.store(new Record(textReportNo.Text, textWardNo.Text, textBhtNo.Text, comboBoxTitle.Text, textPatientName.Text, Int32.Parse(textAge.Text), months, comboBoxGender.Text, textBoxListToStringArray(speciTexts), textMacroscopy.Text, textMicroscopy.Text, textConclusion.Text, dateToString(datePicker2.SelectedDate.Value), dateToString(datePicker1.SelectedDate.Value), comboBoxSeverity.Text, textClinicalDetails.Text));

                database.closeConnection();
                updateSuggestions();
                return true;
            }
            else
            {
                if (MessageBox.Show("A report with this reference number already exists. Do you want to replace it?", "Warning!", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    database.deleteEntry(textReportNo.Text);
                                                                    //**
                    database.store(new Record(textReportNo.Text, textWardNo.Text, textBhtNo.Text, comboBoxTitle.Text, textPatientName.Text, Int32.Parse(textAge.Text), months, comboBoxGender.Text, textBoxListToStringArray(speciTexts), textMacroscopy.Text, textMicroscopy.Text, textConclusion.Text, dateToString(datePicker2.SelectedDate.Value), dateToString(datePicker1.SelectedDate.Value), comboBoxSeverity.Text, textClinicalDetails.Text));

                    database.closeConnection();
                    updateSuggestions();
                    return true;
                }
                else
                {
                    database.closeConnection();
                    return false;
                }
            }

        }

        private String dateToString(DateTime date)
        {
            String ret=date.Year + " / " ;            
            if(date.Month>9)
            {
                ret += date.Month + " / ";
            }
            else
            {
                ret += "0" + date.Month + " / ";
            }

            if (date.Day > 9)
            {
                ret += date.Day;
            }
            else
            {
                ret += "0" + date.Day;
            }
            return ret;
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
            draftWorkingOn = null;
        }

        private void clear()
        {
            //writeEnable();
            textReportNo.IsEnabled = true;
            buttonDelete.Visibility = System.Windows.Visibility.Hidden;
            if (!grid3.IsVisible)
                menu.IsEnabled = true;

            textReportNo.Text = "";
            textWardNo.Text = "";
            textBhtNo.Text = "";
            textPatientName.Text = "";
            textAge.Text = "";
            textMonth.Text = "";
            textMonth.Visibility = System.Windows.Visibility.Hidden;
            labelMonths.Visibility = System.Windows.Visibility.Hidden;
            comboBoxTitle.SelectedIndex = 0;
            comboBoxGender.SelectedIndex = 0;
            comboBoxSeverity.SelectedIndex = 0;//**
            
            menuItemGetReport.IsEnabled = true;

            if ((string)buttonSave.Content == "Edit")
            {
                buttonSave.Content = "Save";
                buttonSave.Click -= buttonEdit_Click;
                buttonSave.Click -= buttonEditInPreview_Click;
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
            textSpecimen.Text = "";
            textClinicalDetails.Text = "";
            textMacroscopy.Text = "";
            textMicroscopy.Text = "";
            textConclusion.Text = "";

            datePicker1.SelectedDate = DateTime.Today;
            datePicker2.SelectedDate = DateTime.Today;

            

        }

        private void exitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void getPreviousReports_Click(object sender, RoutedEventArgs e)
        {
            searchSpecimen.ItemsSource = spSg.specimen;

            searchRefNo.Text = "";
            searchBht.Text = "";
            searchName.Text = "";
            searchWardNo.Text = "";
            searchSpecimen.Text = "";
            searchSeverity.SelectedIndex = -1;
            checkBoxSearchPeriod.IsChecked = false;
            searchFrom.IsEnabled = false;
            searchTo.IsEnabled = false;
            dataGrid1.DataContext = null;



            grid3.Visibility = System.Windows.Visibility.Visible;
            menu.IsEnabled = false;

            buttonSave.Content = "Edit";
            buttonSave.Click -= buttonSave_Click;
            buttonSave.Click += new RoutedEventHandler(buttonEditInPreview_Click);

            buttonClear.Content = "Home";
            buttonClear.Click -= buttonClear_Click;
            buttonClear.Click += new RoutedEventHandler(buttonHome_Click);

            dataGrid1.SelectedIndex = -1;
            textPreview.Text = "";
            labelSearchCount.Content = "";

            
        }

        private void buttonNext_Click(object sender, RoutedEventArgs e)
        {
            database.connectToDatabase();
            if (selectedList == 0)
                dataGrid1.DataContext = database.getAllRecordList(1);
            else
                dataGrid1.DataContext = database.getRecordList(1);
            database.closeConnection();
            if (!hasmore)
                buttonNext.IsEnabled = false;
            buttonBack.IsEnabled = true;
            start = end + 1;
            if (end + MainWindow.listsize <= total)
                end += MainWindow.listsize;
            else
                end = total;
            updateSearchCount();
        }

        private void buttonBack_Click(object sender, RoutedEventArgs e)
        {
            database.connectToDatabase();
            if (selectedList == 0)
                dataGrid1.DataContext = database.getAllRecordList(0);
            else
                dataGrid1.DataContext = database.getRecordList(0);
            
            database.closeConnection();
            if (!hasmore)
                buttonBack.IsEnabled = false;
            buttonNext.IsEnabled = true;
            start -= MainWindow.listsize;
            end = start + MainWindow.listsize - 1;
            updateSearchCount();
        }

        private void buttonHome_Click(object sender, RoutedEventArgs e)
        {
            methodForbuttonHome_Click();
            buttonNext.Visibility = System.Windows.Visibility.Hidden;
            buttonBack.Visibility = System.Windows.Visibility.Hidden;
        }

        private void methodForbuttonHome_Click()
        {
            grid3.Visibility = System.Windows.Visibility.Hidden;
            clear();
            buttonClear.Content = "Clear";
            buttonClear.Click -= buttonHome_Click;
            buttonClear.Click += new RoutedEventHandler(buttonClear_Click);
            dataGrid1.DataContext = null;
        }

        private void buttonEditInPreview_Click(object sender, RoutedEventArgs e)
        {
            if ((Record)dataGrid1.SelectedItem != null)
            {
                textReportNo.IsEnabled = false;
                buttonDelete.Visibility = System.Windows.Visibility.Visible;
                grid3.Visibility = System.Windows.Visibility.Hidden;
                buttonClear.Content = "Clear"; 
                buttonClear.Click -= buttonHome_Click; 
                buttonClear.Click += new RoutedEventHandler(buttonClear_Click);
                //writeEnable();
                buttonSave.Content = "Save";
                buttonSave.Click -= buttonEditInPreview_Click;
                buttonSave.Click += new RoutedEventHandler(buttonSave_Click);
                //dataGrid1.DataContext = null;
                buttonNext.Visibility = System.Windows.Visibility.Hidden;
                buttonBack.Visibility = System.Windows.Visibility.Hidden;
            }
            else
            {
                MessageBox.Show("Please select a report first.", "", MessageBoxButton.OK, MessageBoxImage.Stop);
            }
        }

        private void updateSearchCount()
        {
            labelSearchCount.Content = "Showing "+start+" - "+end+" out of "+total+" total results";
        }

        /*private void writeDisable()
        {
            textReportNo.IsReadOnly = true; 
            textWardNo.IsReadOnly = true; textBhtNo.IsReadOnly = true;
            comboBoxTitle.IsEnabled = false; textPatientName.IsReadOnly = true; textAge.IsReadOnly = true;
            textMonth.IsReadOnly = true;
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
        }*/

        /*private void writeEnable()
        {
            textReportNo.IsReadOnly = false; 
            textWardNo.IsReadOnly = false; textBhtNo.IsReadOnly = false;
            comboBoxTitle.IsEnabled = true; textPatientName.IsReadOnly = false; textAge.IsReadOnly = false;
            textMonth.IsReadOnly = false;
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
        }*/


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
                textSpecimen.ItemsSource = spSg.specimen;
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
            if ((((System.Windows.Controls.DataGrid)sender).CurrentCell).IsValid)
            {
                if (!((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).complete)
                {
                    database.connectToDatabase();
                    database.getTheRest((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item);
                    database.closeConnection();
                }
                textPreview.Text = "";
                if (((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).clinicalDetails == "")//**
                    textPreview.Text += "Clinical Details :\n - \n\n";                                          //**
                else                                                                                            //**
                    textPreview.Text += "Clinical Details :\n" + ((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).clinicalDetails + "\n\n"; //**
                textPreview.Text += "Macroscopy :\n" + ((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).macroscopy + "\n\n";
                textPreview.Text += "Microscopy :\n" + ((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).microscopy + "\n\n";
                textPreview.Text += "Conclusion :\n" + ((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item).conclusion;

                fillFields((Record)((System.Windows.Controls.DataGrid)sender).CurrentCell.Item);
                buttonSave.Content = "Edit";
                buttonSave.Click -= buttonSave_Click;
                buttonSave.Click += new RoutedEventHandler(buttonEditInPreview_Click);
            }

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void toTitlecase(object sender, TextChangedEventArgs e)
        {
            int position = ((TextBox)sender).SelectionStart;
            ((TextBox)sender).Text = textInfo.ToTitleCase(((TextBox)sender).Text);
            ((TextBox)sender).SelectionStart = position;
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to delete this report?", "Delete report!", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                database.connectToDatabase();
                database.deleteEntry(textReportNo.Text);
                database.closeConnection();
                clear();
            }
            else
                return;
        }

        private void buttonTemplate1_Click(object sender, RoutedEventArgs e)
        {
            methodforTemplates("clinicalDetails");
        }
        private void buttonTemplate2_Click(object sender, RoutedEventArgs e)
        {
            methodforTemplates("macroscopy");
        }
        private void buttonTemplate3_Click(object sender, RoutedEventArgs e)
        {
            methodforTemplates("microscopy");
        }
        private void buttonTemplate4_Click(object sender, RoutedEventArgs e)
        {
            methodforTemplates("conclusion");
        }

        private void methodforTemplates(String colomn)
        {
            templField = colomn;
            database.connectToDatabase();
            templates = database.getTemplates(colomn, textSpecimen.Text);
            database.closeConnection();
            if (templates == null)
            {
                MessageBox.Show("No templates found related to this specimen", "Sorry!", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }
            this.IsEnabled = false;
            
            wintemp = new WindowTemplates();
            MainWindow.template = "";
            for (int i = 0; i < templates.Length; i++)
                wintemp.listBox1.Items.Add(new UniqueListItemObject(templates[i]));
            wintemp.ShowDialog();
            
                       
        }

        void timer2_Elapsed(object sender, EventArgs e)
        {
            if (template != "")
            {
                if (templField == "clinicalDetails")
                    textClinicalDetails.Text = template;
                else if (templField == "macroscopy")
                    textMacroscopy.Text = template;
                else if (templField == "microscopy")
                    textMicroscopy.Text = template;
                else if (templField == "conclusion")
                    textConclusion.Text = template;
            }
            this.IsEnabled = true;
            timer2.Stop();
        }

        private void textSpecimen_TextChanged(object sender, RoutedEventArgs e)
        {
            methodforTextSpecimenChanged();
        }

        private void methodforTextSpecimenChanged()
        {
            if (textSpecimen.Text == "")
            {
                buttonTemplate1.Visibility = buttonTemplate2.Visibility = buttonTemplate3.Visibility =
                buttonTemplate4.Visibility = System.Windows.Visibility.Hidden;
            }
            else
            {
                if (textSpecimen.IsEnabled)
                {
                    buttonTemplate1.Visibility = buttonTemplate2.Visibility = buttonTemplate3.Visibility =
                buttonTemplate4.Visibility = System.Windows.Visibility.Visible;
                }
                else
                {
                    buttonTemplate1.Visibility = buttonTemplate2.Visibility = buttonTemplate3.Visibility =
                    buttonTemplate4.Visibility = System.Windows.Visibility.Hidden;
                }
            }
        }

        private void textSpecimen_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            methodforTextSpecimenChanged();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            /*
            if (wintemp != null)
                wintemp.Close();
            */
            if (MessageBox.Show("Are you sure that you want to exit? All unsaved data will be lost.", "Confim!", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
            }
            else
                e.Cancel = true;
        }

        private void searchRefNo_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (searchRefNo.Text == "")
            {
                searchBht.IsEnabled = true;
                searchName.IsEnabled = true;
                searchWardNo.IsEnabled = true;
                searchSpecimen.IsEnabled = true;
                searchSeverity.IsEnabled = true;
                checkBoxSearchPeriod.IsEnabled = true;
                checkBoxSearchPeriod.IsChecked = false;

            }
            else if(searchRefNo.Text.Length==1)
            {
                searchBht.IsEnabled = false;
                searchName.IsEnabled = false;
                searchWardNo.IsEnabled = false;
                searchSpecimen.IsEnabled = false;
                searchSeverity.IsEnabled = false;
                checkBoxSearchPeriod.IsEnabled = false;
                checkBoxSearchPeriod.IsChecked = false;
            }

        }

        private void buttonSearch_Click(object sender, RoutedEventArgs e)
        {
            int count=0;
            topdate = bottomdate = "";
            topid = bottomid = 0;
            hasmore = false;
            MainWindow.searchPhrase = "";

            textPreview.Text = "";

            database.connectToDatabase();

            if (searchRefNo.Text == "")
            {
                
                if (searchWardNo.Text != "")
                {
                    if (count != 0)
                        MainWindow.searchPhrase += " AND";
                    MainWindow.searchPhrase += " ward = '" + searchWardNo.Text + "'";
                    count++;
                }
                if (searchBht.Text != "")
                {
                    if (count != 0)
                        MainWindow.searchPhrase += " AND";
                    MainWindow.searchPhrase += " bht = '" + searchBht.Text + "'";
                    count++;
                }
                if (searchName.Text != "")
                {
                    if (count != 0)
                        MainWindow.searchPhrase += " AND";
                    MainWindow.searchPhrase += " patientName LIKE '%" + searchName.Text + "%'";
                    count++;
                }
                if (searchSpecimen.Text != "")
                {
                    if (count != 0)
                        MainWindow.searchPhrase += " AND";
                    MainWindow.searchPhrase += " specimen LIKE '%" + searchSpecimen.Text + "%'";
                    count++;
                }

                if (searchSeverity.SelectedIndex != -1)
                {
                    if (count != 0)
                        MainWindow.searchPhrase += " AND";
                    MainWindow.searchPhrase += " severity = '" + searchSeverity.Text + "'";
                    count++;
                }

                if (checkBoxSearchPeriod.IsChecked==true)
                {
                    if (count != 0)
                        MainWindow.searchPhrase += " AND";
                    MainWindow.searchPhrase += " testDate >= '" + dateToString(searchFrom.SelectedDate.Value) + "' AND testDate <= '" + dateToString(searchTo.SelectedDate.Value) + "'";
                    count++;
                }

                if (count != 0)
                {
                    selectedList = 1;
                    dataGrid1.DataContext = database.getRecordList(1);//1-next,0-back
                    if (dataGrid1.Items.Count == 0)
                        MessageBox.Show("There are no reports that satisfy the conditions.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
                else
                {

                    if (database.hasanyEntry())
                    {
                        selectedList = 0;
                        dataGrid1.DataContext = database.getAllRecordList(1);//1-next,0-back
                    }
                    else
                    {
                        methodForbuttonHome_Click();
                        MessageBox.Show("There are no reports stored.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
            }
            else
            {
                MainWindow.searchPhrase = " reportNo = '" + searchRefNo.Text + "'";
                selectedList = 1;
                dataGrid1.DataContext = database.getRecordList(1);//1-next,0-back
                if(dataGrid1.Items.Count==0)
                    MessageBox.Show("There are no reports that satisfy the conditions.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }

            if (hasmore)
            {
                buttonBack.Visibility = System.Windows.Visibility.Visible;
                buttonBack.IsEnabled = false;
                buttonNext.Visibility = System.Windows.Visibility.Visible;
                buttonNext.IsEnabled = true;
            }
            else
            {
                buttonBack.Visibility = System.Windows.Visibility.Hidden;
                buttonNext.Visibility = System.Windows.Visibility.Hidden;
            }

            total=database.count();
            if (total != 0)
            {
                start = 1;
                if (start + MainWindow.listsize - 1 <= total)
                    end = start + MainWindow.listsize - 1;
                else
                    end = total;
                database.closeConnection();

                updateSearchCount();
            }
            else
                labelSearchCount.Content="";

        }

        private void buttonClearSearch_Click(object sender, RoutedEventArgs e)
        {
            searchRefNo.Text = "";
            searchBht.Text = "";
            searchName.Text = "";
            searchWardNo.Text = "";
            searchSpecimen.Text = "";
            searchSeverity.SelectedIndex = -1;
            dataGrid1.DataContext = null;
            textPreview.Text="";
            labelSearchCount.Content = "";
            checkBoxSearchPeriod.IsChecked = false;
            buttonNext.Visibility = System.Windows.Visibility.Hidden;
            buttonBack.Visibility = System.Windows.Visibility.Hidden;
        }

        private void checkBoxSearchPeriod_Checked(object sender, RoutedEventArgs e)
        {
            searchFrom.IsEnabled = true;
            searchTo.IsEnabled = true;
            searchFrom.SelectedDate = DateTime.Today;
            searchTo.SelectedDate = DateTime.Today;
        }

        private void checkBoxSearchPeriod_Unchecked(object sender, RoutedEventArgs e)
        {
            searchFrom.IsEnabled = false;
            searchTo.IsEnabled = false;
        }

        private void buttonSaveDraft_Click(object sender, RoutedEventArgs e)
        {
            if (textPatientName.Text != "")
            {
                DraftList list;
                Stream FileStream = null;
                BinaryFormatter deserializer;
                if (dataGridDraftsList.DataContext == null)
                {
                    try
                    {
                        FileStream = File.OpenRead(draftsPath);
                        deserializer = new BinaryFormatter();
                        list = (DraftList)deserializer.Deserialize(FileStream);
                        FileStream.Close();
                    }
                    catch
                    {
                        if (FileStream != null)
                        {
                            FileStream.Close();
                        }
                        list = new DraftList();
                    }
                }
                else
                {
                    list = new DraftList();
                    list.list = (List<Record>)dataGridDraftsList.DataContext;
                }

                int months = 0;
                if (comboBoxTitle.Text == "Baby")
                    Int32.TryParse(textMonth.Text, out months);
                int years = 0;
                Int32.TryParse(textAge.Text, out years);

                list.list.Add(new Record(textReportNo.Text, textWardNo.Text, textBhtNo.Text, comboBoxTitle.Text, textPatientName.Text, years, months, comboBoxGender.Text, textBoxListToStringArray(speciTexts), textMacroscopy.Text, textMicroscopy.Text, textConclusion.Text, dateToString(datePicker2.SelectedDate.Value), dateToString(datePicker1.SelectedDate.Value), comboBoxSeverity.Text, textClinicalDetails.Text));
                if (draftWorkingOn != null)
                {
                    list.list.Remove(draftWorkingOn);
                    draftWorkingOn = null;
                }
                FileStream = File.Create(draftsPath);
                BinaryFormatter serializer = new BinaryFormatter();
                serializer.Serialize(FileStream, list);
                FileStream.Close();

                labelError.Foreground = Brushes.CadetBlue;
                labelError.Content = "Draft Saved successfully";
                labelError.Visibility = System.Windows.Visibility.Visible;
                timer1.Start();
            }
            else
                MessageBox.Show("Can't save a draft without a name.", "Draft Not Saved", MessageBoxButton.OK);
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            gridDraftsList.Visibility = System.Windows.Visibility.Hidden;
            menu.IsEnabled = true;
        }

        private void retrieveDraftMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Stream FileStream;
            BinaryFormatter deserializer;
            try
            {
                FileStream = File.OpenRead(draftsPath);
                deserializer = new BinaryFormatter();
                dataGridDraftsList.DataContext = (List<Record>)(((DraftList)deserializer.Deserialize(FileStream)).list);
                FileStream.Close();
                if (dataGridDraftsList.Items.Count != 0)
                {
                    gridDraftsList.Visibility = System.Windows.Visibility.Visible;
                    menu.IsEnabled = false;
                }
                else
                    MessageBox.Show("There are no drafts saved.", "Empty!", MessageBoxButton.OK, MessageBoxImage.Stop);
            }
            catch
            {
                MessageBox.Show("Error in retrieving drafts.", "Error", MessageBoxButton.OK, MessageBoxImage.Stop);
            }
        }

        private void dataGridDraftsList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            draftWorkingOn=(Record)dataGridDraftsList.SelectedItem;
            fillFields(draftWorkingOn);
            gridDraftsList.Visibility = System.Windows.Visibility.Hidden;
            menu.IsEnabled = false;
        }

        private void buttonDeleteDraft_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Are you sure that you want to delete the selected draft?", "", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                Stream FileStream;
                ((List<Record>)dataGridDraftsList.DataContext).Remove((Record)dataGridDraftsList.SelectedItem);
                DraftList list = new DraftList();
                list.list = (List<Record>)dataGridDraftsList.DataContext;
                FileStream = File.Create(draftsPath);
                BinaryFormatter serializer = new BinaryFormatter();
                serializer.Serialize(FileStream, list);
                FileStream.Close();
                dataGridDraftsList.Items.Refresh();
            }
        }

        private void buttonLogin_Click(object sender, RoutedEventArgs e)
        {
            login();
        }

        private void login()
        {
            if (textBoxPassword.Password == "")
            {
                gridLogin.Visibility = System.Windows.Visibility.Hidden;
            }
            else
            {
                MessageBox.Show("The password you entered is incorrect. Please try again", "Authentication failed", MessageBoxButton.OK, MessageBoxImage.Error);
                textBoxPassword.Password = "";
                textBoxPassword.Focus();
            }
        }

        private void textBoxPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                login();
            }
        }

        private void frameNextTest_MouseDown(object sender, MouseButtonEventArgs e)
        {
            switch (label1.Content.ToString())
            {
                case("HISTOPATHOLOGY REPORT"):
                    {
                        label1.Content = "Test1";
                        break;
                    }
                case ("Test1"):
                    {
                        label1.Content = "Test2";
                        break;
                    }
                case ("Test2"):
                    {
                        label1.Content = "HISTOPATHOLOGY REPORT";
                        break;
                    }
            }
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
            Stream stream = File.Open(MainWindow.specimensPath, FileMode.Create);
            BinaryFormatter bformatter = new BinaryFormatter();

            Console.WriteLine("Writing serialized data into file");
            bformatter.Serialize(stream, this);
            stream.Close();
        }

        public static specimenSuggestions retrieve()
        {
            Stream stream = File.Open(MainWindow.specimensPath, FileMode.Open);
            BinaryFormatter bformatter = new BinaryFormatter();

            Console.WriteLine("Reading serialized data from file");
            specimenSuggestions spSg = (specimenSuggestions)bformatter.Deserialize(stream);
            stream.Close();
            return spSg;
        }

    }

    [Serializable()]
    public class DraftList
    {
        public List<Record> list = new List<Record>();
    }

    [Serializable()]
    public class Record
    {
        public Boolean complete = false;
        public String Name { set; get; }
        public String Gender { set; get; }
        public String TestDate { set; get; }
        public String macroscopy;
        public String microscopy;
        public String conclusion;
        public String Reference_No { set; get; }
        public String BHT { set; get; }
        public String Ward { set; get; }
        public String title;
        public String[] specimenArray;
        public int years;
        public int months;
        public String Age { set; get; }
        public String Specimens { set; get; }
        public String requestDate;                  //**
        public String severity { set; get; }        //**
        public String clinicalDetails;              //**

        public Record(String reportNo, String ward, String bht, String title, String name, int years, int months, String gender, String[] specimen, String macroscopy, String microscopy,
                                                        String conclusion, String date, String requestDate, String severity, String clinicalDetails)
        {                                                                           //**

            this.Name = name;
            this.years = years;
            this.months = months;
            this.Gender = gender;
            this.TestDate = date;
            this.macroscopy = macroscopy;
            this.microscopy = microscopy;
            this.conclusion = conclusion;
            this.Reference_No = reportNo;
            this.specimenArray = specimen;
            this.Ward = ward;
            this.BHT = bht;
            this.title = title;
            this.requestDate = requestDate;         //**
            this.severity = severity;               //**
            this.clinicalDetails = clinicalDetails; //**

            Specimens = "";

            for (int i = 0; i < specimen.Length; i++)
            {
                Specimens += specimen[i];
                if (i < specimen.Length - 1)
                    Specimens += '\n';
            }

            this.Age = "";
            if (this.years != 0)
            {
                this.Age += this.years + " y ";
            }
            if (this.months != 0)
            {
                this.Age += this.months + " m";
            }

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
            string ConnStr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databasePath + ";Persist Security Info=False;Jet OLEDB:Database Password=CSE'10_CSR";
            MyConn = new OleDbConnection(ConnStr);
            MyConn.Open();
        }

        public void closeConnection()
        {
            MyConn.Close();
        }

       /* public Record getReport(String refNo) //**
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

                String name, gender, specimen, date, macroscopy, microscopy, conclusion, reportNo, bht, ward, title,requestDate, severity, clinicalDetails;
                int years,months;                                                                                            //**
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
                requestDate = ObjReader["requestDate"].ToString();  //**
                severity = ObjReader["severity"].ToString();            //**
                clinicalDetails = ObjReader["clinicalDetails"].ToString();  //**
                years = Int32.Parse(ObjReader["age"].ToString());
                months = Int32.Parse(ObjReader["months"].ToString());
                Record p = new Record(reportNo, ward, bht, title, name, years, gender, Record.StringToArray(specimen), macroscopy, microscopy, conclusion, date, requestDate, severity, clinicalDetails,months);
                p.complete = true;                                                                              //**
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
        */
        public List<Record> getAllRecordList(int option)
        {
            List<Record> refNoSet = new List<Record>();
            string StrCmd;
            switch(option)
            {
                case (0):
                    {
                        StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, months, specimen,testDate,severity FROM Table1 where testdate >= '" + MainWindow.topdate + "'  AND iif(testdate = '" + MainWindow.topdate + "' ,dateid > " + MainWindow.topid + ",1) order by testdate asc,dateid ASC";
                        break;
                    }
                case(1):
                default:
                    {
                        if (MainWindow.bottomdate != "")
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, months, specimen,testDate,severity FROM Table1 where testdate <= '" + MainWindow.bottomdate + "' AND iif(testdate = '" + MainWindow.bottomdate + "' ,dateid <= " + MainWindow.bottomid + ",1) order by testdate desc,dateid DESC";
                        else
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, months, specimen,testDate,severity FROM Table1 order by testdate desc, dateid desc";
                        break;
                    }
            }
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
            String name, gender, date, reportNo, bht, ward, severity;
            String[] specimen;
            int age, months;

            while (ObjReader.Read())
            {
                    name = ObjReader["patientName"].ToString();
                    gender = ObjReader["gender"].ToString();
                    ward = ObjReader["ward"].ToString();
                    bht = ObjReader["bht"].ToString();
                    date = ObjReader["testDate"].ToString();
                    reportNo = ObjReader["reportNo"].ToString();
                    severity = ObjReader["severity"].ToString();        //**
                    age = Int32.Parse(ObjReader["age"].ToString());
                    months = Int32.Parse(ObjReader["months"].ToString());
                    specimen = Record.StringToArray(ObjReader["specimen"].ToString());
                    refNoSet.Add(new Record(reportNo, ward, bht, "", name, age, months, gender, specimen, "", "", "", date, "", severity, ""));
            }
            if (refNoSet.Count > 0)
            {
                int tmp = MainWindow.topid;
                string temp = MainWindow.topdate;
                string StrCmd2;
                if(option==1)
                    StrCmd2= "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet.Last().Reference_No + "'";
                else
                    StrCmd2 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet[(MainWindow.listsize-1)].Reference_No + "'";
                string StrCmd3 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet.First().Reference_No + "'";
                OleDbCommand Cmd2 = new OleDbCommand(StrCmd2, MyConn);
                OleDbDataReader ObjReader2 = Cmd2.ExecuteReader();
                if (ObjReader2 == null)
                {
                    MainWindow.hasmore = false;
                    Console.WriteLine();////////////////////////////////////error in connecting to the database
                    return null;
                }
                else
                {
                    ObjReader2.Read();
                    MainWindow.bottomdate = ObjReader2["testdate"].ToString();
                    MainWindow.bottomid = Int32.Parse(ObjReader2["dateid"].ToString());
                    OleDbCommand Cmd3 = new OleDbCommand(StrCmd3, MyConn);
                    OleDbDataReader ObjReader3 = Cmd3.ExecuteReader();
                    if (ObjReader3 == null)
                    {
                        MainWindow.hasmore = false;
                        Console.WriteLine();////////////////////////////////////error in connecting to the database
                        return null;
                    }
                    else
                    {
                        ObjReader3.Read();
                        MainWindow.topdate = ObjReader3["testdate"].ToString();
                        MainWindow.topid = Int32.Parse(ObjReader3["dateid"].ToString());
                    }
                    if (refNoSet.Count == (MainWindow.listsize+1))
                    {
                        MainWindow.hasmore = true;
                        refNoSet.Remove(refNoSet.Last());
                    }
                    else
                        MainWindow.hasmore = false;
                    if (option == 0)
                    {
                        MainWindow.topdate = MainWindow.bottomdate;
                        MainWindow.bottomdate = temp;
                        MainWindow.topid = MainWindow.bottomid;
                        MainWindow.bottomid = tmp;
                        refNoSet.Reverse();                        
                    }
                    return refNoSet;
                }
            }
            else
            {
                MainWindow.hasmore = false;
                return null;
            }
        }

        public List<Record> getRecordList(int option)
        {
            List<Record> refNoSet = new List<Record>();
            string StrCmd;
            switch (option)
            {
                case (0):
                    {
                        StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, months, specimen,testDate,severity FROM Table1 where testdate >= '" + MainWindow.topdate + "' AND iif(testdate = '" + MainWindow.topdate + "' ,dateid > " + MainWindow.topid + ",1) AND" + MainWindow.searchPhrase + " order by testdate asc,dateid ASC";
                        break;
                    }
                case (1):
                default:
                    {
                        if (MainWindow.bottomdate != "")
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, months, specimen,testDate,severity FROM Table1 where testdate <= '" + MainWindow.bottomdate + "' AND iif(testdate = '" + MainWindow.bottomdate + "' ,dateid <= " + MainWindow.bottomid + ",1) AND" + MainWindow.searchPhrase + " order by testdate desc,dateid DESC";
                        else
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, months, specimen,testDate,severity FROM Table1 WHERE " + MainWindow.searchPhrase + " order by testdate desc, dateid desc";
                        break;
                    }
            }
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
            String name, gender, date, reportNo, bht, ward, severity;
            String[] specimen;
            int age, months;

            while (ObjReader.Read())
            {
                name = ObjReader["patientName"].ToString();
                gender = ObjReader["gender"].ToString();
                ward = ObjReader["ward"].ToString();
                bht = ObjReader["bht"].ToString();
                date = ObjReader["testDate"].ToString();
                reportNo = ObjReader["reportNo"].ToString();
                severity = ObjReader["severity"].ToString();        //**
                age = Int32.Parse(ObjReader["age"].ToString());
                months = Int32.Parse(ObjReader["months"].ToString());
                specimen = Record.StringToArray(ObjReader["specimen"].ToString());
                refNoSet.Add(new Record(reportNo, ward, bht, "", name, age, months, gender, specimen, "", "", "", date, "", severity, ""));
            }
            if (refNoSet.Count > 0)
            {
                int tmp = MainWindow.topid;
                string temp = MainWindow.topdate;
                string StrCmd2;
                if (option == 1)
                    StrCmd2 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet.Last().Reference_No + "'";
                else
                    StrCmd2 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet[(MainWindow.listsize - 1)].Reference_No + "'";
                string StrCmd3 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet.First().Reference_No + "'";
                OleDbCommand Cmd2 = new OleDbCommand(StrCmd2, MyConn);
                OleDbDataReader ObjReader2 = Cmd2.ExecuteReader();
                if (ObjReader2 == null)
                {
                    MainWindow.hasmore = false;
                    Console.WriteLine();////////////////////////////////////error in connecting to the database
                    return null;
                }
                else
                {
                    ObjReader2.Read();
                    MainWindow.bottomdate = ObjReader2["testdate"].ToString();
                    MainWindow.bottomid = Int32.Parse(ObjReader2["dateid"].ToString());
                    OleDbCommand Cmd3 = new OleDbCommand(StrCmd3, MyConn);
                    OleDbDataReader ObjReader3 = Cmd3.ExecuteReader();
                    if (ObjReader3 == null)
                    {
                        MainWindow.hasmore = false;
                        Console.WriteLine();////////////////////////////////////error in connecting to the database
                        return null;
                    }
                    else
                    {
                        ObjReader3.Read();
                        MainWindow.topdate = ObjReader3["testdate"].ToString();
                        MainWindow.topid = Int32.Parse(ObjReader3["dateid"].ToString());
                    }
                    if (refNoSet.Count == (MainWindow.listsize + 1))
                    {
                        MainWindow.hasmore = true;
                        refNoSet.Remove(refNoSet.Last());
                    }
                    else
                        MainWindow.hasmore = false;
                    if (option == 0)
                    {
                        MainWindow.topdate = MainWindow.bottomdate;
                        MainWindow.bottomdate = temp;
                        MainWindow.topid = MainWindow.bottomid;
                        MainWindow.bottomid = tmp;
                        refNoSet.Reverse();
                    }
                    return refNoSet;
                }
            }
            else
            {
                MainWindow.hasmore = false;
                return null;
            }
        }


        /*
        public List<Record> getRecordListByPartOfName(String partOfValue,String column, int option)
        {
            List<Record> refNoSet = new List<Record>();
            string StrCmd;
            switch (option)
            {
                case (0):
                    {
                        StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1 where "+column+" LIKE '%"+partOfValue+"%' AND ID > " + MainWindow.topid + " order by ID ASC";
                        break;
                    }
                case (1):
                default:
                    {
                        if (MainWindow.bottomid != 0)
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1 where " + column + " LIKE '%" + partOfValue + "%' AND ID <= " + MainWindow.bottomid + " order by ID DESC";
                        else
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1 where " + column + " LIKE '%" + partOfValue + "%' order by ID DESC";
                        break;
                    }
            }
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
            String name, gender, date, reportNo, bht, ward, severity;
            String[] specimen;
            int age;

            while (ObjReader.Read())
            {
                name = ObjReader["patientName"].ToString();
                gender = ObjReader["gender"].ToString();
                ward = ObjReader["ward"].ToString();
                bht = ObjReader["bht"].ToString();
                date = ObjReader["testDate"].ToString();
                reportNo = ObjReader["reportNo"].ToString();
                severity = ObjReader["severity"].ToString();        //**
                age = Int32.Parse(ObjReader["age"].ToString());
                specimen = Record.StringToArray(ObjReader["specimen"].ToString());
                refNoSet.Add(new Record(reportNo, ward, bht, "", name, age, gender, specimen, "", "", "", date, "", severity, "",0));
            }
            if (refNoSet.Count > 0)
            {
                int temp = MainWindow.topid;
                string StrCmd2;
                if (option == 1)
                    StrCmd2 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet.Last().Reference_No + "'";
                else
                    StrCmd2 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet[(MainWindow.listsize - 1)].Reference_No + "'";
                string StrCmd3 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet.First().Reference_No + "'";
                OleDbCommand Cmd2 = new OleDbCommand(StrCmd2, MyConn);
                OleDbDataReader ObjReader2 = Cmd2.ExecuteReader();
                if (ObjReader2 == null)
                {
                    MainWindow.hasmore = false;
                    Console.WriteLine();////////////////////////////////////error in connecting to the database
                    return null;
                }
                else
                {
                    ObjReader2.Read();
                    MainWindow.bottomid = Int32.Parse(ObjReader2["ID"].ToString());
                    OleDbCommand Cmd3 = new OleDbCommand(StrCmd3, MyConn);
                    OleDbDataReader ObjReader3 = Cmd3.ExecuteReader();
                    if (ObjReader3 == null)
                    {
                        MainWindow.hasmore = false;
                        Console.WriteLine();////////////////////////////////////error in connecting to the database
                        return null;
                    }
                    else
                    {
                        ObjReader3.Read();
                        MainWindow.topid = Int32.Parse(ObjReader3["ID"].ToString());
                    }
                    if (refNoSet.Count == (MainWindow.listsize + 1))
                    {
                        MainWindow.hasmore = true;
                        refNoSet.Remove(refNoSet.Last());
                    }
                    else
                        MainWindow.hasmore = false;
                    if (option == 0)
                    {
                        MainWindow.topid = MainWindow.bottomid;
                        MainWindow.bottomid = temp;
                        refNoSet.Reverse();
                    }
                    return refNoSet;
                }
            }
            else
            {
                MainWindow.hasmore = false;
                return null;
            }
        }

        public List<Record> getReportbyfullVariable(String value, String column,int option)
        {
            List<Record> refNoSet = new List<Record>();
            string StrCmd;
            switch (option)
            {
                case (0):
                    {
                        StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1 WHERE " + column + " = '" + value + "' AND ID > " + MainWindow.topid + " order by ID ASC";
                        break;
                    }
                case (1):
                default:
                    {
                        if (MainWindow.bottomid != 0)
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1 WHERE " + column + " = '" + value + "' AND ID <= " + MainWindow.bottomid + " order by ID DESC";
                        else
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1 WHERE " + column + " = '" + value + "' order by ID DESC";
                        break;
                    }
            }
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);                                    //**
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
            String name, gender, date, reportNo, bht, ward, severity;
            String[] specimen;
            int age,months;

            while (ObjReader.Read())
            {
                    name = ObjReader["patientName"].ToString();
                    gender = ObjReader["gender"].ToString();
                    ward = ObjReader["ward"].ToString();
                    bht = ObjReader["bht"].ToString();
                    date = ObjReader["testDate"].ToString();
                    reportNo = ObjReader["reportNo"].ToString();
                    severity = ObjReader["severity"].ToString();        //**
                    age = Int32.Parse(ObjReader["age"].ToString());
                    months = Int32.Parse(ObjReader["months"].ToString());
                    specimen = Record.StringToArray(ObjReader["specimen"].ToString());
                    refNoSet.Add(new Record(reportNo, ward, bht, "", name, age, gender, specimen, "", "", "", date,"",severity,"",0));                
            }
            if (refNoSet.Count > 0)
            {
                int temp = MainWindow.topid;
                string StrCmd2;
                if (option == 1)
                    StrCmd2 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet.Last().Reference_No + "'";
                else
                    StrCmd2 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet[(MainWindow.listsize - 1)].Reference_No + "'";
                string StrCmd3 = "SELECT * FROM Table1 WHERE reportNo = '" + refNoSet.First().Reference_No + "'";
                OleDbCommand Cmd2 = new OleDbCommand(StrCmd2, MyConn);
                OleDbDataReader ObjReader2 = Cmd2.ExecuteReader();
                if (ObjReader2 == null)
                {
                    MainWindow.hasmore = false;
                    Console.WriteLine();////////////////////////////////////error in connecting to the database
                    return null;
                }
                else
                {
                    ObjReader2.Read();
                    MainWindow.bottomid = Int32.Parse(ObjReader2["ID"].ToString());
                    OleDbCommand Cmd3 = new OleDbCommand(StrCmd3, MyConn);
                    OleDbDataReader ObjReader3 = Cmd3.ExecuteReader();
                    if (ObjReader3 == null)
                    {
                        MainWindow.hasmore = false;
                        Console.WriteLine();////////////////////////////////////error in connecting to the database
                        return null;
                    }
                    else
                    {
                        ObjReader3.Read();
                        MainWindow.topid = Int32.Parse(ObjReader3["ID"].ToString());
                    }
                    if (refNoSet.Count == (MainWindow.listsize + 1))
                    {
                        MainWindow.hasmore = true;
                        refNoSet.Remove(refNoSet.Last());
                    }
                    else
                        MainWindow.hasmore = false;
                    if (option == 0)
                    {
                        MainWindow.topid = MainWindow.bottomid;
                        MainWindow.bottomid = temp;
                        refNoSet.Reverse();
                    }
                    return refNoSet;
                }
            }
            else
            {
                MainWindow.hasmore = false;
                return null;
            }
        }
        */
        public void getTheRest(Record record)
        {
            string StrCmd = "SELECT  title, macroscopy, microscopy, conclusion, requestDate, clinicalDetails FROM Table1  WHERE reportNo = '" + record.Reference_No + "'";
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);                            //**
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
            ObjReader.Read();
            record.title = ObjReader["title"].ToString();
            record.macroscopy = ObjReader["macroscopy"].ToString();
            record.microscopy = ObjReader["microscopy"].ToString();
            record.conclusion = ObjReader["conclusion"].ToString();
            record.requestDate = ObjReader["requestDate"].ToString();   //**
            record.clinicalDetails = ObjReader["clinicalDetails"].ToString();//**
            record.complete = true;
        }

        public void store(Record record)
        {
            int dateid = 0;
            if (hasEntry(record.TestDate, "testdate"))
            {
                string StrCmd = "SELECT top 1 dateid FROM Table1 WHERE testdate = '" + record.TestDate + "' order by testdate desc";
                OleDbCommand Cmd1 = new OleDbCommand(StrCmd, MyConn);                            //**
                OleDbDataReader ObjReader = Cmd1.ExecuteReader();
                ObjReader.Read();
                dateid = Int32.Parse(ObjReader["dateid"].ToString()) + 1;
            }
            OleDbCommand Cmd = new OleDbCommand("INSERT INTO Table1 ( reportNo, ward,bht,title, patientName, age, gender, specimen,macroscopy, microscopy, conclusion,testDate,requestDate, severity, clinicalDetails, dateid, months ) VALUES ('" + record.Reference_No + "'," + "'" + record.Ward + "'," + "'" + record.BHT + "'," + "'" + record.title + "'," + "'" + record.Name + "',"
                + "'" + record.years + "'," + "'" + record.Gender + "'," + "'" + Record.ArrayToString(record.specimenArray) + "'," + "'" + record.macroscopy + "'," + "'" + record.microscopy + "'," + "'" + record.conclusion + "'," + "'" + record.TestDate + "'," + "'" + record.requestDate + "'," + "'" + record.severity + "'," + "'" + record.clinicalDetails + "'," + "'" + dateid + "'," + "'" + record.months + "')", MyConn); ;
            //OleDbCommand Cmd = new OleDbCommand("INSERT INTO Table1 ( name) VALUES ('" +record.microscopy + "')", MyConn); 

            Cmd.ExecuteNonQuery();
        }

        public bool hasEntry(String value,String column)
        {
            OleDbCommand cmdCheck = new OleDbCommand("SELECT COUNT(*) FROM Table1 WHERE "+column+" = '" + value + "'", MyConn);
            if (Convert.ToInt32(cmdCheck.ExecuteScalar()) == 0)
            {
                cmdCheck.ExecuteNonQuery();
                return false;
            }
            cmdCheck.ExecuteNonQuery();
            return true;
        }
        /*
        public bool haspartEntry(String value, String column)
        {            
            String StrCmd = "SELECT reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1";
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);                                    //**
            OleDbDataReader ObjReader = Cmd.ExecuteReader();

            while (ObjReader.Read())
            {
                if (ObjReader[column].ToString().ToLower().Contains(value.ToLower()))
                    return true;
            }
            return false;
        }*/

        public bool hasanyEntry()
        {
            OleDbCommand cmdCheck = new OleDbCommand("SELECT COUNT(*) FROM Table1", MyConn);
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

       public String[] getTemplates(String column,String speci)
        {
            List<String> templ = new List<String>();
            string StrCmd = "SELECT top 10 " + column + " FROM Table1 where specimen LIKE '%" + speci + "%' AND NOT (" + column + " = '') group by " + column + " order by max(ID) desc";
            OleDbCommand Cmd = new OleDbCommand(StrCmd, MyConn);
            OleDbDataReader ObjReader = Cmd.ExecuteReader();
            while (ObjReader.Read())
            {
                templ.Add(ObjReader[column].ToString());
            }
            if (templ.Count > 0)
                return templ.ToArray();
            else
                return null;
        }

       public int count()
       {
           int ret;
           OleDbCommand cmdCheck;
           if(MainWindow.searchPhrase=="")
                cmdCheck = new OleDbCommand("SELECT COUNT(*) FROM Table1", MyConn);
           else
               cmdCheck = new OleDbCommand("SELECT COUNT(*) FROM Table1 WHERE "+MainWindow.searchPhrase, MyConn);

           ret = Convert.ToInt32(cmdCheck.ExecuteScalar());
           cmdCheck.ExecuteNonQuery();
           return ret;
       }
    }
    class UniqueListItemObject
    {
        private string _text;
        public string Text { get { return _text; } set { _text = value; } }

        public UniqueListItemObject(string input)
        {
            Text = input;
        }
        public UniqueListItemObject()
        {
            Text = string.Empty;
        }

        public override string ToString()
        {
            return Text;
        }
    }
}