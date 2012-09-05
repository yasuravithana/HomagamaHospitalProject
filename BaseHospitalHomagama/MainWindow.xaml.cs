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
        Database database = new Database("database.mdb");
        CultureInfo cultureInfo;//**
        TextInfo textInfo;//**
        int selectedList;           //0-all, 1-bht, 2-name, 3-specimen, 4-severity
        public static int topid;
        public static int bottomid;
        public static bool hasmore;
        public static int listsize=10;

        private DispatcherTimer timer1 = new DispatcherTimer();
        
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
            comboBoxSeverity.SelectedIndex = 0;                                          //**

            datePicker1.SelectedDate = DateTime.Today;
            datePicker2.SelectedDate = DateTime.Today;
            cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;//**
            textInfo = cultureInfo.TextInfo;//**

            spSg = new specimenSuggestions();

            timer1.Interval = new TimeSpan(0, 0, 4);
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

        private Boolean retrieveReport(String reportNo)
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
        }

        private void fillFields(Record record)
        {
            clear();
            textReportNo.Text = record.Reference_No;
            textPatientName.Text = record.Name;
            textWardNo.Text = record.Ward;
            textBhtNo.Text = record.BHT;
            comboBoxTitle.SelectedIndex = -1;
            comboBoxTitle.Text = record.title;
            textPatientName.Text = record.Name;
            textAge.Text = record.Age.ToString();
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
            if (textClinicalDetails.Text=="")                                   //**
                FindindReplace(wordApp, "<clinical>", replaceNewLines("-"));    //**
            else                                                              //**
                FindindReplace(wordApp, "<clinical>", replaceNewLines(textClinicalDetails.Text));//**
            FindindReplace(wordApp, "<macro>", replaceNewLines(textMacroscopy.Text));
            FindindReplace(wordApp, "<micro>", replaceNewLines(textMicroscopy.Text));
            FindindReplace(wordApp, "<con>", replaceNewLines(textConclusion.Text));
            FindindReplace(wordApp, "<date>", datePicker2.SelectedDate.Value.Date.Day + " / " + datePicker2.SelectedDate.Value.Date.Month + " / " + datePicker2.SelectedDate.Value.Date.Date.Year);
            FindindReplace(wordApp, "<reqdate>", datePicker1.SelectedDate.Value.Date.Day + " / " + datePicker1.SelectedDate.Value.Date.Month + " / " + datePicker1.SelectedDate.Value.Date.Date.Year);//**
            FindindReplace(wordApp, "<printdate>", DateTime.Today.Date.Day + " / " + DateTime.Today.Date.Month + " / " + DateTime.Today.Date.Year);//**

            aDoc.SaveAs(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing,
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
            if (grid3.Visibility == System.Windows.Visibility.Visible)
            {
                if ((Record)dataGrid1.SelectedItem == null)
                {
                    MessageBox.Show("Please select a report first.", "", MessageBoxButton.OK, MessageBoxImage.Stop);
                    return;
                }
            }
            labelError.Foreground = Brushes.CadetBlue;
            labelError.Content = "Saving report in the database....";
            labelError.Visibility = System.Windows.Visibility.Visible;
            labelError.UpdateLayout();
            if (grid3.Visibility == System.Windows.Visibility.Hidden)
            {
                if (!save())
                {
                    return;
                }
            }
            labelError.Foreground = Brushes.CadetBlue;
            labelError.Visibility = System.Windows.Visibility.Visible;
            labelError.Content = "Transferring report to the printer....";

            print();

            labelError.Content = "Report is saved and has been transferred to the printer....";

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

        private void buttonEdit_Click(object sender, RoutedEventArgs e)
        {
            writeEnable();
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

            database.connectToDatabase();
            if (!database.hasEntry(textReportNo.Text,"reportNo"))
            {                                                                   //**
                database.store(new Record(textReportNo.Text, textWardNo.Text, textBhtNo.Text, comboBoxTitle.Text, textPatientName.Text, Int32.Parse(textAge.Text), comboBoxGender.Text, textBoxListToStringArray(speciTexts), textMacroscopy.Text, textMicroscopy.Text, textConclusion.Text, dateToString(datePicker2.SelectedDate.Value), dateToString(datePicker1.SelectedDate.Value), comboBoxSeverity.Text, textClinicalDetails.Text));

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
                    database.store(new Record(textReportNo.Text, textWardNo.Text, textBhtNo.Text, comboBoxTitle.Text, textPatientName.Text, Int32.Parse(textAge.Text), comboBoxGender.Text, textBoxListToStringArray(speciTexts), textMacroscopy.Text, textMicroscopy.Text, textConclusion.Text, dateToString(datePicker2.SelectedDate.Value), dateToString(datePicker1.SelectedDate.Value), comboBoxSeverity.Text, textClinicalDetails.Text));

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
        }

        private void clear()
        {
            writeEnable();
            textReportNo.IsEnabled = true;
            buttonDelete.Visibility = System.Windows.Visibility.Hidden;
            if (!grid3.IsVisible)
                menu.IsEnabled = true;

            textReportNo.Text = "";
            textWardNo.Text = "";
            textBhtNo.Text = "";
            textPatientName.Text = "";
            textAge.Text = "";
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
                    menuItemGetReport.IsEnabled = false;
                }
                else
                {
                    MessageBox.Show("The reference number you entered did not match any report.\nPlease check the reference number and try again.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            else  // advanced search   
            {
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

                topid = bottomid = 0;
                hasmore = false;

                if (textPatientName.Text != "")
                {
                    database.connectToDatabase();
                    if (database.haspartEntry(textPatientName.Text, "patientName"))
                    {
                        selectedList = 2;
                        dataGrid1.DataContext = database.getRecordListByPartOfName(textPatientName.Text, "patientName",1);
                        database.closeConnection();
                    }
                    else
                    {
                        database.closeConnection();
                        methodForbuttonHome_Click();
                        MessageBox.Show("The Patient Name you entered did not match any report.\nPlease check and try again.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }

                else if (textBhtNo.Text != "")
                {
                    database.connectToDatabase();
                    if (database.hasEntry(textBhtNo.Text, "bht"))
                    {
                        selectedList = 1;
                        dataGrid1.DataContext = database.getReportbyfullVariable(textBhtNo.Text,"bht",1);
                        database.closeConnection();                       
                    }
                    else
                    {
                        database.closeConnection();
                        methodForbuttonHome_Click();
                        MessageBox.Show("The BHT number you entered did not match any report.\nPlease check and try again.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }

                else if (textSpecimen.Text != "")
                {
                    database.connectToDatabase();
                    if (database.haspartEntry(textSpecimen.Text, "specimen"))
                    {
                        selectedList = 3;
                        dataGrid1.DataContext = database.getRecordListByPartOfName(textSpecimen.Text, "specimen",1);
                        database.closeConnection();
                    }
                    else
                    {
                        database.closeConnection();
                        methodForbuttonHome_Click();
                        MessageBox.Show("The Specimen you entered did not match any report.\nPlease check and try again.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }

                else if (comboBoxSeverity.Text != "No significance")
                {
                    database.connectToDatabase();
                    if (database.hasEntry(comboBoxSeverity.Text, "severity"))
                    {
                        selectedList = 4;
                        dataGrid1.DataContext = database.getReportbyfullVariable(comboBoxSeverity.Text, "severity",1);
                        database.closeConnection();
                    }
                    else
                    {
                        database.closeConnection();
                        methodForbuttonHome_Click();
                        MessageBox.Show("The Severity you entered did not match any report.\nPlease check and try again.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }

                else
                {
                    database.connectToDatabase();
                    if (database.hasanyEntry())
                    {
                        selectedList = 0;
                        dataGrid1.DataContext = database.getAllRecordList(1);//1-next,0-back
                        database.closeConnection();                        
                    }
                    else
                    {
                        database.closeConnection();
                        methodForbuttonHome_Click();
                        MessageBox.Show("There are no reports stored.", "Report not found!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }

                if (hasmore)
                {
                    buttonBack.Visibility = System.Windows.Visibility.Visible;
                    buttonBack.IsEnabled = false;
                    buttonNext.Visibility = System.Windows.Visibility.Visible;
                    buttonNext.IsEnabled = true;
                }
                
                dataGrid1.SelectionChanged+=new SelectionChangedEventHandler(DataGrid_SelectionChanged);
            }
        }

        private void buttonNext_Click(object sender, RoutedEventArgs e)
        {
            database.connectToDatabase();
            if (selectedList == 0)
                dataGrid1.DataContext = database.getAllRecordList(1);
            else if (selectedList == 1)
                dataGrid1.DataContext = database.getReportbyfullVariable(textBhtNo.Text, "bht", 1);
            else if (selectedList == 2)
                dataGrid1.DataContext = database.getRecordListByPartOfName(textPatientName.Text, "patientName", 1);
            else if (selectedList == 3)
                dataGrid1.DataContext = database.getRecordListByPartOfName(textSpecimen.Text, "specimen", 1);
            else if (selectedList == 4)
                dataGrid1.DataContext = database.getReportbyfullVariable(comboBoxSeverity.Text, "severity", 1);
            database.closeConnection();
            if (!hasmore)
                buttonNext.IsEnabled = false;
            buttonBack.IsEnabled = true;            
        }

        private void buttonBack_Click(object sender, RoutedEventArgs e)
        {
            database.connectToDatabase();
            if (selectedList == 0)
                dataGrid1.DataContext = database.getAllRecordList(0);
            else if (selectedList == 1)
                dataGrid1.DataContext = database.getReportbyfullVariable(textBhtNo.Text, "bht", 0);
            else if (selectedList == 2)
                dataGrid1.DataContext = database.getRecordListByPartOfName(textPatientName.Text, "patientName", 0);
            else if (selectedList == 3)
                dataGrid1.DataContext = database.getRecordListByPartOfName(textSpecimen.Text, "specimen", 3);
            else if (selectedList == 4)
                dataGrid1.DataContext = database.getReportbyfullVariable(comboBoxSeverity.Text, "severity", 0);
            database.closeConnection();
            if (!hasmore)
                buttonBack.IsEnabled = false;
            buttonNext.IsEnabled = true;            
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
            dataGrid1.SelectionChanged -= DataGrid_SelectionChanged;
            buttonClear.Content = "Clear";
            buttonClear.Click -= buttonHome_Click;
            buttonClear.Click += new RoutedEventHandler(buttonClear_Click);
            dataGrid1.DataContext = null;
        }

        private void buttonEditInPreview_Click(object sender, RoutedEventArgs e)
        {
            textReportNo.IsEnabled = false;
            buttonDelete.Visibility = System.Windows.Visibility.Visible;
            if ((Record)dataGrid1.SelectedItem != null)
            {
                grid3.Visibility = System.Windows.Visibility.Hidden;
                dataGrid1.SelectionChanged -= DataGrid_SelectionChanged; 
                buttonClear.Content = "Clear"; 
                buttonClear.Click -= buttonHome_Click; 
                buttonClear.Click += new RoutedEventHandler(buttonClear_Click);
                writeEnable();
                buttonSave.Content = "Save";
                buttonSave.Click -= buttonEditInPreview_Click;
                buttonSave.Click += new RoutedEventHandler(buttonSave_Click);
                dataGrid1.DataContext = null;
                buttonNext.Visibility = System.Windows.Visibility.Hidden;
                buttonBack.Visibility = System.Windows.Visibility.Hidden;
            }
            else
            {
                MessageBox.Show("Please select a report first.", "", MessageBoxButton.OK, MessageBoxImage.Stop);
            }
        }


        private void writeDisable()
        {
            textReportNo.IsReadOnly = true; 
            textWardNo.IsReadOnly = true; textBhtNo.IsReadOnly = true;
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
            textReportNo.IsReadOnly = false; 
            textWardNo.IsReadOnly = false; textBhtNo.IsReadOnly = false;
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
        public int Age { set; get; }
        public String Specimens { set; get; }
        public String requestDate;                  //**
        public String severity { set; get; }        //**
        public String clinicalDetails;              //**

        public Record(String reportNo, String ward, String bht, String title, String name, int age, String gender, String[] specimen, String macroscopy, String microscopy,
                                                        String conclusion, String date, String requestDate, String severity, String clinicalDetails)
        {                                                                           //**

            this.Name = name;
            this.Age = age;
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

        public Record getReport(String refNo) //**
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
                int age;                                                                                            //**
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
                age = Int32.Parse(ObjReader["age"].ToString());
                Record p = new Record(reportNo, ward, bht, title, name, age, gender, Record.StringToArray(specimen), macroscopy, microscopy, conclusion, date, requestDate, severity, clinicalDetails);
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

        public List<Record> getAllRecordList(int option)
        {
            List<Record> refNoSet = new List<Record>();
            string StrCmd;
            switch(option)
            {
                case (0):
                    {
                        StrCmd = "SELECT top " + (MainWindow.listsize+1) + " reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1 where ID > " + MainWindow.topid + " order by ID ASC";
                        break;
                    }
                case(1):
                default:
                    {
                        if (MainWindow.bottomid != 0)
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1 where ID <= " + MainWindow.bottomid + " order by ID DESC";
                        else
                            StrCmd = "SELECT top " + (MainWindow.listsize + 1) + " reportNo, patientName,gender, ward, bht, age, specimen,testDate,severity FROM Table1 order by ID DESC";
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
                    refNoSet.Add(new Record(reportNo, ward, bht, "", name, age, gender, specimen, "", "", "", date, "", severity, ""));
            }
            if (refNoSet.Count > 0)
            {
                int temp = MainWindow.topid;
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
                    if (refNoSet.Count == (MainWindow.listsize+1))
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
                refNoSet.Add(new Record(reportNo, ward, bht, "", name, age, gender, specimen, "", "", "", date, "", severity, ""));
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
                    refNoSet.Add(new Record(reportNo, ward, bht, "", name, age, gender, specimen, "", "", "", date,"",severity,""));                
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
        {                                                                               //**
            OleDbCommand Cmd = new OleDbCommand("INSERT INTO Table1 ( reportNo, ward,bht,title, patientName, age, gender, specimen,macroscopy, microscopy, conclusion,testDate,requestDate, severity, clinicalDetails ) VALUES ('" + record.Reference_No + "'," + "'" + record.Ward + "'," + "'" + record.BHT + "'," + "'" + record.title + "'," + "'" + record.Name + "',"
                + "'" + record.Age + "'," + "'" + record.Gender + "'," + "'" + Record.ArrayToString(record.specimenArray) + "'," + "'" + record.macroscopy + "'," + "'" + record.microscopy + "'," + "'" + record.conclusion + "'," + "'" + record.TestDate + "'," + "'" + record.requestDate + "'," + "'" + record.severity + "'," + "'" + record.clinicalDetails + "')", MyConn); ;
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
        }

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

       /* public int count(int option)
        {
            String command;
            switch (option)
            {
                case 1:
                    {
                        command = "SELECT COUNT(*) FROM Table1";
                        break;
                    }
                default: 
                    {
                        return 0;
                    }                    
            }
            OleDbCommand cmdCheck = new OleDbCommand(command, MyConn);
            return Convert.ToInt32(cmdCheck.ExecuteScalar());
        }*/
    }
}