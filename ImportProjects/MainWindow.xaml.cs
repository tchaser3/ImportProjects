using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using NewEventLogDLL;
using NewEmployeeDLL;
using ProjectMatrixDLL;
using ProjectsDLL;
using DepartmentDLL;
using DataValidationDLL;
using System.Runtime.CompilerServices;

namespace ImportProjects
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        ProjectMatrixClass TheProjectMatrixClass = new ProjectMatrixClass();
        ProjectClass TheProjectClass = new ProjectClass();
        DepartmentClass TheDepartmentClass = new DepartmentClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();

        //setting up the data
        FindProjectByAssignedProjectIDDataSet TheFindProjectByAssignedProjectIDDataSet = new FindProjectByAssignedProjectIDDataSet();
        FindProjectMatrixByAssignedProjectIDDataSet TheFindProjectMatrixByAssignedProjectIDDataSet = new FindProjectMatrixByAssignedProjectIDDataSet();
        FindProjectMatrixByCustomerProjectIDDataSet TheFindProjectMatrixByCustomerProjectIDDataSet = new FindProjectMatrixByCustomerProjectIDDataSet();
        FindWarehousesDataSet TheFindWarehousesDataSet = new FindWarehousesDataSet();
        FindSortedDepartmentDataSet TheFindSortedDepartmentDataSet = new FindSortedDepartmentDataSet();
        FindEmployeeByEmployeeIDDataSet TheFindEmployeeByEmployeeIDDataSet = new FindEmployeeByEmployeeIDDataSet();
        ImportProjectsDataSet TheImportProjectsDataSet = new ImportProjectsDataSet();

        ChangeProjectIDOnEmployeeProjectAssignmentEntryTableAdapters.QueriesTableAdapter aChangeProjectIDOnEmployeeProjectAssignmentTableAdapter;
        ChangeProjectIDEmployeeCrewAssignmentEntryTableAdapters.QueriesTableAdapter aChangeProjectIDOnEmployeeCrewAssignmentTableAdapter;
        ChangeProjectIDOnProjectTaskEntryTableAdapters.QueriesTableAdapter aChangeProjectIDOnProjectTaskTableAdapter;


        //setting up global variables
        int gintOfficeID;
        int gintDepartmentID;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void expCloseProgram_Expanded(object sender, RoutedEventArgs e)
        {
            expCloseProgram.IsExpanded = false;
            TheMessagesClass.CloseTheProgram();
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ResetControls();
        }
        private void ResetControls()
        {
            //setting local variables
            int intCounter;
            int intNumberOfRecords;

            try
            {
                TheImportProjectsDataSet.importprojects.Rows.Clear();
                txtEmployeeID.Text = "";
                cboSelectDepartment.Items.Clear();
                cboSelectDepartment.Items.Add("Select Department");

                TheFindSortedDepartmentDataSet = TheDepartmentClass.FindSortedDepartment();

                intNumberOfRecords = TheFindSortedDepartmentDataSet.FindSortedDepartment.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    cboSelectDepartment.Items.Add(TheFindSortedDepartmentDataSet.FindSortedDepartment[intCounter].Department);
                }

                cboSelectDepartment.SelectedIndex = 0;

                cboSelectOffice.Items.Clear();
                cboSelectOffice.Items.Add("Select Office");

                TheFindWarehousesDataSet = TheEmployeeClass.FindWarehouses();

                intNumberOfRecords = TheFindWarehousesDataSet.FindWarehouses.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    cboSelectOffice.Items.Add(TheFindWarehousesDataSet.FindWarehouses[intCounter].FirstName);
                }

                cboSelectOffice.SelectedIndex = 0;
                dgrProjects.ItemsSource = TheImportProjectsDataSet.importprojects;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Projects // Main Window // Reset Controls " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void expImportExcel_Expanded(object sender, RoutedEventArgs e)
        {
            //setting up local variables
            string strValueForValidation = "";
            int intEmployeeID = 0;
            bool blnFatalError = false;
            bool blnThereIsAProblem = false;
            string strErrorMessage = "";
            int intRecordsReturned;
            string strAssignedProjectID = "";
            string strCustomerProjectID = "";
            int intProjectID = 0;
            DateTime datTransactionDate = DateTime.Now;
            int intSecondProjectID = 0;
            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            bool blnDoNotImport;

            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            try
            {
                expImportExcel.IsExpanded = false;

                if (cboSelectDepartment.SelectedIndex < 1)
                {
                    blnFatalError = true;
                    strErrorMessage += "The Department Was Not Selected\n";
                }
                if (cboSelectOffice.SelectedIndex < 1)
                {
                    blnFatalError = true;
                    strErrorMessage += "The Office Was Not Selected\n";
                }
                strValueForValidation = txtEmployeeID.Text;
                blnThereIsAProblem = TheDataValidationClass.VerifyIntegerData(strValueForValidation);
                if (blnThereIsAProblem == true)
                {
                    blnFatalError = true;
                    strErrorMessage += "The Employee ID Was Not An Integer\n";
                }
                else
                {
                    intEmployeeID = Convert.ToInt32(strValueForValidation);

                    TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intEmployeeID);

                    intRecordsReturned = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID.Rows.Count;

                    if (intRecordsReturned == 0)
                    {
                        blnFatalError = true;
                        strErrorMessage += "Employee Was Not Found\n";
                    }
                }
                if (blnFatalError == true)
                {
                    TheMessagesClass.ErrorMessage(strErrorMessage);
                    return;
                }

                TheImportProjectsDataSet.importprojects.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 1; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strCustomerProjectID = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2).ToUpper();
                    strAssignedProjectID = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2).ToUpper();
                    intProjectID = 0;
                    intSecondProjectID = 0;
                    blnDoNotImport = false;

                    TheFindProjectByAssignedProjectIDDataSet = TheProjectClass.FindProjectByAssignedProjectID(strCustomerProjectID);

                    if(TheFindProjectByAssignedProjectIDDataSet.FindProjectByAssignedProjectID.Rows.Count > 0)
                    {
                        intProjectID = TheFindProjectByAssignedProjectIDDataSet.FindProjectByAssignedProjectID[0].ProjectID;
                    }

                    TheFindProjectByAssignedProjectIDDataSet = TheProjectClass.FindProjectByAssignedProjectID(strAssignedProjectID);

                    if (TheFindProjectByAssignedProjectIDDataSet.FindProjectByAssignedProjectID.Rows.Count > 0)
                    {
                        intSecondProjectID = TheFindProjectByAssignedProjectIDDataSet.FindProjectByAssignedProjectID[0].ProjectID;
                    }

                    TheFindProjectMatrixByCustomerProjectIDDataSet = TheProjectMatrixClass.FindProjectMatrixByCustomerProjectID(strCustomerProjectID);

                    if(TheFindProjectMatrixByCustomerProjectIDDataSet.FindProjectMatrixByCustomerProjectID.Rows.Count > 0)
                    {
                        if(TheFindProjectMatrixByCustomerProjectIDDataSet.FindProjectMatrixByCustomerProjectID[0].AssignedProjectID == strAssignedProjectID)
                        {
                            blnDoNotImport = true;
                        }
                    }

                    if(intProjectID == 0)
                    {
                        if(intSecondProjectID != 0)
                        {
                            intProjectID = intSecondProjectID;
                        }
                    }

                    if((intProjectID == 0) && intSecondProjectID == 0)
                    {
                        blnFatalError = TheProjectClass.InsertProject(strCustomerProjectID, strCustomerProjectID);

                        if (blnFatalError == true)
                            throw new Exception();

                        TheFindProjectByAssignedProjectIDDataSet = TheProjectClass.FindProjectByAssignedProjectID(strCustomerProjectID);

                        intProjectID = TheFindProjectByAssignedProjectIDDataSet.FindProjectByAssignedProjectID[0].ProjectID;
                    }

                    ImportProjectsDataSet.importprojectsRow NewProjectRow = TheImportProjectsDataSet.importprojects.NewimportprojectsRow();

                    NewProjectRow.AssignedProjectID = strAssignedProjectID;
                    NewProjectRow.CustomerProjectID = strCustomerProjectID;
                    NewProjectRow.DepartmentID = gintDepartmentID;
                    NewProjectRow.EmployeeID = intEmployeeID;
                    NewProjectRow.ProjectID = intProjectID;
                    NewProjectRow.SecondProjectID = intSecondProjectID;
                    NewProjectRow.TransactionDate = DateTime.Now;
                    NewProjectRow.WarehouseID = gintOfficeID;
                    NewProjectRow.DoNotImport = blnDoNotImport;

                    TheImportProjectsDataSet.importprojects.Rows.Add(NewProjectRow);
                }

                PleaseWait.Close();

                dgrProjects.ItemsSource = TheImportProjectsDataSet.importprojects;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Projects // Main Window // Import Excel Expander " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void cboSelectDepartment_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;

            intSelectedIndex = cboSelectDepartment.SelectedIndex - 1;

            if(intSelectedIndex > -1)
            {
                gintDepartmentID = TheFindSortedDepartmentDataSet.FindSortedDepartment[intSelectedIndex].DepartmentID;
            }
        }

        private void cboSelectOffice_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;

            intSelectedIndex = cboSelectOffice.SelectedIndex - 1;

            if(intSelectedIndex > -1)
            {
                gintOfficeID = TheFindWarehousesDataSet.FindWarehouses[intSelectedIndex].EmployeeID;
            }
        }

        private void expProcess_Expanded(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            int intOldProjectID;
            int intNewProjectID;
            string strAssignedProjectID;
            string strCustomerProjectID;
            DateTime datTransactionDate;
            int intEmployeeID;
            int intDepartmentID;
            int intWarehouseID;
            bool blnFatalError = false;

            try
            {
                expProcess.IsExpanded = false;

                intNumberOfRecords = TheImportProjectsDataSet.importprojects.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    if(TheImportProjectsDataSet.importprojects[intCounter].DoNotImport == false)
                    {
                        intNewProjectID = TheImportProjectsDataSet.importprojects[intCounter].ProjectID;
                        intOldProjectID = TheImportProjectsDataSet.importprojects[intCounter].SecondProjectID;
                        strAssignedProjectID = TheImportProjectsDataSet.importprojects[intCounter].AssignedProjectID;
                        strCustomerProjectID = TheImportProjectsDataSet.importprojects[intCounter].CustomerProjectID;
                        datTransactionDate = DateTime.Now;
                        intEmployeeID = TheImportProjectsDataSet.importprojects[intCounter].EmployeeID;
                        intDepartmentID = TheImportProjectsDataSet.importprojects[intCounter].DepartmentID;
                        intWarehouseID = TheImportProjectsDataSet.importprojects[intCounter].WarehouseID;                        

                        blnFatalError = TheProjectMatrixClass.InsertProjectMatrix(intNewProjectID, strAssignedProjectID, strCustomerProjectID, datTransactionDate, intEmployeeID, intWarehouseID, intDepartmentID);

                        if (blnFatalError == true)
                            throw new Exception();                        

                        if(intNewProjectID != 0)
                        {
                            if(intOldProjectID != 0)
                            {
                                if (intNewProjectID != intOldProjectID)
                                {
                                    blnFatalError = ResetProjectIDEmployeeCrewAssignment(intOldProjectID, intNewProjectID);

                                    if (blnFatalError == true)
                                        throw new Exception();

                                    blnFatalError = ResetProjectIDEmployeeProjectAssignment(intOldProjectID, intNewProjectID);

                                    if (blnFatalError == true)
                                        throw new Exception();

                                    blnFatalError = ResetProjectIDProjectTask(intOldProjectID, intNewProjectID);

                                    if (blnFatalError == true)
                                        throw new Exception();

                                    blnFatalError = TheProjectClass.RemoveProjectEntry(intOldProjectID);

                                    if (blnFatalError == true)
                                        throw new Exception();
                                }
                            }
                            
                        }
                        
                    }
                }

                TheMessagesClass.InformationMessage("Import Complete");

                ResetControls();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Projects // Main Window // Process Expander " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

        }
        private bool ResetProjectIDProjectTask(int intOldProjectID, int intNewProjectID)
        {
            bool blnFatalError = false;

            try
            {
                aChangeProjectIDOnProjectTaskTableAdapter = new ChangeProjectIDOnProjectTaskEntryTableAdapters.QueriesTableAdapter();
                aChangeProjectIDOnProjectTaskTableAdapter.ChangeProjectIDOnProjectTask(intOldProjectID, intNewProjectID);
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Projects // Main Window // Reset Project ID Project Task " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }
        private bool ResetProjectIDEmployeeCrewAssignment(int intOldProjectID, int intNewProjectID)
        {
            bool blnFatalError = false;

            try
            {
                aChangeProjectIDOnEmployeeCrewAssignmentTableAdapter = new ChangeProjectIDEmployeeCrewAssignmentEntryTableAdapters.QueriesTableAdapter();
                aChangeProjectIDOnEmployeeCrewAssignmentTableAdapter.ChangeProjectIDEmployeeCrewAssignment(intOldProjectID, intNewProjectID);
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Projects // Main Window // Reset Project ID Employee Crew Assignment " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }
        private bool ResetProjectIDEmployeeProjectAssignment(int intOldProjectID, int intNewProjectID)
        {
            bool blnFatalError = false;

            try
            {
                aChangeProjectIDOnEmployeeProjectAssignmentTableAdapter = new ChangeProjectIDOnEmployeeProjectAssignmentEntryTableAdapters.QueriesTableAdapter();
                aChangeProjectIDOnEmployeeProjectAssignmentTableAdapter.ChangeProjectIDOnEmployeeProjectAssignment(intOldProjectID, intNewProjectID);
            }
            catch(Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Projects // Main Window // Reset Project ID Employee Project Assignment " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }
    }
}
