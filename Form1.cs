using EPDM.Interop.epdm;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Asn1.X509;
using System.Text.RegularExpressions;

/**

The rules for the spreadsheet are:

1) folder names in the project must match a sheet name in the spread sheet.

2) columns in the sheet must be titled one of these for the program to update from PDM:

                "Document #";
                "Rev"
                "# Sht"
                "Title"
                "DESIGNER"
                "REVIEWER"
                "APPROVER"
                "Inspection Notes"
                "Notes"
                "Status"

3) To convert a status the PDM values must match one of these, otherwise it stays the same.  It matches the first value and converts to the second.

            "NULL", "CREATE"
            "CREATE", "CREATE"
            "CONCEPT MODEL COMPLETE"
            "REVIEW", "REVIEW"
            "REVISE", "REVIEW"
            "APPROVE", "APPROVE"
            "RE-REVISE", "APPROVE"
            "PENDING EXWC REVIEW", "SMT EXWC"
            "RELEASED", "RELEASED"
            "CREATE REV", "CREATE"
            "CONCEPT MODEL COMPLETE REV", "CMC"
            "REVIEW REV", "REVIEW"
            "REVISE REV", "REVIEW"
            "APPROVE REV", "APPROVE"
            "RE-REVISE REV", "APPROVE"
            "PENDING EXWC REVIEW REV", "SMT EXWC"

4) the spreadsheet name must match the project folder name, apart from the DWG INDEX.xlsx part.

*/

namespace EinhornExportIndex
{
    public partial class Form1 : Form
    {
        const int TitleRow = 2; // this should be the row with all the column titles
        const int FirstFileRow = 3; // this should be the first row with file info
        const int LastFileRow = 94; // this should be the last row with posible file info

        //String Host = "Einhorn PDM";
        String Host = "TEST";
        String OutputFolder = "\\ENGINEERING DATA\\PDM INDEX OUTPUT\\";
        Dictionary<string, string> statusConversion = new Dictionary<string, string>();

        /**
         * this stores all the column numbers based on their titles in the title row
         */
        private struct ColumnNumbers
        {
            public void UpdateColumns(ISheet sheet)
            {
                fileNameColumn = FindColumn(sheet, "Document #");
                revisionColumn = FindColumn(sheet, "Rev");
                numberOfSheetsColumn = FindColumn(sheet, "# Sht");
                descriptionColumn = FindColumn(sheet, "Title");
                drawnByColumn = FindColumn(sheet, "DESIGNER");
                reviewerColumn = FindColumn(sheet, "REVIEWER");
                approverColumn = FindColumn(sheet, "APPROVER");
                inspectionNotesColumn = FindColumn(sheet, "Inspection Notes");
                notesColumn = FindColumn(sheet, "Notes");
                stateColumn = FindColumn(sheet, "Status");
            }
            public int fileNameColumn { get; set; }
            public int revisionColumn { get; set; }
            public int numberOfSheetsColumn { get; set; }
            public int descriptionColumn { get; set; }
            public int drawnByColumn { get; set; }
            public int reviewerColumn { get; set; }
            public int approverColumn { get; set; }
            public int inspectionNotesColumn { get; set; }
            public int notesColumn { get; set; }
            public int stateColumn { get; set; }
            private int FindColumn(ISheet sheet, String title)
            {
                IRow row = sheet.GetRow(TitleRow);
                foreach (ICell cell in row)
                {
                    if (cell.StringCellValue == title)
                    {
                        return cell.ColumnIndex;
                    }
                }

                return -1; // not found
            }

            public ICell? GetCell(IRow row, int columnIndex)
            {
                if (columnIndex != -1)
                {
                    return row.GetCell(columnIndex);
                }
                else
                {
                    return null;
                }
            }

        }

        IEdmVault7? vault;
        public Form1()
        {
            InitializeComponent();

            // Set the status column conversion table
            statusConversion.Add("NULL", "CREATE");
            statusConversion.Add("CREATE", "CREATE");
            statusConversion.Add("CONCEPT MODEL COMPLETE", "CMC");
            statusConversion.Add("REVIEW", "REVIEW");
            statusConversion.Add("REVISE", "REVIEW");
            statusConversion.Add("APPROVE", "APPROVE");
            statusConversion.Add("RE-REVISE", "APPROVE");
            statusConversion.Add("PENDING EXWC REVIEW", "SMT EXWC");
            statusConversion.Add("RELEASED", "RELEASED");
            statusConversion.Add("CREATE REV", "CREATE");
            statusConversion.Add("CONCEPT MODEL COMPLETE REV", "CMC");
            statusConversion.Add("REVIEW REV", "REVIEW");
            statusConversion.Add("REVISE REV", "REVIEW");
            statusConversion.Add("APPROVE REV", "APPROVE");
            statusConversion.Add("RE-REVISE REV", "APPROVE");
            statusConversion.Add("PENDING EXWC REVIEW REV", "SMT EXWC");
        }

        void EinhornExportIndex_Load(System.Object sender, System.EventArgs e)
        {
        }

        private void EinhornExportIndex_Click(System.Object sender, System.EventArgs e)
        {
            try
            {
                vault = new EdmVault5();

                //Log into selected vault as the current user
                vault.LoginAuto(Host, this.Handle.ToInt32());

                IEdmFolder5 Folder = vault.BrowseForFolder(0, "Select folder to traverse");

                if (Folder != null)
                {
                    String Path = GetWorkbookPath(Folder);

                    if (File.Exists(Path))
                    {
                        UpdateProject(Folder, Path);
                    }
                    else
                    {
                        IEdmPos5 FolderPos;
                        FolderPos = Folder.GetFirstSubFolderPosition();
                        while (!FolderPos.IsNull)
                        {
                            IEdmFolder5 SubFolder;
                            SubFolder = Folder.GetNextSubFolder(FolderPos);
                            Path = GetWorkbookPath(SubFolder);
                            if (File.Exists(Path))
                            {
                                UpdateProject(Folder, Path);
                            }
                        }
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                textBox1.AppendText("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                textBox1.AppendText(ex.Message);
            }

            textBox1.AppendText(Environment.NewLine + "Program Done");

        }

        private string GetWorkbookPath(IEdmFolder5 Folder)
        {
            return "C:\\" + Host + OutputFolder + Folder.Name + " DWG INDEX.xlsx";
        }

        private void UpdateProject(IEdmFolder5 Folder, string Path)
        {
            textBox1.AppendText("Workbook " + Path + Environment.NewLine);

            textBox1.AppendText("Locking " + Path + Environment.NewLine);

            //LockFile(Path);

            IWorkbook workBook;

            using (FileStream str = new FileStream(Path, FileMode.Open, FileAccess.Read))
            {
                textBox1.AppendText("Updating " + Path + Environment.NewLine);

                workBook = new XSSFWorkbook(str);

                TraverseProject(Folder, workBook);

                str.Close();

                textBox1.AppendText("Done Updating " + Path + Environment.NewLine);
            }

            using (FileStream str = new FileStream(Path, FileMode.Create, FileAccess.Write))
            {
                textBox1.AppendText("Writing " + Path + Environment.NewLine);
                workBook.Write(str);
                str.Close();

                textBox1.AppendText("Done Writing " + Path + Environment.NewLine);
            }

            textBox1.AppendText("Unlocking " + Path + Environment.NewLine);

            //UnlockFile(Path);

            textBox1.AppendText(Environment.NewLine);
            textBox1.AppendText("Done Processing for " + Path + Environment.NewLine);
        }

        /*
         * Traverse all the subfolders in the project folder and update the matching worksheet in the workbook
         */
        private void TraverseProject(IEdmFolder5 CurFolder, IWorkbook workBook)
        {
            try
            {
                if (CurFolder.Name.EndsWith("-DRAWINGS") || CurFolder.Name.EndsWith("-ANALYSIS"))
                {
                    UpdateWorksheet(workBook, CurFolder);
                }

                //Enumerate the sub-folders in the folder
                IEdmPos5 FolderPos;
                FolderPos = CurFolder.GetFirstSubFolderPosition();
                while (!FolderPos.IsNull)
                {
                    IEdmFolder5 SubFolder;
                    SubFolder = CurFolder.GetNextSubFolder(FolderPos);
                    TraverseProject(SubFolder, workBook);
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                textBox1.AppendText("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                textBox1.AppendText(ex.Message);
            }
        }

        /*
         * Update the worksheet whose name that matches the folder name
         * Ignore folders that do not have a matching worksheet
         */
        private void UpdateWorksheet(IWorkbook workBook, IEdmFolder5 Folder)
        {
            ISheet sheet = workBook.GetSheet(Folder.Name);

            if (sheet != null)
            {
                sheet.ForceFormulaRecalculation = true;

                ColumnNumbers columnNumbers = new ColumnNumbers();
                columnNumbers.UpdateColumns(sheet);

                textBox1.AppendText(Environment.NewLine);
                textBox1.AppendText("Updating worksheet " + Folder.Name + Environment.NewLine);

                IEdmPos5? FilePos = default;
                FilePos = Folder.GetFirstFilePosition();
                IEdmFile5? file = default;

                // Instantiate the regular expression object to match file names
                // File name should be 2 or 3 characters, a dash, two numbers, dash, then four characters or digits, then a dot.
                Regex r = new Regex("^[A-Za-z]{2,3}-\\d\\d-[a-zA-Z_0-9][a-zA-Z_0-9][a-zA-Z_0-9][a-zA-Z_0-9]..*$");
                List<IEdmFile5> newFiles = new List<IEdmFile5>();

                while (!FilePos.IsNull)
                {
                    file = Folder.GetNextFile(FilePos);

                    if (r.IsMatch(file.Name))
                    {
                        textBox1.AppendText("Reading File " + file.Name + Environment.NewLine);
                        file.Refresh();
                        if (!UpdateRows(sheet, columnNumbers, file))
                        {
                            newFiles.Add(file);
                        }
                    }
                }

                AddNewFiles(newFiles, workBook, sheet, columnNumbers);

            }
        }


        /**
         * add new files to this sheet
         */
        private void AddNewFiles(List<IEdmFile5> newFiles, IWorkbook workBook, ISheet sheet, ColumnNumbers columnNumbers)
        {
            textBox1.AppendText(Environment.NewLine);

            foreach (var file in newFiles)
            {
                textBox1.AppendText("new file " + file.Name + Environment.NewLine);

                int emptyRow = FindEmptyRow(sheet);

                if (emptyRow == -1)
                {
                    textBox1.AppendText(Environment.NewLine);
                    textBox1.AppendText("Please insert more empty rows manually" + Environment.NewLine);
                    textBox1.AppendText(Environment.NewLine);
                    break;
                }

                IRow row = sheet.GetRow(emptyRow);
                row.GetCell(1).SetCellValue(System.IO.Path.GetFileNameWithoutExtension(file.Name));
                UpdateRow(sheet, columnNumbers, row, file);
                row.Hidden = false;
            }

            XSSFFormulaEvaluator.EvaluateAllFormulaCells(workBook);
        }

        int FindEmptyRow(ISheet sheet)
        {
            for (int row = FirstFileRow; row <= LastFileRow; row++)
            {
                string fileName = sheet.GetRow(row).GetCell(1).StringCellValue;
                if (fileName == "")
                {
                    return row;
                }
            }

            return -1;
        }


        /**
         * update rows that match a drawing file 
         */
        private Boolean UpdateRows(ISheet sheet, ColumnNumbers columnNumbers, IEdmFile5 file)
        {
            foreach (IRow row in sheet)
            {
                try
                {
                    ICell? cell = columnNumbers.GetCell(row, columnNumbers.fileNameColumn);

                    if (cell != null && cell.CellType == CellType.String)
                    {
                        // If the file name matches this pattern with the first 11 chars, open it.  Like TX-27-D524.
                        string fileName = cell.StringCellValue + ".";

                        if (file.Name.Contains(fileName))
                        {
                            textBox1.AppendText("matched " + fileName + Environment.NewLine);

                            UpdateRow(sheet, columnNumbers, row, file);

                            return true;  // found it.
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    textBox1.AppendText("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
                }
                catch (Exception ex)
                {
                    textBox1.AppendText(ex.Message);
                }
            }

            return false; // Did not find
        }

        private void UpdateRow(ISheet sheet, ColumnNumbers columnNumbers, IRow row, IEdmFile5 file)
        {
            IEdmEnumeratorVariable8 EnumVarObj;
            EnumVarObj = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable();
            object VarObj;
            ICell? cell;

            if (EnumVarObj.GetVar("Revision", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.revisionColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Revision " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            if (EnumVarObj.GetVar("# Sheets", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.numberOfSheetsColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating # Sheets " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToInt64(VarObj));
                }
            }

            if (EnumVarObj.GetVar("Description", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.descriptionColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Description " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            if (EnumVarObj.GetVar("Drawn By", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.drawnByColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Drawn By " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            if (EnumVarObj.GetVar("Resp Eng", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.reviewerColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Resp Eng " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            if (EnumVarObj.GetVar("Approved By", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.approverColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Approved By " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            // get the state and convert it
            cell = columnNumbers.GetCell(row, columnNumbers.stateColumn);
            if (cell != null)
            {
                String status = file.CurrentState.Name;
                try
                {
                    status = statusConversion[status];
                    if (status == null)
                    {
                        status = statusConversion["NULL"];
                    }


                }
                catch (Exception)
                {
                    // no conversion available
                    status = statusConversion["NULL"];
                }

                textBox1.AppendText("Updating Status to: " + status + Environment.NewLine);
                cell.SetCellValue(status);

            }
        }

        private void UpdateColumns(ISheet sheet, IRow row, ColumnNumbers columnNumbers, IEdmFile5 file, string fileName)
        {
            textBox1.AppendText("matched " + fileName + Environment.NewLine);

            IEdmEnumeratorVariable8 EnumVarObj;
            EnumVarObj = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable();
            object VarObj;
            ICell? cell = null;

            if (EnumVarObj.GetVar("Revision", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.revisionColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Revision " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            if (EnumVarObj.GetVar("# Sheets", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.numberOfSheetsColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating # Sheets " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToInt64(VarObj));
                }
            }

            if (EnumVarObj.GetVar("Description", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.descriptionColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Description " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            if (EnumVarObj.GetVar("Drawn By", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.drawnByColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Drawn By " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            if (EnumVarObj.GetVar("Resp Eng", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.reviewerColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Resp Eng " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            if (EnumVarObj.GetVar("Approved By", "@", out VarObj) == true)
            {
                cell = columnNumbers.GetCell(row, columnNumbers.approverColumn);
                if (cell != null)
                {
                    textBox1.AppendText("Updating Approved By " + Environment.NewLine);

                    cell.SetCellValue(Convert.ToString(VarObj));
                }
            }

            // get the state and convert it
            cell = columnNumbers.GetCell(row, columnNumbers.stateColumn);
            if (cell != null)
            {
                String status = file.CurrentState.Name;
                try
                {
                    status = statusConversion[status];
                    if (status == null)
                    {
                        status = statusConversion["NULL"];
                    }

                    textBox1.AppendText("Updating Status " + Environment.NewLine);

                    cell.SetCellValue(status);
                }
                catch (Exception)
                {
                    textBox1.AppendText("No conversion for status " + status + Environment.NewLine);
                }

            }
        }
        private Boolean LockFile(String path)
        {
            Boolean altered = false;
            IEdmFile5? file;
            IEdmFolder5? folder = null;
            if (vault != null)
            {
                file = vault.GetFileFromPath(path, out folder);

                if (file != null && !file.IsLocked)
                {
                    altered = true;
                    file.LockFile(folder.ID, this.Handle.ToInt32());

                    textBox1.AppendText("Locked file " + path + Environment.NewLine);
                }

                return altered;
            }

            return altered;
        }

        private Boolean UnlockFile(String path)
        {
            Boolean altered = false;
            IEdmFile5? file = default;
            IEdmFolder5 folder;
            if (vault != null)
            {
                file = vault.GetFileFromPath(path, out folder);

                if (file != null && file.IsLocked)
                {
                    altered = true;
                    file.UnlockFile(folder.ID, "update");

                    textBox1.AppendText("Unlocked file " + path + Environment.NewLine);
                }
            }

            return altered;
        }

        /* old code
        private void AddVarColumn(XLWorkbook workbook, IEdmEnumeratorVariable8 EnumVarObj, String Name)
        {
            object VarObj = null;
            if (EnumVarObj.GetVar(Name, "@", out VarObj) == true)
            {
                workbook.AddNextCell(VarObj.ToString());
            }
            else
            {
                workbook.CurrentWorksheet.AddNextCell("");
            }
        }

        private void AddWorksheet(XLWorkbook workbook, IEdmFolder5 Folder)
        {

            try
            {
                workbook.RemoveWorksheet(Folder.Name);
            }
            catch
            {
            }

            workbook.AddWorksheet(Folder.Name);

            List<object> values = new List<object>() { "Name", "Revision", "# Sheets", "Description", "Resp Eng", "Drawn By", "State", "Version", "Notes", "Inspection Notes", "State Comments", "Checkin Comments" };
            workbook.CurrentWorksheet.AddCellRange(values, new Address(0, 0), new Address(11, 0));
            workbook.CurrentWorksheet.Cells["A1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["B1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["C1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["D1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["E1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["F1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["G1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["H1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["I1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["J1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["K1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.Cells["L1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.GoToNextRow();

            IEdmPos5 FilePos = default(IEdmPos5);
            FilePos = Folder.GetFirstFilePosition();
            IEdmFile5 file = default(IEdmFile5);

            // Instantiate the regular expression object to match file names
            Regex r = new Regex("^[A-Za-z][A-Za-z]-\\d\\d-[A-Za-z]\\d\\d\\d\\..*$");

            while (!FilePos.IsNull)
            {

                file = Folder.GetNextFile(FilePos);

                if (r.IsMatch(file.Name))
                {

                    textBox1.AppendText("Reading File " + file.Name + Environment.NewLine);

                    IEdmEnumeratorVariable8 EnumVarObj = default(IEdmEnumeratorVariable8);
                    EnumVarObj = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable();
                    object VarObj = null;

                    workbook.CurrentWorksheet.AddNextCell(Path.GetFileNameWithoutExtension(file.Name));

                    AddVarColumn(workbook, EnumVarObj, "Revision");
                    AddVarColumn(workbook, EnumVarObj, "# Sheets");
                    AddVarColumn(workbook, EnumVarObj, "Description");
                    AddVarColumn(workbook, EnumVarObj, "Resp Eng");
                    AddVarColumn(workbook, EnumVarObj, "Drawn By");

                    workbook.CurrentWorksheet.AddNextCell(file.CurrentState.Name);
                    workbook.CurrentWorksheet.AddNextCell(file.CurrentVersion.ToString());

                    AddVarColumn(workbook, EnumVarObj, "Notes");
                    AddVarColumn(workbook, EnumVarObj, "Inspection Notes");

                    IEdmHistory2 history = (IEdmHistory2)vault.CreateUtility(EdmUtility.EdmUtil_History);
                    history.AddFile(file.ID);
                    EdmHistoryItem[] ppoRethistory = null;

                    history.GetHistory(ref ppoRethistory, (int)EdmHistoryType.Edmhist_FileState);
                    workbook.CurrentWorksheet.AddNextCell(ppoRethistory[0].mbsComment);

                    history.GetHistory(ref ppoRethistory, (int)EdmHistoryType.Edmhist_FileVersion);
                    workbook.CurrentWorksheet.AddNextCell(ppoRethistory[0].mbsComment);

                    workbook.CurrentWorksheet.GoToNextRow();
                }
            }

            workbook.Save();

        }

        private void EinhornExportIndex_Click(System.Object sender, System.EventArgs e)
        {
            try
            {
                vault = new EdmVault5();

                //Log into selected vault as the current user
                vault.LoginAuto("Einhorn PDM", this.Handle.ToInt32());

                IEdmFolder5 Folder = vault.BrowseForFolder(0, "Select folder to traverse");
                String FileName = Folder.Name + " INDEX OUTPUT.xlsx";

                if (Folder != null)
                {
                    String Path = RootFolder + OutputFolder + Folder.Name + " INDEX OUTPUT.xlsx";

                    textBox1.AppendText("Workbook " + Path + Environment.NewLine);

                    Workbook workbook;

                    if (File.Exists(Path))
                    {
                        LockFile(Path);
                        workbook = getWorkbook(Path);
                    }
                    else
                    {
                        workbook = newWorkbook(Path);

                        IEdmFolder5 VaultFolder = default(IEdmFolder5);
                        VaultFolder = (IEdmFolder5)vault.RootFolder.GetSubFolder("ENGINEERING DATA").GetSubFolder("PDM INDEX OUTPUT"); ;
                        int ret = VaultFolder.AddFile(this.Handle.ToInt32(), "", FileName, 1);

                        textBox1.AppendText("Added to vault: " + Path +  Environment.NewLine);
                    }

                    TraverseFolder(Folder, workbook);
                    workbook.Save();

                    UnlockFile(Path);

                    textBox1.AppendText("Done" + Environment.NewLine);
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    */

    }

}






