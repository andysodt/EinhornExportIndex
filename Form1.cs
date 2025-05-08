using EPDM.Interop.epdm;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;
using System.Security.Policy;
using System.Text.RegularExpressions;


namespace EinhornExportIndex
{
    public partial class Form1 : Form
    {
        String RootFolder = "C:\\Einhorn PDM\\";
        String OutputFolder = "ENGINEERING DATA\\PDM INDEX OUTPUT\\";

        /**
         * this stores all the column numbers based on their titles in the title row
         */
        private struct ColumnNumbers
        {
            const int TitleRow = 2; // this should be the row with all the column titles
            public void UpdateColumns(ISheet sheet)
            {
                fileNameColumn = FindColumn(sheet, "Drawing #");
                revisionColumn = FindColumn(sheet, "Rev");
                numberOfSheetsColumn = FindColumn(sheet, "# Sht");
                descriptionColumn = FindColumn(sheet, "Title");
                drawnByColumn = FindColumn(sheet, "DESIGNER");
                reviewerColumn = FindColumn(sheet, "REVIEWER");
                inspectionNotesColumn = FindColumn(sheet, "Inspection Notes");
                notesColumn = FindColumn(sheet, "Notes");
            }
            public int fileNameColumn { get; set; }
            public int revisionColumn { get; set; }
            public int numberOfSheetsColumn { get; set; }
            public int descriptionColumn { get; set; }
            public int drawnByColumn { get; set; }
            public int reviewerColumn { get; set; }
            public int inspectionNotesColumn { get; set; }
            public int notesColumn { get; set; }
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
        }

        IEdmVault7 vault;
        public Form1()
        {
            InitializeComponent();
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
                vault.LoginAuto("Einhorn PDM", this.Handle.ToInt32());

                IEdmFolder5 Folder = vault.BrowseForFolder(0, "Select folder to traverse");

                if (Folder != null)
                {
                    String Path = RootFolder + OutputFolder + Folder.Name + " INDEX (EINENG).xlsx";
                    String NewPath = RootFolder + OutputFolder + Folder.Name + " INDEX (EINENG) NEW.xlsx";

                    textBox1.AppendText("Workbook " + Path + Environment.NewLine);

                    //LockFile(NewPath);

                    IWorkbook workBook = null;

                    using (FileStream str = new FileStream(Path, FileMode.Open, FileAccess.Read))
                    {
                        workBook = new XSSFWorkbook(str);

                        TraverseFolder(Folder, workBook);
                    }

                    using (FileStream str = new FileStream(NewPath, FileMode.OpenOrCreate, FileAccess.Write))
                    {
                        workBook.Write(str);
                    }

                    //UnlockFile(NewPath);

                    textBox1.AppendText("Done Processing" + Environment.NewLine);
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


        /*
         * Traverse all the subfolders in the chosen folder andupdate the matching worksheet in the workbook
         */
        private void TraverseFolder(IEdmFolder5 CurFolder, IWorkbook workBook)
        {
            try
            {
                if (CurFolder.Name.EndsWith("-DRAWINGS") || CurFolder.Name.EndsWith("-ANALYSIS"))
                {
                    UpdateWorksheet(workBook, CurFolder);
                }

                //Enumerate the sub-folders in the folder
                IEdmPos5 FolderPos = default(IEdmPos5);
                FolderPos = CurFolder.GetFirstSubFolderPosition();
                while (!FolderPos.IsNull)
                {
                    IEdmFolder5 SubFolder = default(IEdmFolder5);
                    SubFolder = CurFolder.GetNextSubFolder(FolderPos);
                    TraverseFolder(SubFolder, workBook);
                }

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                ColumnNumbers columnNumbers = new ColumnNumbers();
                columnNumbers.UpdateColumns(sheet);

                textBox1.AppendText("Updating worksheet " + Folder.Name + Environment.NewLine);

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
                        UpdateRow(sheet, columnNumbers, file);
                    }
                }
            }
        }

        /**
         * update a row that matches a drawing file 
         */
        private void UpdateRow(ISheet sheet, ColumnNumbers columnNumbers, IEdmFile5 file)
        {
            foreach (IRow row in sheet)
            {
                ICell cell = row.GetCell(columnNumbers.fileNameColumn);

                if (cell != null) 
                {
                    string fileName = cell.StringCellValue + ".SLDDRW";

                    if (fileName == file.Name)
                    {
                        textBox1.AppendText("matched " + fileName + Environment.NewLine);

                        IEdmEnumeratorVariable8 EnumVarObj = default(IEdmEnumeratorVariable8);
                        EnumVarObj = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable();
                        object VarObj = null;

                        if (EnumVarObj.GetVar("Revision", "@", out VarObj) == true)
                        {
                            cell = row.GetCell(columnNumbers.revisionColumn);
                            cell.SetCellValue(Convert.ToString(VarObj));
                        }

                        if (EnumVarObj.GetVar("# Sheets", "@", out VarObj) == true)
                        {
                            cell = row.GetCell(columnNumbers.numberOfSheetsColumn);
                            cell.SetCellValue(Convert.ToInt64(VarObj));
                        }

                        if (EnumVarObj.GetVar("Description", "@", out VarObj) == true)
                        {
                            cell = row.GetCell(columnNumbers.descriptionColumn);
                            cell.SetCellValue(Convert.ToString(VarObj));
                        }

                        if (EnumVarObj.GetVar("Drawn By", "@", out VarObj) == true)
                        {
                            cell = row.GetCell(columnNumbers.drawnByColumn);
                            cell.SetCellValue(Convert.ToString(VarObj));
                        }

                        if (EnumVarObj.GetVar("Resp Eng", "@", out VarObj) == true)
                        {
                            cell = row.GetCell(columnNumbers.reviewerColumn);
                            cell.SetCellValue(Convert.ToString(VarObj));
                        }

                        if (EnumVarObj.GetVar("Inspection Notes", "@", out VarObj) == true)
                        {
                            cell = row.GetCell(columnNumbers.inspectionNotesColumn);
                            cell.SetCellValue(Convert.ToString(VarObj));
                        }

                        if (EnumVarObj.GetVar("Notes", "@", out VarObj) == true)
                        {
                            cell = row.GetCell(columnNumbers.notesColumn);
                            cell.SetCellValue(Convert.ToString(VarObj));
                        }

                        /* perhaps not needed
                        IEdmHistory2 history = (IEdmHistory2)vault.CreateUtility(EdmUtility.EdmUtil_History);
                        history.AddFile(file.ID);
                        EdmHistoryItem[] ppoRethistory = null;

                        history.GetHistory(ref ppoRethistory, (int)EdmHistoryType.Edmhist_FileState);
                        //sheet.GetRow(row).GetCell(11).SetCellValue(ppoRethistory[0].mbsComment);

                        history.GetHistory(ref ppoRethistory, (int)EdmHistoryType.Edmhist_FileVersion);
                        //sheet.GetRow(row).GetCell(12).SetCellValue(ppoRethistory[0].mbsComment);

                        //sheet.GetRow(row).GetCell(8).SetCellValue(file.CurrentVersion);
                        //sheet.GetRow(row).GetCell(11).SetCellValue(file.CurrentState.Name);
                        */
                    }
                }
            }
        }




        private void UpdateCell(ISheet sheet, IEdmFile5 file, ICell cell, String name, Boolean isString)
        {
            IEdmEnumeratorVariable8 EnumVarObj = default(IEdmEnumeratorVariable8);
            EnumVarObj = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable();
            object VarObj = null;

            if (EnumVarObj.GetVar(name, "@", out VarObj) == true)
            {
                if (isString)
                {
                    cell.SetCellValue(Convert.ToString(VarObj));
                }
                else
                {
                    cell.SetCellValue(Convert.ToInt64(VarObj));
                }

                //textBox1.AppendText("Row " + row + " Column " + column + " set to " + VarObj.ToString() + Environment.NewLine);
            }
        }

        private Boolean LockFile(String path)
        {
            Boolean altered = false;
            IEdmFile5 file = default(IEdmFile5);
            IEdmFolder5 folder = null;
            file = this.vault.GetFileFromPath(path, out folder);

            if (file != null && !file.IsLocked)
            {
                altered = true;
                file.LockFile(folder.ID, this.Handle.ToInt32());

                textBox1.AppendText("Locked file " + path + Environment.NewLine);
            }

            return altered;
        }

        private Boolean UnlockFile(String path)
        {
            Boolean altered = false;
            IEdmFile5 file = default(IEdmFile5);
            IEdmFolder5 folder = null;
            file = this.vault.GetFileFromPath(path, out folder);

            if (file != null && file.IsLocked)
            {
                altered = true;
                file.UnlockFile(folder.ID, "update");

                textBox1.AppendText("Unlocked file " + path + Environment.NewLine);
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






