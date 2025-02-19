using EPDM.Interop.epdm;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace EinhornExportIndex
{
    public partial class Form1 : Form
    {
        String RootFolder = "C:\\Einhorn PDM\\";
        String OutputFolder = "ENGINEERING DATA\\PDM INDEX OUTPUT\\";

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
                    String Path = RootFolder + OutputFolder + Folder.Name + " INDEX.xlsx";

                    textBox1.AppendText("Workbook " + Path + Environment.NewLine);

                    XLWorkbook workbook;

                    if (File.Exists(Path))
                    {
                        //LockFile(Path);
                        workbook = getWorkbook(Path);
                        TraverseFolder(Folder, workbook);
                        workbook.SaveAs(Path);
                        textBox1.AppendText("Done Processing" + Environment.NewLine);
                        //UnlockFile(Path);
                    }
                    else
                    {
                        textBox1.AppendText("Template file does not exist: " + Path + Environment.NewLine);
                    }


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
        private void TraverseFolder(IEdmFolder5 CurFolder, XLWorkbook workbook)
        {
            try
            {
                if (CurFolder.Name.EndsWith("-DRAWINGS") || CurFolder.Name.EndsWith("-ANALYSIS"))
                {
                    UpdateWorksheet(workbook, CurFolder);
                }

                //Enumerate the sub-folders in the folder
                IEdmPos5 FolderPos = default(IEdmPos5);
                FolderPos = CurFolder.GetFirstSubFolderPosition();
                while (!FolderPos.IsNull)
                {
                    IEdmFolder5 SubFolder = default(IEdmFolder5);
                    SubFolder = CurFolder.GetNextSubFolder(FolderPos);
                    TraverseFolder(SubFolder, workbook);
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
        private void UpdateWorksheet(XLWorkbook workbook, IEdmFolder5 Folder)
        {
            IXLWorksheet sheet = null;
            if (workbook.TryGetWorksheet(Folder.Name, out sheet))
            {
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
                        UpdateRow(sheet, file);
                    }
                }
            }
        }

        private void UpdateRow(IXLWorksheet sheet, IEdmFile5 file)
        {
            Boolean done = false;
            for (int row = 1; !done && row < sheet.RowCount(); row++)
            {
                IXLCell cell = sheet.Cell(row, 1);
                string fileName = cell.CachedValue + ".SLDDRW";
                if (fileName == file.Name)
                {
                    textBox1.AppendText("Row " + Convert.ToString(row) + " matched " + fileName + Environment.NewLine);

                    IEdmEnumeratorVariable8 EnumVarObj = default(IEdmEnumeratorVariable8);
                    EnumVarObj = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable();
                    object VarObj = null;

                    UpdateColumn(sheet, file, row, 2, "Revision", true);
                    UpdateColumn(sheet, file, row, 3, "# Sheets", false);
                    UpdateColumn(sheet, file, row, 4, "Description", true);
                    UpdateColumn(sheet, file, row, 5, "Resp Eng", true);
                    UpdateColumn(sheet, file, row, 6, "Drawn By", true);

                    sheet.Cell(row, 7).Value = file.CurrentState.Name;
                    sheet.Cell(row, 8).Value = file.CurrentVersion;

                    UpdateColumn(sheet, file, row, 9, "Notes", true);
                    UpdateColumn(sheet, file, row, 10, "Inspection Notes", true);

                    IEdmHistory2 history = (IEdmHistory2)vault.CreateUtility(EdmUtility.EdmUtil_History);
                    history.AddFile(file.ID);
                    EdmHistoryItem[] ppoRethistory = null;

                    history.GetHistory(ref ppoRethistory, (int)EdmHistoryType.Edmhist_FileState);
                    sheet.Cell(row, 11).Value = ppoRethistory[0].mbsComment;

                    history.GetHistory(ref ppoRethistory, (int)EdmHistoryType.Edmhist_FileVersion);
                    sheet.Cell(row, 12).Value = ppoRethistory[0].mbsComment;
                    done = true;
                }
            }
        }

        private void UpdateColumn(IXLWorksheet sheet, IEdmFile5 file, int row, int column, String name, Boolean isString)
        {
            IEdmEnumeratorVariable8 EnumVarObj = default(IEdmEnumeratorVariable8);
            EnumVarObj = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable();
            object VarObj = null;

            if (EnumVarObj.GetVar(name, "@", out VarObj) == true)
            {
                if (isString)
                {
                    sheet.Cell(row, column).Value = Convert.ToString(VarObj);
                }
                else
                {
                    sheet.Cell(row, column).Value = Convert.ToInt64(VarObj);
                }

                textBox1.AppendText("Row " + row + " Column " + column + " set to " + VarObj.ToString() + Environment.NewLine);
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

        private XLWorkbook getWorkbook(String Path)
        {
            var workbook = new XLWorkbook(Path);
            textBox1.AppendText("Loaded workbook " + Path + Environment.NewLine);
            return workbook;
        }
        private XLWorkbook newWorkbook(String Path)
        {
            XLWorkbook workbook = new XLWorkbook();
            textBox1.AppendText("Created workbook " + Path + Environment.NewLine);
            return workbook;
        }

        /*
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






