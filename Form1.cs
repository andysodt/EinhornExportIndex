using EPDM.Interop.epdm;
using NanoXLSX;

namespace EinhornExportIndex
{
    public partial class Form1 : Form, IEdmCallback6
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
                String FileName = Folder.Name + " INDEX OUTPUT.xlsx";

                if (Folder != null)
                {
                    String Path = RootFolder + OutputFolder + Folder.Name + " INDEX OUTPUT.xlsx";

                    textBox1.AppendText("Workbook " + Path + Environment.NewLine);

                    Workbook workbook;

                    if (File.Exists(Path))
                    {
                        workbook = getWorkbook(Path);
                    }
                    else
                    {
                        workbook = newWorkbook(Path);

                        IEdmFolder5 VaultRootFolder = default(IEdmFolder5);
                        VaultRootFolder = (IEdmFolder5)vault.RootFolder;
                        VaultRootFolder.AddFile(this.Handle.ToInt32(), "", Path, 0);

                        textBox1.AppendText("Added to vault: " + Path);

                        UnlockFile(Path);
                    }

                    LockFile(Path);
                    TraverseFolder(Folder, workbook);
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

        private void TraverseFolder(IEdmFolder5 CurFolder, Workbook workbook)
        {
            try
            {
                if (CurFolder.Name.EndsWith("-DRAWINGS"))
                {
                    textBox1.AppendText("Updating worksheet " + CurFolder.Name + Environment.NewLine);
                    AddWorksheet(workbook, CurFolder);
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

        private Workbook getWorkbook(String Path)
        {
            Workbook workbook = Workbook.Load(Path);
            workbook.Filename = Path;
            textBox1.AppendText("Loaded workbook " + Path);
            return workbook;
        }
        private Workbook newWorkbook(String Path)
        {
            Workbook workbook = new Workbook();
            workbook.Filename = Path;
            textBox1.AppendText("Created workbook " + Path);
            return workbook;
        }

        private void AddVarColumn(Workbook workbook, IEdmEnumeratorVariable8 EnumVarObj, String Name)
        {
            object VarObj = null;
            if (EnumVarObj.GetVar(Name, "@", out VarObj) == true)
            {
                workbook.CurrentWorksheet.AddNextCell(VarObj.ToString());
            }
            else
            {
                workbook.CurrentWorksheet.AddNextCell("");
            }
        }

        private void AddWorksheet(Workbook workbook, IEdmFolder5 Folder)
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

            while (!FilePos.IsNull)
            {

                file = Folder.GetNextFile(FilePos);
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

            workbook.Save();

        }

        private EdmMBoxResult IEdmCallback6_MsgBox(int lParentWnd, int lMsgID, string bsMsg, EdmMBoxType eType = 0L)
        {
            MessageBox.Show(bsMsg);
            return EdmMBoxResult.EdmMbr_OK;
        }
        EdmMBoxResult IEdmCallback6.MsgBox(int lParentWnd, int lMsgID, string bsMsg, EdmMBoxType eType)
        {
            return IEdmCallback6_MsgBox(lParentWnd, lMsgID, bsMsg, eType);
        }

        private void IEdmCallback6_Resolve(int lParentWnd, ref EdmCmdData[] ppoItems)
        {
        }
        void IEdmCallback6.Resolve(int lParentWnd, ref EdmCmdData[] ppoItems)
        {
            IEdmCallback6_Resolve(lParentWnd, ref ppoItems);
        }

        private bool IEdmCallback6_SetProgress(int lBarIndex, int lPos, string bsMsg)
        {
            return true;
        }
        bool IEdmCallback6.SetProgress(int lBarIndex, int lPos, string bsMsg)
        {
            return IEdmCallback6_SetProgress(lBarIndex, lPos, bsMsg);
        }

        private void IEdmCallback6_SetProgressRange(int lBarIndex, int lMax)
        {
        }
        void IEdmCallback6.SetProgressRange(int lBarIndex, int lMax)
        {
            IEdmCallback6_SetProgressRange(lBarIndex, lMax);
        }

        private void IEdmCallback6_SetStatusMessage(int lBarIndex, string bsMessage)
        {
        }
        void IEdmCallback6.SetStatusMessage(int lBarIndex, string bsMessage)
        {
            IEdmCallback6_SetStatusMessage(lBarIndex, bsMessage);
        }

    }
}






