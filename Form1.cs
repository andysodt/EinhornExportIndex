using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using EPDM.Interop.epdm;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Net;
using NanoXLSX;
using NanoXLSX.Styles;

namespace EinhornExportIndex
{
    public partial class Form1 : Form
    {
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

                Workbook workbook = getWorkbook();

                TraverseFolder(Folder, workbook);

                MessageBox.Show("Done!");


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

        private Workbook getWorkbook()
        {
            String filename = "C:\\temp\\index.xlsx";
            Workbook workbook = null;
            if (System.IO.File.Exists(filename))
            {
                workbook = Workbook.Load(filename);
            }
            else
            {
                workbook = new Workbook();
            }
            workbook.Filename = filename;

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

            List<object> values = new List<object>() { "File.ID", "Name", "Revision", "# Sheets", "Description", "Resp Eng", "Drawn By", "State", "Version", "Notes", "Inspection Notes", "State Comments", "Checkin Comments" };
            workbook.CurrentWorksheet.AddCellRange(values, new Address(0, 0), new Address(12, 0));
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
            workbook.CurrentWorksheet.Cells["M1"].SetStyle(NanoXLSX.Styles.BasicStyles.Bold);
            workbook.CurrentWorksheet.GoToNextRow();

            IEdmHistory2 history = (IEdmHistory2)vault.CreateUtility(EdmUtility.EdmUtil_History);
            IEdmPos5 FilePos = default(IEdmPos5);
            FilePos = Folder.GetFirstFilePosition();
            IEdmFile5 file = default(IEdmFile5);
            while (!FilePos.IsNull)
            {
                file = Folder.GetNextFile(FilePos);

                workbook.CurrentWorksheet.AddNextCell(file.ID);
                history.AddFile(file.ID);

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
                AddVarColumn(workbook, EnumVarObj, "State Comments");
                AddVarColumn(workbook, EnumVarObj, "Checkin Comments");

                workbook.CurrentWorksheet.GoToNextRow();
            }

            EdmHistoryItem[] ppoRethistory = null;
            history.GetHistory(ref ppoRethistory, (int)EdmHistoryType.Edmhist_FileVersion);
            workbook.Save();

        }
    }
}






