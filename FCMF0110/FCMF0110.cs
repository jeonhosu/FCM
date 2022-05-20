using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using System.IO;
using Syncfusion.GridExcelConverter;
using Syncfusion.XlsIO;

namespace FCMF0110
{
    public partial class FCMF0110 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public FCMF0110(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Property / Method -----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void DEFAULT_ACCOUNT_SET()
        {
            IDC_DFV_ACCOUNT_SET.ExecuteNonQuery();

            W_ACCOUNT_SET_NAME.EditValue = IDC_DFV_ACCOUNT_SET.GetCommandParamValue("O_ACCOUNT_SET_NAME");
            W_ACCOUNT_SET_ID.EditValue = IDC_DFV_ACCOUNT_SET.GetCommandParamValue("O_ACCOUNT_SET_ID");
        }

        private void Init_Account_Set()
        {
            AS_ACCOUNT_LEVEL.EditValue = 0; 
            AS_ACCOUNT_SET_ID.Focus();
        }

        private void Init_Account_Control_Insert()
        {
            ACCOUNT_BALANCE_FLAG.CheckBoxValue = "Y";
            ENABLED_FLAG.CheckBoxValue = "Y";
            EFFECTIVE_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            ADD_ON_MONTH.EditValue = 0;
            CLOSED_DAY.EditValue = 0;

            ACCOUNT_CODE.Focus();
        }
         
        private void isSetCommonParameter(string pGroup_Code, string pEnabled_Flag)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_Flag);
        }

        private void SEARCH_DB()
        {
            if (TB_ACCOUNT_CONTROL.SelectedTab.TabIndex == TP_ACCOUNT_LIST.TabIndex)
            {
                IDA_ACCOUNT_CONTROL_LIST.Fill();
                IGR_ACCOUNT_CONTROL_LIST.Focus();
            }
            else
            {                
                IDA_ACCOUNT_SET.Fill();
                AS_ACCOUNT_SET_ID.Focus(); 
            }
        }

        private void Show_Print_Form(string pOutChoice)
        {
            if (iString.ISNull(W_ACCOUNT_SET_ID.EditValue) == string.Empty)
            {
                W_ACCOUNT_SET_NAME.Focus();
                return;
            }

            string vMessageText = string.Empty;
            object vAccount_Code_Fr = null;
            object vAccount_Code_To = null;
            DialogResult dlgResult;

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            FCMF0110_PRN vFCMF0110_PRN = new FCMF0110_PRN(isAppInterfaceAdv1.AppInterface, pOutChoice, W_ENABLED_FLAG.CheckBoxValue);
            dlgResult = vFCMF0110_PRN.ShowDialog();
            if (dlgResult != DialogResult.OK)
            {                
                vFCMF0110_PRN.Dispose();
                Application.UseWaitCursor = false;
                this.Cursor = System.Windows.Forms.Cursors.Default;
                Application.DoEvents();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
            vAccount_Code_Fr = vFCMF0110_PRN.Get_Account_Code_Fr;
            vAccount_Code_To = vFCMF0110_PRN.Get_Account_Code_To;

            idaPRINT_ACCOUNT.SetSelectParamValue("W_ACCOUNT_CODE_FR", vAccount_Code_Fr);
            idaPRINT_ACCOUNT.SetSelectParamValue("W_ACCOUNT_CODE_TO", vAccount_Code_To);
            idaPRINT_ACCOUNT.Fill();
            int vCountRow = idaPRINT_ACCOUNT.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                vMessageText = string.Format(isMessageAdapter1.ReturnText("FCM_10386"));
                isAppInterfaceAdv1.OnAppMessage(vMessageText);

                Application.UseWaitCursor = false;
                this.Cursor = System.Windows.Forms.Cursors.Default;
                Application.DoEvents();
                return;
            }
            XLPrinting_1(pOutChoice, idaPRINT_ACCOUNT);
        }

        #endregion

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx pGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                //xls 저장방법
                //vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                //                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);



                //if (MessageBox.Show("Do you wish to open the xls file now?",
                //                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                //    vProc.StartInfo.FileName = saveFileDialog.FileName;
                //    vProc.Start();
                //}

                //xlsx 파일 저장 방법
                GridExcelConverterControl converter = new GridExcelConverterControl();
                ExcelEngine excelEngine = new ExcelEngine();
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2007;
                IWorkbook workBook = ExcelUtils.CreateWorkbook(1);
                workBook.Version = ExcelVersion.Excel2007;
                IWorksheet sheet = workBook.Worksheets[0];
                //used to convert grid to excel 
                converter.GridToExcel(pGrid.BaseGrid, sheet, ConverterOptions.ColumnHeaders);
                //used to save the file
                workBook.SaveAs(saveFileDialog.FileName);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                        "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = saveFileDialog.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion


        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pOutChoice, ISDataAdapter pAdapter)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            object vAccountBook = AS_ACCOUNT_SET_NAME.EditValue;
            object vToday = DateTime.Today.ToShortDateString();

            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            //출력구분이 파일인 경우 처리.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("Accounts_{0}", vToday);

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vSaveFileName = saveFileDialog1.FileName;
                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                vMessageText = string.Format(" Writing Starting...");
            }
            else
            {
                vMessageText = string.Format(" Printing Starting...");
            }
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int vPageNumber = 0;
            //int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "FCMF0110_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    // 헤더 부분 인쇄.
                    xlPrinting.HeaderWrite(vAccountBook, vToday);

                    // 라인 인쇄
                    vPageNumber = xlPrinting.LineWrite(idaPRINT_ACCOUNT);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Application_MainButtonClick -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_ACCOUNT_SET.IsFocused)
                    {
                        IDA_ACCOUNT_SET.AddOver();
                        Init_Account_Set();
                    } 
                    else if (IDA_ACCOUNT_CONTROL.IsFocused)
                    {
                        IDA_ACCOUNT_CONTROL.AddOver();
                        Init_Account_Control_Insert();
                    }
                    else if (IDA_ACCOUNT_CONTROL_ITEM.IsFocused)
                    {
                        IDA_ACCOUNT_CONTROL_ITEM.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_ACCOUNT_SET.IsFocused)
                    {
                        IDA_ACCOUNT_SET.AddUnder();
                        Init_Account_Set();
                    } 
                    else if (IDA_ACCOUNT_CONTROL.IsFocused)
                    {
                        IDA_ACCOUNT_CONTROL.AddUnder();
                        Init_Account_Control_Insert();
                    }
                    else if (IDA_ACCOUNT_CONTROL_ITEM.IsFocused)
                    {
                        IDA_ACCOUNT_CONTROL_ITEM.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_ACCOUNT_SET.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ACCOUNT_SET.IsFocused)
                    {
                        IDA_ACCOUNT_CONTROL.Cancel();
                        IDA_ACCOUNT_SET.Cancel();
                    } 
                    else if(IDA_ACCOUNT_CONTROL.IsFocused)
                    {
                        IDA_ACCOUNT_CONTROL_ITEM.Cancel();
                        IDA_ACCOUNT_CONTROL.Cancel();
                    }
                    else if (IDA_ACCOUNT_CONTROL_ITEM.IsFocused)
                    {
                        IDA_ACCOUNT_CONTROL_ITEM.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ACCOUNT_SET.IsFocused)
                    {
                        IDA_ACCOUNT_SET.Delete();
                    } 
                    else if (IDA_ACCOUNT_CONTROL_ITEM.IsFocused)
                    {
                        IDA_ACCOUNT_CONTROL_ITEM.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    Show_Print_Form("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(IGR_ACCOUNT_CONTROL_LIST);
                }
            }
        }

        #endregion

        #region ----- Form Event ------

        private void FCMF0110_Load(object sender, EventArgs e)
        {
            IDA_ACCOUNT_SET.FillSchema();            
        }

        private void FCMF0110_Shown(object sender, EventArgs e)
        {
            DEFAULT_ACCOUNT_SET(); 
        }

        private void BTN_LANG_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(AS_ACCOUNT_SET_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("Account Set must selected. Check it", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dlgResult;
            FCMF0110_TL vFCMF0110_TL = new FCMF0110_TL(this.MdiParent, isAppInterfaceAdv1.AppInterface,
                                                        AS_ACCOUNT_SET_ID.EditValue, AS_ACCOUNT_SET_CODE.EditValue, 
                                                        AS_ACCOUNT_SET_NAME.EditValue, AS_ACCOUNT_LEVEL.EditValue);
            vFCMF0110_TL.ShowDialog();
            vFCMF0110_TL.Dispose();
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaACCOUNT_SET_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["ACCOUNT_SET_ID"] == DBNull.Value)
            {// 계정SET ID
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10086", "&&VALUE:=Account Set ID(계정 세트ID)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_SET_CODE"]) == string.Empty)
            {// 계정세트코드
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10086", "&&VALUE:=Account Set Code(계정 세트코드)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_SET_NAME"]) == string.Empty)
            {// 계정세트명
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10086", "&&VALUE:=Account Set Name(계정 세트명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ACCOUNT_LEVEL"] == DBNull.Value)
            {// 계정세트레벨
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10086", "&&VALUE:=Account Level(계정 세트 레벨)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
        }

        private void idaACCOUNT_SET_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaACCOUNT_CONTROL_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (AS_ACCOUNT_LEVEL.EditValue == null)
            {// 계정세트레벨
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Account Level(계정 세트 레벨)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_CODE"]) == string.Empty)
            {// 계정코드
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Account Code(계정 코드)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_DESC"]) == string.Empty)
            {// 계정명
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10087", "&&VALUE:=Account Desc(계정명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(e.Row["ACCOUNT_LEVEL"], 0) == Convert.ToInt32(0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10160"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(e.Row["ACCOUNT_LEVEL"], 0) > 1 && iString.ISNull(UPPER_ACCOUNT_CONTROL_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(UPPER_ACCOUNT_CONTROL_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(AS_ACCOUNT_LEVEL.EditValue, 0) < iString.ISNumtoZero(e.Row["ACCOUNT_LEVEL"], 0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10133"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_DR_CR"]) == string.Empty)
            {// 차대구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10088", "&&VALUE:=Account DR/CR(차대 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_PROPERTY_NAME"]) == string.Empty)
            {// 차대구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ACCOUNT_PROPERTY_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_FS_TYPE_NAME"]) == string.Empty)
            {// 차대구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ACCOUNT_FS_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }   
            if (e.Row["EFFECTIVE_DATE_FR"] == DBNull.Value)
            {// 적용 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10088", "&&VALUE:=Effective Date From(적용 시작일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }               
        }

        private void idaACCOUNT_CONTROL_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10047"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaACCOUNT_CONTROL_ITEM_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            ////////////////
            if (e.Row["MANAGEMENT1_ID"] == DBNull.Value && e.Row["MANAGEMENT1_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item1(관리항목1)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["MANAGEMENT2_ID"] == DBNull.Value && e.Row["MANAGEMENT2_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item2(관리항목2)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER1_ID"] == DBNull.Value && e.Row["REFER1_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item3(관리항목3)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER2_ID"] == DBNull.Value && e.Row["REFER2_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item4(관리항목4)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER3_ID"] == DBNull.Value && e.Row["REFER3_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item5(관리항목5)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER4_ID"] == DBNull.Value && e.Row["REFER4_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item6(관리항목6)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER5_ID"] == DBNull.Value && e.Row["REFER5_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item7(관리항목7)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER6_ID"] == DBNull.Value && e.Row["REFER6_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item8(관리항목8)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER7_ID"] == DBNull.Value && e.Row["REFER7_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item9(관리항목9)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER8_ID"] == DBNull.Value && e.Row["REFER8_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item10(관리항목10)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER9_ID"] == DBNull.Value && e.Row["REFER9_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Item11(관리항목11)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER_RATE_ID"] == DBNull.Value && e.Row["REFER_RATE_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Rate(관리율)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER_AMOUNT_ID"] == DBNull.Value && e.Row["REFER_AMOUNT_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Amount(관리 금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER_DATE1_ID"] == DBNull.Value && e.Row["REFER_DATE1_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Date1(관리일자1)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["REFER_DATE2_ID"] == DBNull.Value && e.Row["REFER_DATE2_YN"].ToString() == "Y".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Management Date2(관리일자2)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            //if (e.Row["VOUCH_ID"] == DBNull.Value && e.Row["VOUCH_YN"].ToString() == "Y".ToString())
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10090", "&&FIELD_NAME:=Vouch Y/N(증빙 유무)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            /////////////////////////////////////
            
            //if (e.Row["MANAGEMENT1_ID"] == DBNull.Value && e.Row["MANAGEMENT2_ID"] != DBNull.Value)
            //{// 계정코드
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10089"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}

            //if (e.Row["REFER1_ID"] == DBNull.Value && e.Row["REFER2_ID"] != DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10089"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["REFER2_ID"] == DBNull.Value && e.Row["REFER3_ID"] != DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10089"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["REFER3_ID"] == DBNull.Value && e.Row["REFER4_ID"] != DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10089"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["REFER4_ID"] == DBNull.Value && e.Row["REFER5_ID"] != DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10089"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["REFER5_ID"] == DBNull.Value && e.Row["REFER6_ID"] != DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10089"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["REFER6_ID"] == DBNull.Value && e.Row["REFER7_ID"] != DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10089"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["REFER7_ID"] == DBNull.Value && e.Row["REFER8_ID"] != DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10089"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["REFER8_ID"] == DBNull.Value && e.Row["REFER9_ID"] != DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10089"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        }

        private void idaACCOUNT_CONTROL_ITEM_PreDelete(ISPreDeleteEventArgs e)
        {
        }

        private void idaACCOUNT_SET_ExcuteKeySearch(object pSender)
        {
            SEARCH_DB();
        }

        #endregion
        
        #region ----- Lookup Event -----

        private void ilaACCOUNT_GROUP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL_ALL.SetLookupParamValue("W_ENABLED_YN", W_ENABLED_FLAG.CheckBoxValue);
        }

        private void ilaACCOUNT_DR_CR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("ACCOUNT_DR_CR", "Y");
        }

        private void ILA_ACCOUNT_PROPERTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("ACCOUNT_PROPERTY", "Y");
        }

        private void ILA_UPPER_ACCOUNT_CONTROL_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_UPPER_ACCOUNT_CONTROL.SetLookupParamValue("W_CURR_ACC_LEVEL", null);
            ILD_UPPER_ACCOUNT_CONTROL.SetLookupParamValue("W_CHILD_ACC_LEVEL", null);
            ILD_UPPER_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_UPPER_ACCOUNT_CONTROL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_UPPER_ACCOUNT_CONTROL.SetLookupParamValue("W_CURR_ACC_LEVEL", null);
            ILD_UPPER_ACCOUNT_CONTROL.SetLookupParamValue("W_CHILD_ACC_LEVEL", ACCOUNT_LEVEL.EditValue);
            ILD_UPPER_ACCOUNT_CONTROL.SetLookupParamValue("W_ACCOUNT_CODE_FR", null);
            ILD_UPPER_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ilaGL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("GL_TYPE", "Y");
        }

        private void ilaFS_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("FS_TYPE", "Y");
        }

        private void ilaLIQUIDATE_METHOD_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("LIQUIDATE_METHOD_TYPE", "Y");
        }

        private void ilaACCOUNT_CLASS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("ACCOUNT_CLASS", "Y");
        }

        private void ilaBALANCE_MANAGEMENT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("MANAGEMENT_CODE", "Y");
        }

        private void ilaREFER_ITEM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("MANAGEMENT_CODE", "Y");
        }

        private void ilaREFER_RATE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("REFERENCE_RATE", "Y");
        }

        private void ilaREFER_AMOUNT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("REFERENCE_AMOUNT", "Y");
        }

        private void ilaREFER_DATE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("REFERENCE_DATE", "Y");
        }

        private void ILA_CLOSED_DAY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("DAY_NUM", "Y"); 
        }

        private void ilaVOUCH_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonParameter("VOUCH_CODE", "Y");
        }

        #endregion


    }
}