using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;       //호환되지 않은DLL을 사용할 때.

using System.IO;
using System.Diagnostics;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using System.IO;

namespace FCMF0814
{
    public partial class FCMF0814 : Office2007Form
    {
        #region ----- API Dll Import -----
  
        [DllImport("fcrypt_es.dll")]
        extern public static int DSFC_EncryptFile(int hWnd, string pszPlainFilePathName, string pszEncFilePathName, string pszPassword, uint nOption);

        #endregion;

        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string inputPath;
        string OutputPath;
        string Password;
        uint DSFC_OPT_OVERWRITE_OUTPUT;
        int nRet;

        #endregion;

        #region ----- Constructor -----

        public FCMF0814()
        {
            InitializeComponent();
        }

        public FCMF0814(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private DateTime GetDate()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                idcGetDate.ExecuteNonQuery();
                object vObject = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch (Exception ex)
            {
                string vMessage = ex.Message;
                vDateTime = new DateTime(9999, 12, 31, 23, 59, 59);
            }
            return vDateTime;
        }

        private void Set_Default_Value()
        {
            //세금계산서 발행기간.
            DateTime vGetDateTime = GetDate();
            W_PERIOD_YEAR.EditValue = iDate.ISYear(vGetDateTime);

            //사업장 구분.
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "TAX_CODE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            W_TAX_CODE_NAME.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            W_TAX_CODE.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");
             
            //제출자
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "VAT_PRESENTER_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            VAT_PRESENTER_TYPE_NAME.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            VAT_PRESENTER_TYPE.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");

            //환급 구분.
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "TAX_REFUND_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            VAT_REFUND_TYPE_NAME.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            VAT_REFUND_TYPE.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");

            WRITE_DATE.EditValue = vGetDateTime;
        }

        private void SEARCH_DB()
        {
            if (iString.ISNull(W_TAX_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_TAX_CODE_NAME.Focus();
                return;
            }

            if (iString.ISNull(W_VAT_PERIOD_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10487"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_PERIOD_DESC.Focus();
                return;
            }
            if (iString.ISNull(W_ISSUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(W_ISSUE_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            if (Convert.ToDateTime(W_ISSUE_DATE_FR.EditValue) > Convert.ToDateTime(W_ISSUE_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }

            idaREPORT.Fill();
            idaREPORT_FILE.Fill(); 
        }

        private bool VAT_PERIOD_CHECK()
        {
            //신고기간 검증.
            string vCHECK_YN = "N";
            idcVAT_PERIOD_CHECK.ExecuteNonQuery();
            vCHECK_YN = iString.ISNull(idcVAT_PERIOD_CHECK.GetCommandParamValue("O_YN"));
            if (vCHECK_YN == "N")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10396"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return false;
            }
            return true;
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }
         
        #endregion;

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

        #region ----- XL Print 1 (매입) Method ----

        private void XLPrinting_1(string pOutChoice, ISDataAdapter pData1)
        {// pOutChoice : 출력구분.
            //string vMessageText = string.Empty;
            //string vSaveFileName = string.Empty;

            //int vCountRow = pData1.OraSelectData.Rows.Count;

            //if (vCountRow < 1)
            //{
            //    vMessageText = string.Format("Without Data");
            //    isAppInterfaceAdv1.OnAppMessage(vMessageText);
            //    System.Windows.Forms.Application.DoEvents();
            //    return;
            //}

            //System.Windows.Forms.Application.UseWaitCursor = true;
            //this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            //System.Windows.Forms.Application.DoEvents();

            //int vPageNumber = 0;

            //vMessageText = string.Format(" Printing Starting...");
            //isAppInterfaceAdv1.OnAppMessage(vMessageText);
            //System.Windows.Forms.Application.DoEvents();

            //XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            //try
            //{// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
            //    idcVAT_PERIOD.ExecuteNonQuery();
            //    string vPeriod = string.Format("( {0} )", idcVAT_PERIOD.GetCommandParamValue("O_PERIOD"));
            //    string vISSUE_PERIOD = String.Format("({0:D2}월 {1:D2}일 ~ {2:D2}월 {3:D2}일)", ISSUE_PERIOD_FR.DateTimeValue.Month, ISSUE_PERIOD_FR.DateTimeValue.Day, ISSUE_DATE_TO.DateTimeValue.Month, ISSUE_DATE_TO.DateTimeValue.Day);
                
            //    // open해야 할 파일명 지정.
            //    //-------------------------------------------------------------------------------------
            //    xlPrinting.OpenFileNameExcel = "FCMF0814_001.xls";
            //    //-------------------------------------------------------------------------------------
            //    // 파일 오픈.
            //    //-------------------------------------------------------------------------------------
            //    bool isOpen = xlPrinting.XLFileOpen();
            //    //-------------------------------------------------------------------------------------

            //    //-------------------------------------------------------------------------------------
            //    if (isOpen == true)
            //    {
            //        // 헤더 인쇄.
            //        idaREPORT.Fill();
            //        if (idaREPORT.SelectRows.Count > 0)
            //        {
            //            xlPrinting.HeaderWrite(idaREPORT, vPeriod, vISSUE_PERIOD);
            //        }

            //        //과세표준인쇄.
            //        idaPRINT_TAX_STANDARD.Fill();
            //        if (igrPRINT_TAX_STANDARD.RowCount > 0)
            //        {
            //            xlPrinting.XLLine_3(igrPRINT_TAX_STANDARD);
            //        }

            //        // 실제 인쇄
            //        vPageNumber = xlPrinting.LineWrite(pData1, pData2);

            //        //출력구분에 따른 선택(인쇄 or file 저장)
            //        if (pOutChoice == "PRINT")
            //        {
            //            xlPrinting.Printing(1, vPageNumber);
            //        }
            //        else if (pOutChoice == "FILE")
            //        {
            //            xlPrinting.SAVE("VAT_1_");
            //        }

            //        //-------------------------------------------------------------------------------------
            //        xlPrinting.Dispose();
            //        //-------------------------------------------------------------------------------------

            //        vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
            //        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            //        System.Windows.Forms.Application.DoEvents();
            //    }
            //    else
            //    {
            //        vMessageText = "Excel File Open Error";
            //        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            //        System.Windows.Forms.Application.DoEvents();
            //    }
            //    //-------------------------------------------------------------------------------------
            //}
            //catch (System.Exception ex)
            //{
            //    xlPrinting.Dispose();

            //    vMessageText = ex.Message;
            //    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            //    System.Windows.Forms.Application.DoEvents();
            //}

            //System.Windows.Forms.Application.UseWaitCursor = false;
            //this.Cursor = System.Windows.Forms.Cursors.Default;
            //System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Text File Export Methods ----

        private void ExportTXT(object pENCRYPT_PASSWORD, ISDataAdapter pData)
        {
            int vCountRow = pData.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                return;
            }

            isAppInterfaceAdv1.OnAppMessage("Export Text Start...");

            string vSaveTextFileName = String.Empty;
            string vFileName = string.Empty;
            string vFilePath = "C:\\ersdata"; 

            int euckrCodepage = 51949; 
            
            System.IO.FileStream vWriteFile = null;
            System.Text.StringBuilder vSaveString = new System.Text.StringBuilder();

            //파일 경로 디렉토리 존재 여부 체크(없으면 생성).
            if (System.IO.Directory.Exists(vFilePath) == false)
            {
                System.IO.Directory.CreateDirectory(vFilePath);
            }
            vFileName = WRITE_DATE.DateTimeValue.ToShortDateString().Replace("-", "");

            saveFileDialog1.Title = "Save File";
            saveFileDialog1.FileName = vFileName;
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(vFilePath);
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Text Files (*.101)|*.101";
            saveFileDialog1.DefaultExt = ".101";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                Application.DoEvents();
            
                vSaveTextFileName = saveFileDialog1.FileName;
                try
                {
                    vWriteFile = System.IO.File.Open(vSaveTextFileName, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);
                    foreach(DataRow cRow in pData.OraSelectData.Rows)
                    {
                        vSaveString = new System.Text.StringBuilder();  //초기화.
                        vSaveString.Append(cRow["REPORT_FILE"]);
                        vSaveString.Append("\r\n");

                        //기존
                        //byte[] vSaveBytes = new System.Text.UnicodeEncoding().GetBytes(vSaveString.ToString());

                        //신규.
                        System.Text.Encoding vEUCKR = System.Text.Encoding.GetEncoding(euckrCodepage);
                        byte[] vSaveBytes = vEUCKR.GetBytes(vSaveString.ToString()); 

                        int vSaveStrigLength = vSaveBytes.Length;
                        vWriteFile.Write(vSaveBytes, 0, vSaveStrigLength);
                    }
                }
                catch (System.Exception ex)
                {
                    string vMessage = ex.Message;
                    isAppInterfaceAdv1.OnAppMessage(vMessage);
                    Application.DoEvents();
                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                }
                isAppInterfaceAdv1.OnAppMessage("Export Text End");
                vWriteFile.Dispose();

                //기존 동일한 파일 삭제.
                if (System.IO.File.Exists(vSaveTextFileName) == false)
                {
                    MessageBoxAdv.Show("암호화 대상 전자파일이 존재하지 않습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                //암호화는 안함.
                //nRet = 0;
                //inputPath = vSaveTextFileName;// "20120410.201";//pFileName;
                //OutputPath = string.Format("{0}.erc", vSaveTextFileName);
                //Password = pENCRYPT_PASSWORD.ToString();
                //DSFC_OPT_OVERWRITE_OUTPUT = 1;
                //nRet = DSFC_EncryptFile(0, inputPath, OutputPath, Password, DSFC_OPT_OVERWRITE_OUTPUT);
                //if (nRet != 0)
                //{
                //    MessageBox.Show("Encrypt Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}

                //System.IO.File.Delete(vSaveTextFileName);
                //System.IO.File.Copy(inputPath, inputPath, true);
                //System.IO.File.Delete(OutputPath);                

                //폴더 열기.
                Process.Start(System.IO.Path.GetDirectoryName(vSaveTextFileName));
            }
            Application.DoEvents();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default; 
        }

        public void ExportTXT_File(ISDataAdapter pData)
        {
        //    int vCountRow = pData.OraSelectData.Rows.Count;
        //    if (vCountRow < 1)
        //    {
        //        return;
        //    }

        //    isAppInterfaceAdv1.OnAppMessage("Export Text Start...");

        //    System.IO.Stream vWrite = null; ;
        //    System.Text.StringBuilder vSaveString = new System.Text.StringBuilder();

        //    saveFileDialog1.Title = "Save File";
        //    saveFileDialog1.FileName = WRITE_DATE.DateTimeValue.ToShortDateString().Replace("-", "");
        //    saveFileDialog1.DefaultExt = ".101";
        //    System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
        //    saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
        //    saveFileDialog1.Filter = "Text Files (*.101)|*.101";
        //    if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        //    {
        //        Application.UseWaitCursor = true;
        //        this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //        Application.DoEvents();

        //        string vsSaveTextFileName = saveFileDialog1.FileName;
        //        try
        //        {
        //            //vWriteFile = System.IO.File.Open(vsSaveTextFileName, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);
        //            vWrite = System.IO.File.OpenWrite(vsSaveTextFileName);
        //            foreach (DataRow cRow in pData.OraSelectData.Rows)
        //            {
        //                vSaveString = new System.Text.StringBuilder();  //초기화.
        //                vSaveString.Append(cRow["REPORT_FILE"]);
        //                vSaveString.Append("\r\n");

        //                System.IO.StreamWriter(vWrite, Encoding.Default);

        //                //byte[] vSaveBytes = new System.Text.UnicodeEncoding().GetBytes(vSaveString.ToString());
        //                //int vSaveStrigLength = vSaveBytes.Length;
        //                //vWriteFile.Write(vSaveBytes, 0, vSaveStrigLength);
        //            }
        //        }
        //        catch (System.Exception ex)
        //        {
        //            string vMessage = ex.Message;
        //            isAppInterfaceAdv1.OnAppMessage(vMessage);
        //            Application.DoEvents();
        //            Application.UseWaitCursor = false;
        //            this.Cursor = System.Windows.Forms.Cursors.Default;
        //        }

        //        isAppInterfaceAdv1.OnAppMessage("Export Text End");
        //        vWriteFile.Dispose();
        //    }
        //    Application.DoEvents();
        //    Application.UseWaitCursor = false;
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        } 

        #endregion;
        
        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                //신고기간 검증.
                if (VAT_PERIOD_CHECK() == false)
                {
                    return;
                }

                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        idaREPORT.Update();
                    }
                    catch (Exception Ex)
                    {
                        MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }                
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaREPORT.IsFocused)
                    {
                        idaREPORT.Cancel();
                    }                  
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    //XLPrinting_1("PRINT", idaREPORT);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    //XLPrinting_1("FILE", idaREPORT);
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void FCMF0814_Load(object sender, EventArgs e)
        {
             
        }

        private void FCMF0814_Shown(object sender, EventArgs e)
        {
            W_ISSUE_DATE_FR.BringToFront();
            W_ISSUE_DATE_TO.BringToFront();

            Set_Default_Value();
        }

        private void itbREPORT_Click(object sender, EventArgs e)
        {
            if (itbREPORT.SelectedTab.TabIndex == 1)
            {
                ISSUE_DATE_FR.Focus();
            }            
        }

        private void BTN_SET_REPORT_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_TAX_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_TAX_CODE_NAME.Focus();
                return;
            }

            if (iString.ISNull(W_VAT_PERIOD_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10487"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_VAT_PERIOD_DESC.Focus();
                return;
            }
            if (iString.ISNull(ISSUE_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(ISSUE_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_TO.Focus();
                return;
            }
            if (Convert.ToDateTime(ISSUE_DATE_FR.EditValue) > Convert.ToDateTime(ISSUE_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_ISSUE_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(ISSUE_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10298"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WRITE_DATE.Focus();
                return;
            }

            try
            {
                idaREPORT.Update();
            }
            catch (Exception Ex)
            {
                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //신고기간 검증.
            if (VAT_PERIOD_CHECK() == false)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            IDC_SET_REPORT_FILE.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_SET_REPORT_FILE.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_SET_REPORT_FILE.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_REPORT_FILE.ExcuteError)
            {
                MessageBoxAdv.Show(IDC_SET_REPORT_FILE.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            if(vSTATUS == "S")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10112"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            SEARCH_DB();
        }

        private void BTN_SAVE_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(E_FILE_YN.EditValue) != "Y")
            {
                MessageBoxAdv.Show("먼저 신고파일을 생성하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);                
                return;
            }
             
            ////전산매체 암호화 암호 입력 받기.
            //DialogResult vdlgResult;
            object vENCRYPT_PASSWORD = String.Empty;
            //FCMF0814_FILE vFCMF0814_FILE = new FCMF0814_FILE(isAppInterfaceAdv1.AppInterface);
            //vdlgResult = vFCMF0814_FILE.ShowDialog();
            //if (vdlgResult == DialogResult.OK)
            //{
            //    vENCRYPT_PASSWORD = vFCMF0814_FILE.Get_Encrypt_Password;
            //}

            //if (iString.ISNull(vENCRYPT_PASSWORD) == string.Empty)
            //{
            //    return;
            //}

            idaREPORT_FILE.Fill();
            ExportTXT(vENCRYPT_PASSWORD, idaREPORT_FILE);

            //string mMESSAGE;
            //idcSET_REPORT_FILE.ExecuteNonQuery();
            //mMESSAGE = iString.ISNull(idcSET_REPORT_FILE.GetCommandParamValue("O_MESSAGE"));
            //if (mMESSAGE != String.Empty)
            //{
            //    MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaTAX_CODE_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("TAX_CODE", "Y");
        }

        private void ilaTAX_CODE_0_SelectedRowData(object pSender)
        {
            W_VAT_PERIOD_DESC.EditValue = string.Empty;
            W_VAT_PERIOD_ID.EditValue = string.Empty;
            W_ISSUE_DATE_FR.EditValue = DBNull.Value;
            W_ISSUE_DATE_TO.EditValue = DBNull.Value;
        }

        private void ilaVAT_LEVIER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("VAT_LEVIER", "Y");
        }

        private void ILA_VAT_REPORT_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("VAT_REPORT_TYPE", "Y");
        }

        private void ILA_VAT_PRESENTER_TYPE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            SetCommonParameter("VAT_PRESENTER_TYPE", "Y");
        }
   
        private void ilaVAT_REFUND_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("VAT_REFUND_TYPE", "Y");
        }
         
        #endregion

        #region ----- Adapter Event ------
         
        private void idaREPORT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["DECLARATION_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show("부가가치세 신고서 정보를 찾을수 없습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
            if (iString.ISNull(e.Row["HOMETAX_LOGIN_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show("홈택스ID는 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["VAT_LEVIER_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("일반과세자구분은 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["VAT_PRESENTER_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("제출자구분은 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["VAT_REFUND_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("환급구분은 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        private void idaREPORT_PreDelete(ISPreDeleteEventArgs e)
        {       
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10047"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            
        }

        #endregion




    }
}