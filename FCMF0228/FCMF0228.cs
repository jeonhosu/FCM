using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace FCMF0228
{
    public partial class FCMF0228 : Office2007Form
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);


        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private ISFileTransferAdv mFileTransfer;
        private string mHost = string.Empty;
        private string mPort = string.Empty;
        private string mPassive = "N";
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;
        private string mFTP_Folder = string.Empty;
        private string mClient_Folder = string.Empty;

        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 실행 디렉토리.        
        private string mDownload_Folder = string.Empty;             // Download Folder 
        private bool mFTP_Connect_Status = false;                   // FTP 정보 상태.
        private bool mSave_Appr_Status = false;
        private string mView_Only_Flag = "N";

        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        Form mMdi_Parent = null;
        string mSource_Category = "";

        #endregion;

        #region ----- Constructor -----

        public FCMF0228()
        {
            InitializeComponent();
        }

        public FCMF0228(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            F_SLIP_DATE.EditValue = "2022-02-11";
            F_SLIP_NUM.EditValue = "EXP220211002";
            mView_Only_Flag = "N";
        }

        public FCMF0228(Form pMainForm, ISAppInterface pAppInterface, string pSource_Category, object pSlip_Date, object pSlip_Num, string pView_Only_Flag)
        {
            InitializeComponent();
            mMdi_Parent = pMainForm;

            pMainForm.Activated += PMainForm_Activated;
            //this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mSource_Category = pSource_Category;
            F_SLIP_DATE.EditValue = pSlip_Date;
            F_SLIP_NUM.EditValue = pSlip_Num;
            mView_Only_Flag = pView_Only_Flag;
        }

        private void PMainForm_Activated(object sender, EventArgs e)
        {
            Form vFrm = GetForm(this.Name);
            if (vFrm == null)
                return;

            //System.Diagnostics.Debug.WriteLine("Here!");
            IntPtr vIntPtr = GetIntPtr(vFrm.Text);

            ShowWindowAsync(vIntPtr, 1);
            SetForegroundWindow(vIntPtr);
        }
         
        public static Form GetForm(string pFormName)
        {
            foreach (Form vFrm in Application.OpenForms)
            {
                if(vFrm.Name.Replace("{", "").StartsWith(pFormName))
                {
                    return vFrm;
                }
            }
            return null;
        }

        public static IntPtr GetIntPtr(string pFormName)
        {
            IntPtr vIntPtr = FindWindow(null, pFormName);
            return vIntPtr; 
        }

        #endregion;

        #region ----- Private Methods -----
         
        private void Search_DB()
        {
             
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
           
        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                } 
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                { 
                }
            }
        }

        #endregion;

        #region ----- Form Event ----- 
        
        private void FCMF0228_Load(object sender, EventArgs e)
        {
            if (mMdi_Parent != null)
            {
                int left = mMdi_Parent.Left + mMdi_Parent.Width / 2 - this.Width / 2;
                int top = mMdi_Parent.Top + mMdi_Parent.Height / 2 - this.Height / 2;

                this.Location = new Point(left, top); 
            }
            this.TopLevel = true;

            IDA_DOC_ATTACHMENT.Fill();
            if(mView_Only_Flag.Equals("Y"))
            {
                BTN_ATT_SELECT.Visible = false;
                BTN_ATT_DELETE.Visible = false;
            }
            else
            {
                BTN_ATT_SELECT.Visible = true;
                BTN_ATT_DELETE.Visible = true;
            }
        }

        private void FCMF0228_Shown(object sender, EventArgs e)
        {
            Set_FTP_Info();
        }

        private void FCMF0228_Deactivate(object sender, EventArgs e)
        {
            Form vFrm = GetForm(this.Name);
            if (vFrm == null)
                return;

            IntPtr vIntPtr = GetIntPtr(vFrm.Text);
            //System.Diagnostics.Debug.WriteLine("Here!" + vIntPtr); 
            ShowWindowAsync(vIntPtr, 1);
            SetForegroundWindow(vIntPtr);
        }

        #endregion

        #region ---- Doc Att / Appr Step ----

        private void BTN_FILE_ATTACH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(iString.ISNull(F_SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(new Form { TopMost = true }, isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 

            //FTP 정보//
            Set_FTP_Info();
             
            IDA_DOC_ATTACHMENT.Fill();
            IGR_DOC_ATTACHMENT.Focus();
        }

        private void BTN_DOC_ATT_L_ButtonClick(object pSender, EventArgs pEventArgs)
        { 
            if (iString.ISNull(F_SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(new Form { TopMost = true }, isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 

            //FTP 정보//
            Set_FTP_Info();

            IDA_DOC_ATTACHMENT.Fill();
            IGR_DOC_ATTACHMENT.Focus();
        }

        private void IGR_DOC_ATTACHMENT_CellDoubleClick(object pSender)
        {
            //if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            //{
            //    return;
            //}

            //string vFTP_FILE_NAME = iString.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            //string vUSER_FILE_NAME = string.Format("{0}{1}", mDownload_Folder, IGR_DOC_ATTACHMENT.GetCellValue("USER_FILE_NAME"));
            //if (DownLoadFile(vFTP_FILE_NAME, vUSER_FILE_NAME) == false)
            //{
            //    return;
            //} 
        }

        private void BTN_ATT_SELECT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDC_GET_DOC_ATT_STATUS.SetCommandParamValue("P_SOURCE_CATEGORY", mSource_Category);
            IDC_GET_DOC_ATT_STATUS.ExecuteNonQuery();
            String vSTATUS = iString.ISNull(IDC_GET_DOC_ATT_STATUS.GetCommandParamValue("O_STATUS"));
            String vMESSAGE  = iString.ISNull(IDC_GET_DOC_ATT_STATUS.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(new Form { TopMost = true }, vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            UpLoadFile(F_SLIP_DATE.EditValue, F_SLIP_NUM.EditValue);
            IDA_DOC_ATTACHMENT.Fill();
            IGR_DOC_ATTACHMENT.Focus();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_ATT_DOWN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            {
                return;
            }

            object vDOC_ATTACHMENT_ID = IGR_DOC_ATTACHMENT.GetCellValue("DOC_ATTACHMENT_ID");
            string vFTP_FILE_NAME = iString.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            string vUSER_FILE_NAME = string.Format("{0}{1}", mDownload_Folder, IGR_DOC_ATTACHMENT.GetCellValue("USER_FILE_NAME"));
            if (DownLoadFile(vDOC_ATTACHMENT_ID, vFTP_FILE_NAME, vUSER_FILE_NAME) == false)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            } 
        }

        private void BTN_ATT_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(new Form { TopMost = true }, isMessageAdapter1.ReturnText("EAPP_10220"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            if (iString.ISNull(F_SLIP_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(new Form { TopMost = true }, isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDC_GET_DOC_ATT_STATUS.SetCommandParamValue("P_SOURCE_CATEGORY", mSource_Category);
            IDC_GET_DOC_ATT_STATUS.ExecuteNonQuery();
            String vSTATUS = iString.ISNull(IDC_GET_DOC_ATT_STATUS.GetCommandParamValue("O_STATUS"));
            String vMESSAGE = iString.ISNull(IDC_GET_DOC_ATT_STATUS.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(new Form { TopMost = true }, vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            {
                return;
            }

            object vDOC_ATTACHMENT_ID = IGR_DOC_ATTACHMENT.GetCellValue("DOC_ATTACHMENT_ID"); 
            string vFTP_FileName = iString.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            DeleteFile(vDOC_ATTACHMENT_ID, vFTP_FileName);
            IDA_DOC_ATTACHMENT.Fill();
            IGR_DOC_ATTACHMENT.Focus();
        }

        private void BTN_ATT_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        } 

        #region ----- FTP Infomation -----
        //ftp 접속정보 및 환경 정보 설정 
        private void Set_FTP_Info()
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            mFTP_Connect_Status = false;
            mHost = string.Empty;
            mPort = string.Empty;
            mPassive = "N";
            mUserID = string.Empty;
            mPassword = string.Empty;
            mFTP_Folder = string.Empty;
            mClient_Folder = String.Empty;
            try
            {
                IDC_FTP_INFO.SetCommandParamValue("W_FTP_CODE", "SLIP_DOC");
                IDC_FTP_INFO.ExecuteNonQuery();
                if (IDC_FTP_INFO.ExcuteError)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();
                    return;
                }  
                mHost = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_IP"));
                mPort = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_PORT"));
                mUserID = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_NO"));
                mPassword = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_PWD"));
                mPassive = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_PASSIVE_FLAG")); 
                mFTP_Folder = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_FOLDER"));
                mClient_Folder = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_CLIENT_FOLDER"));
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            if (mHost == string.Empty)
            {
                //ftp접속정보 오류          
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            try
            {
                //FileTransfer Initialze
                mFileTransfer = new ISFileTransferAdv();
                mFileTransfer.Host = mHost;
                mFileTransfer.Port = mPort;
                if (mPassive == "Y")
                {
                    mFileTransfer.UsePassive = true; 
                }
                else
                {
                    mFileTransfer.UsePassive = false;
                }
                mFileTransfer.UserId = mUserID;
                mFileTransfer.Password = mPassword;

                mDownload_Folder = string.Format("{0}{1}", mClient_Base_Path, mClient_Folder.Replace("/", "\\"));
            }
            catch (System.Exception Ex)
            {
                //ftp접속정보 오류 
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            //Client Download Folder 없으면 생성 
            System.IO.DirectoryInfo vDownload_Folder = new System.IO.DirectoryInfo(mDownload_Folder);
            if (vDownload_Folder.Exists == false) //있으면 True, 없으면 False
            {
                vDownload_Folder.Create();
            }

            mFTP_Connect_Status = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- File Upload Methods -----
        //ftp에 file upload 처리 
        private bool UpLoadFile(object pSLIP_DATE, object pSLIP_NUM)
        {
            bool isUpload = false;
            OpenFileDialog vOpenFileDialog1 = new OpenFileDialog();
            vOpenFileDialog1.RestoreDirectory = true; 

            if (mFTP_Connect_Status == false)
            {
                isAppInterfaceAdv1.OnAppMessage("FTP Server Connect Fail. Check FTP Server");
                return isUpload;
            }

            if (iString.ISNull(pSLIP_NUM) != string.Empty)
            {
                string vSTATUS = "F";
                string vMESSAGE = string.Empty;

                //openFileDialog1.FileName = string.Format("*{0}", vFileExtension);
                //openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", vFileExtension, vFileExtension);

                vOpenFileDialog1.Title = "Select Open File";
                vOpenFileDialog1.Filter = "All File(*.*)|*.*|pdf File(*.pdf)|*.pdf|jpg file(*.jpg)|*.jpg|bmp file(*.bmp)|*.bmp";
                vOpenFileDialog1.DefaultExt = "*.pdf";
                vOpenFileDialog1.FileName = "";
                vOpenFileDialog1.Multiselect = true;
                 

                if (vOpenFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Application.UseWaitCursor = true;
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();

                    string vSelectFullPath = string.Empty;
                    string vSelectDirectoryPath = string.Empty;

                    string vFileName = string.Empty;
                    string vFileExtension = string.Empty;

                    //1. 사용자 선택 파일 
                    for (int i = 0; i < vOpenFileDialog1.FileNames.Length; i++)
                    {
                        vSelectFullPath = vOpenFileDialog1.FileNames[i];
                        vSelectDirectoryPath = System.IO.Path.GetDirectoryName(vSelectFullPath);

                        vFileName = System.IO.Path.GetFileName(vSelectFullPath);
                        vFileExtension = System.IO.Path.GetExtension(vSelectFullPath).ToUpper(); 

                        //2. 첨부파일 DB 저장 
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_CATEGORY", "SLIP_DOC"); //구분 
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_DATE", pSLIP_DATE);
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_NUM", pSLIP_NUM);
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_USER_FILE_NAME", vFileName);
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_FTP_FILE_NAME", vFileName);
                        IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_EXTENSION_NAME", vFileExtension);
                        IDC_INSERT_DOC_ATTACHMENT.ExecuteNonQuery();

                        vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));
                        object vDOC_ATTACHMENT_ID = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_DOC_ATTACHMENT_ID");
                        object vFTP_FILE_NAME = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_FTP_FILE_NAME");

                        //O_DOC_ATTACHMENT_ID.EditValue = vDOC_ATTACHMENT_ID;
                        //O_FTP_FILE_NAME.EditValue = vFTP_FILE_NAME;

                        if (IDC_INSERT_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            this.Cursor = Cursors.Default;
                            Application.DoEvents();

                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(new Form { TopMost = true }, vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return isUpload;
                        }

                        //3. 첨부파일 로그 저장 
                        IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
                        IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "IN");
                        IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
                        vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
                        if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            this.Cursor = Cursors.Default;
                            Application.DoEvents();
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(new Form { TopMost = true }, vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return isUpload;
                        }

                        //4. 파일 업로드
                        try
                        {
                            mFileTransfer.ShowProgress = true;      //진행바 보이기 

                            //업로드 환경 설정 
                            mFileTransfer.SourceDirectory = vSelectDirectoryPath;
                            mFileTransfer.SourceFileName = vFileName;
                            mFileTransfer.TargetDirectory = mFTP_Folder;
                            mFileTransfer.TargetFileName = iString.ISNull(vFTP_FILE_NAME);

                            bool isUpLoad = mFileTransfer.Upload();

                            if (isUpLoad == true)
                            {
                                isUpload = true;
                            }
                            else
                            {
                                isUpload = false;
                                MessageBoxAdv.Show(new Form { TopMost = true }, isMessageAdapter1.ReturnText("EAPP_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            
                            //5. 적용 
                        }
                        catch (Exception Ex)
                        {
                            isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                            return isUpload;
                        }
                    } 
                }
            }
            return isUpload;
        }

        #endregion;


        #region ----- file Download Methods -----
        //ftp file download 처리 
        private bool DownLoadFile(object pDOC_ATTACHMENT_ID, string pFTP_FileName, string pClient_FileName)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            ////1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            //isDataTransaction1.BeginTran();            
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", pDOC_ATTACHMENT_ID);
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "OUT");
            IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                 
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(new Form { TopMost = true }, vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return IsDownload;
            }

            //2. 실제 다운로드 
            string vTempFileName = string.Format("_{0}", pFTP_FileName);
            try
            {
                System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                if (vDownFileInfo.Exists == true)
                {
                    try
                    {
                        System.IO.File.Delete(vTempFileName);
                    }
                    catch
                    {

                        // ignore
                    }
                }
            }
            catch
            {
                //ignore                        
            } 
            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------
            mFileTransfer.SourceDirectory = mFTP_Folder;
            mFileTransfer.SourceFileName = pFTP_FileName;
            mFileTransfer.TargetDirectory = mDownload_Folder;
            mFileTransfer.TargetFileName = vTempFileName;

            IsDownload = mFileTransfer.Download();

            if (IsDownload == true)
            {
                try
                {
                    //isDataTransaction1.Commit();

                    //다운 파일 FullPath적용 
                    string vTempFullPath = string.Format("{0}\\{1}", mDownload_Folder, vTempFileName);      //임시

                    System.IO.File.Delete(pClient_FileName);                 //기존 파일 삭제 
                    System.IO.File.Move(vTempFullPath, pClient_FileName);    //ftp 이름으로 이름 변경 

                    IsDownload = true;
                }
                catch
                {
                    //isDataTransaction1.RollBack();
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTempFileName);
                            }
                            catch
                            {
                             
                                // ignore
                            }
                        }
                    }
                    catch
                    {
                        //ignore                        
                    }
                }
            }
            else
            {
                //isDataTransaction1.RollBack();
                //download 실패 
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTempFileName);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
                catch
                {
                    //ignore                    
                }
            }
            if (IsDownload == true)
            {  
                System.Diagnostics.Process.Start(pClient_FileName);
            }
            else
            {
                string vMessage = string.Format("{0} {1}", isMessageAdapter1.ReturnText("EAPP_10212"), isMessageAdapter1.ReturnText("QM_10102"));
                MessageBoxAdv.Show(new Form { TopMost = true }, vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return IsDownload;
        }

        #endregion;

        #region ----- file Delete Methods -----
        //ftp file delete 처리 
        private bool DeleteFile(object pDOC_ATTACHMENT_ID, string pFTP_FileName)
        {
            bool IsDelete = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty; 
            
            if (iString.ISNull(pDOC_ATTACHMENT_ID) == string.Empty)
            {
                MessageBoxAdv.Show(new Form { TopMost = true }, isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }
            if (pFTP_FileName == string.Empty)
            {
                MessageBoxAdv.Show(new Form { TopMost = true }, isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
 

            //1. 첨부파일 로그 저장 : Transaction을 이용해서 처리  
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", pDOC_ATTACHMENT_ID);
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "DELETE");
            IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                 
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(new Form { TopMost = true }, vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } 
                return IsDelete;
            }

            //2. 파일 삭제 
            IDC_DELETE_DOC_ATTACHMENT.SetCommandParamValue("W_DOC_ATTACHMENT_ID", pDOC_ATTACHMENT_ID);
            IDC_DELETE_DOC_ATTACHMENT.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));

            if (IDC_DELETE_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
            {
                IsDelete = false; 
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();


                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(new Form { TopMost = true }, vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                } 
                return IsDelete;
            }

            //3. 실제 삭제 
            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransfer.SourceDirectory = mFTP_Folder;
            mFileTransfer.SourceFileName = pFTP_FileName;
            mFileTransfer.TargetDirectory = mFTP_Folder;
            mFileTransfer.TargetFileName = pFTP_FileName;

            IsDelete = mFileTransfer.Delete();
            if (IsDelete == false)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                 
                return IsDelete;
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            return IsDelete;
        }

        #endregion; 
          
        private void IGR_DOC_ATTACHMENT_CellDoubleClick_1(object pSender)
        {
            if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            {
                return;
            }

            object vDOC_ATTACHMENT_ID = IGR_DOC_ATTACHMENT.GetCellValue("DOC_ATTACHMENT_ID");
            string vFTP_FILE_NAME = iString.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            string vUSER_FILE_NAME = string.Format("{0}{1}", mDownload_Folder, IGR_DOC_ATTACHMENT.GetCellValue("USER_FILE_NAME"));
            if (DownLoadFile(vDOC_ATTACHMENT_ID, vFTP_FILE_NAME, vUSER_FILE_NAME) == false)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            }
        }

        #endregion

        private void FCMF0228_Leave(object sender, EventArgs e)
        {

        }
    }
}