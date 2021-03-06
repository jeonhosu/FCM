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
using System.IO;
using ISCommonUtil;


namespace EAPF0402
{
    public partial class EAPF0402 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- UpLoad / DownLoad Variables -----

        private InfoSummit.Win.ControlAdv.ISFileTransferAdv mFileTransferAdv;
        private ItemImageInfomationFTP mImageFTP;

        private string mFTP_Source_Directory = string.Empty;            // ftp 소스 디렉토리.
        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 디렉토리.
        private string mClient_Target_Directory = string.Empty;         // 실제 디렉토리 
        private string mClient_ImageDirectory = string.Empty;           // 클라이언트 이미지 디렉토리.
        private string mFileExtension = string.Empty;                   // 확장자명.

        private bool mIsGetInformationFTP = false;                      // FTP 정보 상태.
        private bool mIsFormLoad = false;                               // NEWMOVE 이벤트 제어.

        #endregion;

        #region ----- Constructor -----

        public EAPF0402()
        {
            InitializeComponent();
        }

        public EAPF0402(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mIsFormLoad = false;
        }



        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
          //  IDA_CSR.Fill();
            

            string vCSR_NO = iConvert.ISNull(ISG_CSR_LIST.GetCellValue("CSR_NO"));
            int vIDX_Col = ISG_CSR_LIST.GetColumnToIndex("CSR_NO");

            IDA_CSR.Fill();

            if (ISG_CSR_LIST.RowCount > 0)
            {
                for (int vRow = 0; vRow < ISG_CSR_LIST.RowCount; vRow++)
                {
                    if (vCSR_NO == iConvert.ISNull(ISG_CSR_LIST.GetCellValue(vRow, vIDX_Col)))
                    {
                        ISG_CSR_LIST.CurrentCellActivate(vRow, vIDX_Col);
                        ISG_CSR_LIST.CurrentCellMoveTo(vRow, vIDX_Col);
                    }
                }
            }


        }

        private void Cancel_Col()
        {
            string vCSR_NO = iConvert.ISNull(ISG_CSR_LIST.GetCellValue("CSR_NO"));
            int vIDX_Col = ISG_CSR_LIST.GetColumnToIndex("CSR_NO");

            IDA_CSR.Cancel();

            if (ISG_CSR_LIST.RowCount > 0)
            {
                for (int vRow = 0; vRow < ISG_CSR_LIST.RowCount; vRow++)
                {
                    if (vCSR_NO == iConvert.ISNull(ISG_CSR_LIST.GetCellValue(vRow, vIDX_Col)))
                    {
                        ISG_CSR_LIST.CurrentCellActivate(vRow, vIDX_Col);
                        ISG_CSR_LIST.CurrentCellMoveTo(vRow, vIDX_Col);
                    }
                }
            }
        }

        #endregion;

        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pOutChoice)
        {// pOutChoice : 출력구분.
            //object mTitle = string.Empty;
            object mCORP_NAME = string.Empty;
            object mPERIOD_DATE = string.Empty;
            object mPRINTED_DATE = string.Empty;
            //object mPRINTED_BY = string.Empty;

            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = ISG_CSR_LIST.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                Application.DoEvents();
                return;
            }

            //파일 저장시 파일명 지정.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("전산 시스템 의뢰서{0}", DateTime.Today.ToShortDateString());

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";
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
            }
            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int vPageNumber = 0;
            //int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "EAPF0401_001.xls";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {

                    // 실제 인쇄
                    int vRow = ISG_CSR_LIST.RowIndex;

                    vPageNumber = xlPrinting.ExcelWrite(vRow, ISG_CSR_LIST);

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

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                    isViewItemImage();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_CSR.IsFocused == true)
                    {
                        if (S_RECEIPT_FLAG.CheckedState == ISUtil.Enum.CheckedState.Unchecked)
                        {
                            if (S_COMPLETE_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
                            {
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10044"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                Cancel_Col();
                                return;
                            }
                            S_RECEIPT_DEVELOPER_DESC.EditValue = null;
                            S_COMPLETE_CHARGE_PERSON_LCODE.EditValue = null;
                            S_COMPLETE_DUE_DATE.EditValue = null;
                            S_COMPLETE_CHARGE_PERSON_DESC.EditValue = null;
                            S_COMPLETE_CHARGE_PERSON_LCODE.EditValue = null;

                            S_COMPLETE_DUE_DATE.Updatable = true;
                            S_COMPLETE_CHARGE_PERSON_DESC.Updatable = true;

                            IDA_CSR.Update();
                        }
                        else if (S_COMPLETE_FLAG.CheckedState == ISUtil.Enum.CheckedState.Unchecked)
                        {
                            S_COMPLETE_DEVELOPER_DESC.EditValue = null;
                            S_COMPLETE_DEVELOPER_LCODE.EditValue = null;
                            S_COMPLETE_TYPE_DESC.EditValue = null;
                            S_COMPLETE_TYPE_LCODE.EditValue = null;
                            S_COMPLETE_ASSEMBLY_VERSION.EditValue = null;
                            S_COMPLETE_COMMENT.EditValue = null;

                            S_COMPLETE_DUE_DATE.Updatable = true;
                            S_COMPLETE_CHARGE_PERSON_DESC.Updatable = true;

                            IDA_CSR.Update();
                        }
                        else if (S_COMPLETE_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
                        {
                            IDA_CSR.Update();
                            IDC_MAIL_SEND.ExecuteNonQuery();

                        }
                        else
                        {
                            IDA_CSR.Update();
                        }
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_CSR.IsFocused == true)
                    {
                        IDA_CSR.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Events -----

        private void EAPF0402_Load(object sender, EventArgs e)
        {
            mIsFormLoad = true;


            IDA_CSR.FillSchema();

            //REQ일자
            V_REQ_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            V_REQ_DATE_TO.EditValue = DateTime.Today;
        }

        private void S_RECEIPT_FLAG_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if (S_RECEIPT_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                S_RECEIPT_DATE.EditValue = DateTime.Now;


                if (iConvert.ISNull(S_COMPLETE_CHARGE_PERSON_LCODE.EditValue) == string.Empty)
                {
                    IDC_CSR_DEVELOPER.ExecuteNonQuery();
                    S_RECEIPT_DEVELOPER_DESC.EditValue = IDC_CSR_DEVELOPER.GetCommandParamValue("O_ENTRY_DESCRIPTION");
                    S_RECEIPT_DEVELOPER_LCODE.EditValue = IDC_CSR_DEVELOPER.GetCommandParamValue("O_ENTRY_CODE");
                }
            }
            else
            {
                S_RECEIPT_DATE.EditValue = null;
                S_RECEIPT_DEVELOPER_DESC.EditValue = null;
                S_RECEIPT_DEVELOPER_LCODE.EditValue = null;
            }
        }

        private void S_COMPLETE_FLAG_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if (S_COMPLETE_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                S_COMPLETE_DUE_DATE.Updatable = false;
                S_COMPLETE_CHARGE_PERSON_DESC.Updatable = false;


                S_COMPLETE_DATE.EditValue = DateTime.Now;

                if (iConvert.ISNull(S_COMPLETE_DEVELOPER_LCODE.EditValue) == string.Empty)
                {
                    IDC_CSR_DEVELOPER.ExecuteNonQuery();
                    S_COMPLETE_DEVELOPER_DESC.EditValue = IDC_CSR_DEVELOPER.GetCommandParamValue("O_ENTRY_DESCRIPTION");
                    S_COMPLETE_DEVELOPER_LCODE.EditValue = IDC_CSR_DEVELOPER.GetCommandParamValue("O_ENTRY_CODE");
                }
                
            }
            else
            {
                S_COMPLETE_DUE_DATE.Updatable = true;
                S_COMPLETE_CHARGE_PERSON_DESC.Updatable = true;
                S_COMPLETE_DEVELOPER_DESC.EditValue = null;
                S_COMPLETE_DEVELOPER_LCODE.EditValue = null;
                S_COMPLETE_DATE.EditValue = null;
            }
        }

        private void EAPF0402_Shown(object sender, EventArgs e)
        {
            mIsGetInformationFTP = GetInfomationFTP();
            if (mIsGetInformationFTP == true)
            {
                MakeDirectory();
                FTPInitializtion();
            }
            mIsFormLoad = false;

            //Default CSR_STATUS
            IDC_DEFAULT_CSR_STATUS.SetCommandParamValue("P_LOOKUP_TYPE", "CSR_STATUS");
            IDC_DEFAULT_CSR_STATUS.ExecuteNonQuery();
            V_CSR_STATUS.EditValue = IDC_DEFAULT_CSR_STATUS.GetCommandParamValue("X_ENTRY_DESCRIPTION");
            V_CSR_STATUS_ID.EditValue = IDC_DEFAULT_CSR_STATUS.GetCommandParamValue("X_ENTRY_CODE");

        }

        private void IDA_CSR_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (mIsFormLoad == true)
            {
                return;
            }

            isViewItemImage();

            if (S_COMPLETE_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                S_COMPLETE_DUE_DATE.Updatable = false;
                S_COMPLETE_CHARGE_PERSON_DESC.Updatable = false;
            }
            else
            {
                S_COMPLETE_DUE_DATE.Updatable = true;
                S_COMPLETE_CHARGE_PERSON_DESC.Updatable = true;
            }
        }

        private void ibtSHOW_CSR_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                string vFileName = iConvert.ISNull(ISG_CSR_LIST.GetCellValue("R_ATTACH_FILE_NAME"));

                if (vFileName == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                vFileName = string.Format("{0}\\{1}", mClient_ImageDirectory, vFileName);
                System.Diagnostics.Process.Start(vFileName);
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }

        private void ibtUPLOAD_CSR_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (S_RECEIPT_FLAG.CheckedState == ISUtil.Enum.CheckedState.Unchecked)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10074"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


            if (mIsGetInformationFTP == true)
            {
                UpLoadItem();
            }
        }

        private void ibtSHOW_C_ATTACH_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                string vFileName = iConvert.ISNull(ISG_CSR_LIST.GetCellValue("C_ATTACH_FILE_NAME"));

                if (vFileName == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                vFileName = string.Format("{0}\\{1}", mClient_ImageDirectory, vFileName);
                System.Diagnostics.Process.Start(vFileName);
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }

        #endregion;

        #region ----- is View Item Image Method -----

        private void isViewItemImage()
        {

            string vFileName1 = iConvert.ISNull(ISG_CSR_LIST.GetCellValue("R_ATTACH_FILE_NAME"));
            string vFileName2 = iConvert.ISNull(ISG_CSR_LIST.GetCellValue("C_ATTACH_FILE_NAME"));

            string vDownLoadFile = string.Empty;

            bool isDown1 = DownLoadItem(vFileName1);
            if (isDown1 == true)
            {
                vDownLoadFile = string.Format("{0}\\{1}", mClient_ImageDirectory, vFileName1);
            }

            bool isDown2 = DownLoadItem(vFileName2);
            if (isDown2 == true)
            {
                vDownLoadFile = string.Format("{0}\\{1}", mClient_ImageDirectory, vFileName2);
            }

        }

        #endregion;

        #region -- Default Value Setting --
        private void DefaultValue()
        {
            idcLOCAL_DATE.ExecuteNonQuery();
            S_REQ_DATE.EditValue = idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE");
        }


        #endregion

        #region ----- Make Directory ----

        private void MakeDirectory()
        {
            System.IO.DirectoryInfo vClient_ImageDirectory = new System.IO.DirectoryInfo(mClient_ImageDirectory);
            if (vClient_ImageDirectory.Exists == false) //있으면 True, 없으면 False
            {
                vClient_ImageDirectory.Create();
            }
        }

        #endregion;

        #region ----- Get Information FTP Methods -----

        private bool GetInfomationFTP()
        {
            bool isGet = false;
            try
            {
                idcFTP_INFO.SetCommandParamValue("W_FTP_INFO_CODE", "20");
                idcFTP_INFO.ExecuteNonQuery();
                mImageFTP = new ItemImageInfomationFTP();

                mImageFTP.Host = iConvert.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_IP"));
                mImageFTP.Port = iConvert.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_PORT"));
                mImageFTP.UserID = iConvert.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_USER_ID"));
                mImageFTP.Password = iConvert.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_PASSWORD"));

                mFTP_Source_Directory = iConvert.ISNull(idcFTP_INFO.GetCommandParamValue("O_FTP_SOURCEPATH"));
                mClient_Target_Directory = iConvert.ISNull(idcFTP_INFO.GetCommandParamValue("O_CLIENT_TARGETPATH"));
                //mFileExtension = ".JPG";

                mClient_ImageDirectory = string.Format("{0}\\{1}", mClient_Base_Path, mClient_Target_Directory);

                Application.DoEvents();

                if (mImageFTP.Host != string.Empty)
                {
                    isGet = true;
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return isGet;
        }

        #endregion;

        #region ----- FTP Initialize -----

        private void FTPInitializtion()
        {
            mFileTransferAdv = new ISFileTransferAdv();
            mFileTransferAdv.Host = mImageFTP.Host;
            mFileTransferAdv.Port = mImageFTP.Port;
            mFileTransferAdv.UserId = mImageFTP.UserID;
            mFileTransferAdv.Password = mImageFTP.Password;
        }

        #endregion;

        #region ----- Image Upload Methods -----

        private bool UpLoadItem()
        {
            bool isUp = false;

            if (iConvert.ISNull(S_CSR_NO.EditValue) != "")
            {
                //bool isUp = false;
                string vFileExtension = Path.GetExtension(openFileDialog1.FileName);

                //openFileDialog1.FileName = string.Format("*{0}", vFileExtension);
                //openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", vFileExtension, vFileExtension);
                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|jpg file(*.jpg)|*.jpg|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "";

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string vChoiceFileFullPath = openFileDialog1.FileName;

                        string vChoiceFilePath = vChoiceFileFullPath.Substring(0, vChoiceFileFullPath.LastIndexOf(@"\"));
                        string vChoiceFileName = vChoiceFileFullPath.Substring(vChoiceFileFullPath.LastIndexOf(@"\") + 1);
                        vFileExtension = Path.GetExtension(openFileDialog1.FileName);
                        mFileTransferAdv.ShowProgress = true;
                        //--------------------------------------------------------------------------------

                        string vSourceFileName = vChoiceFileName;

                        string vTargetFileName = S_CSR_NO.EditValue as string;
                        vTargetFileName = string.Format("C_{0}{1}", vTargetFileName.ToUpper(), vFileExtension);

                        mFileTransferAdv.SourceDirectory = vChoiceFilePath;
                        mFileTransferAdv.SourceFileName = vSourceFileName;
                        mFileTransferAdv.TargetDirectory = mFTP_Source_Directory;
                        mFileTransferAdv.TargetFileName = vTargetFileName;

                        S_C_ATTACH_FILE_NAME.EditValue = vTargetFileName;

                        bool isUpLoad = mFileTransferAdv.Upload();

                        if (isUpLoad == true)
                        {
                            isUp = true;
                            ICB_C_ATTACH_YN.CheckedState = ISUtil.Enum.CheckedState.Checked;
                            IDA_CSR.Update();
                            Search_DB();

                            //bool isView = ImageView(vChoiceFileFullPath);
                        }
                        else
                        {
                        }
                    }
                    catch
                    {
                    }
                }
                System.IO.Directory.SetCurrentDirectory(mClient_Base_Path);
                return isUp;
            }
            if (iConvert.ISNull(S_COMPLETE_DATE.EditValue) == "")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10071"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return isUp;
            } 
            else
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10072"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return isUp;
            }
        }

        #endregion;

        #region ----- Image Download Methods -----

        private bool DownLoadItem(string pFileName)
        {

            bool isDown = false;

            string vSourceDownLoadFile = string.Format("{0}\\{1}", mClient_ImageDirectory, pFileName);
            string vTargetDownLoadFile = string.Format("{0}\\_{1}", mClient_ImageDirectory, pFileName);

            string vBeforeSourceFileName = string.Format("{0}", pFileName);
            string vBeforeTargetFileName = string.Format("_{0}", pFileName);

            mFileTransferAdv.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransferAdv.SourceDirectory = mFTP_Source_Directory;
            mFileTransferAdv.SourceFileName = vBeforeSourceFileName;
            mFileTransferAdv.TargetDirectory = mClient_ImageDirectory;
            mFileTransferAdv.TargetFileName = vBeforeTargetFileName;

            isDown = mFileTransferAdv.Download();

            if (isDown == true)
            {
                try
                {
                    System.IO.File.Delete(vSourceDownLoadFile);
                    System.IO.File.Move(vTargetDownLoadFile, vSourceDownLoadFile);

                    isDown = true;
                }
                catch
                {
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTargetDownLoadFile);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTargetDownLoadFile);
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
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTargetDownLoadFile);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTargetDownLoadFile);
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

            return isDown;
        }

        #endregion;  



    }

    #region ----- User Make Class -----

    public class ItemImageInfomationFTP
    {
        #region ----- Variables -----

        private string mHost = string.Empty;
        private string mPort = string.Empty;
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public ItemImageInfomationFTP()
        {
        }

        public ItemImageInfomationFTP(string pHost, string pPort, string pUserID, string pPassword)
        {
            mHost = pHost;
            mPort = pPort;
            mUserID = pUserID;
            mPassword = pPassword;
        }

        #endregion;

        #region ----- Property -----

        public string Host
        {
            get
            {
                return mHost;
            }
            set
            {
                mHost = value;
            }
        }

        public string Port
        {
            get
            {
                return mPort;
            }
            set
            {
                mPort = value;
            }
        }

        public string UserID
        {
            get
            {
                return mUserID;
            }
            set
            {
                mUserID = value;
            }
        }

        public string Password
        {
            get
            {
                return mPassword;
            }
            set
            {
                mPassword = value;
            }
        }

        #endregion;
    }

    #endregion;

}