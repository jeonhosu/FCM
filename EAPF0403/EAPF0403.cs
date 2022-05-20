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

namespace EAPF0403
{
    public partial class EAPF0403 : Office2007Form
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

        public EAPF0403(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB_DETAIL(object pCSR_ID)
        {
            if (iConvert.ISNull(pCSR_ID) != string.Empty)
            {
                ITB_CSR.SelectedIndex = 1;
                ITB_CSR.SelectedTab.Focus();

                S_CSR_ID.EditValue = pCSR_ID;

                IDA_CSR_DETAIL.Refillable = true;

                IDA_CSR_DETAIL.Fill();

                if (mIsFormLoad == true)
                {
                    return;
                }

                isViewItemImage();

            }
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    if (ITB_CSR.SelectedIndex == 0)
                    {
                        IDA_CSR.Fill();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {

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
            }
        }

        #endregion;

        #region ----- Form Events -----

        private void EAPF0403_Load(object sender, EventArgs e)
        {
            IDA_CSR.FillSchema();

            //REQ일자
            V_REQ_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            V_REQ_DATE_TO.EditValue = DateTime.Today;

        }

        private void EAPF0403_Shown(object sender, EventArgs e)
        {
            mIsGetInformationFTP = GetInfomationFTP();
            if (mIsGetInformationFTP == true)
            {
                MakeDirectory();
                FTPInitializtion();
            }

            mIsFormLoad = false;
        }


        private void ISG_CSR_LIST_CellDoubleClick(object pSender)
        {
            if (IDA_CSR_DETAIL.Refillable == true)
            {
                if (ISG_CSR_LIST.RowCount > 0)
                {
                    Search_DB_DETAIL(ISG_CSR_LIST.GetCellValue("CSR_ID"));
                }
            }
            else
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("PO_10014"), "WARRING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ibtSHOW_CSR_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                string vFileName = iConvert.ISNull(S_ATTACH_FILE_NAME.EditValue);
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