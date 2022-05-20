using System;
using ISCommonUtil;

namespace FCMF0580
{
    public class XL_Upload
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private string mMessageError = string.Empty;

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;
        
        private XL.XLPrint mExcel_Upload = null;

        private string mXLOpenFileName = string.Empty;

        private int mTotalROW = 0;    //Excel Active Sheet Row Count
        private int mTotalCOLUMN = 0; //Excel Active Sheet Column Count

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        public string OpenFileName
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        public int TotalROW
        {
            get
            {
                return mTotalROW;
            }
            set
            {
                mTotalROW = value;
            }
        }

        public int TotalCOLUMN
        {
            get
            {
                return mTotalCOLUMN;
            }
            set
            {
                mTotalCOLUMN = value;
            }
        }

        //public int ReadRow
        //{
        //    get
        //    {
        //        return mStartRowRead;
        //    }
        //    set
        //    {
        //        mStartRowRead = value;
        //    }
        //}

        #endregion;

        #region ----- Constructor -----

        public XL_Upload()
        {
            mExcel_Upload = new XL.XLPrint();
        }

        public XL_Upload(InfoSummit.Win.ControlAdv.ISAppInterfaceAdv pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mAppInterface = pAppInterface;
            mMessageAdapter = pMessageAdapter;

            mExcel_Upload = new XL.XLPrint();
        }

        #endregion;

        #region ----- XLDispose -----

        public void DisposeXL()
        {
            mExcel_Upload.XLOpenFileClose();
            mExcel_Upload.XLClose();
        }

        #endregion;

        #region ----- XL File Open -----

        public bool OpenXL()
        {
            bool IsOpen = false;

            try
            {
                IsOpen = mExcel_Upload.XLFileOpen(mXLOpenFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }

            return IsOpen;
        }

        #endregion;

        #region ----- Convert String Methods ----

        private string ConvertString(object pObject)
        {
            string vString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is string;
                    if (IsConvert == true)
                    {
                        vString = pObject as string;
                    }
                }
            }
            catch
            {
            }

            return vString;
        }

        #endregion;

        #region ----- Convert Date Methods ----

        private System.DateTime ConvertDate(object pObject)
        {
            bool isConvert = false;
            string vTextDateTimeShort = string.Empty;
            System.DateTime vDate = DateTime.Today;

            try
            {
                if (pObject != null)
                {
                    isConvert = pObject is double;
                    if (isConvert == true)
                    {
                        double isConvertDouble = (double)pObject;
                        vDate = System.DateTime.FromOADate(isConvertDouble);
                    }
                    else if (iDate.ISDate(pObject) == true)
                    {
                        vDate = iDate.ISGetDate(pObject);
                    }
                    else
                    {
                        vDate = iDate.ISGetDate("-");
                    }
                }
            }
            catch
            {
                vDate = iDate.ISGetDate("-");
            }
            return vDate;
        }

        #endregion;

        #region ----- Convert Decimal Methods ----

        private decimal ConvertDecimal(object pObject)
        {
            bool isConvert = false;
            decimal vConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    isConvert = pObject is decimal;
                    if (isConvert == true)
                    {
                        decimal isConvertNum = (decimal)pObject;
                        vConvertDecimal = isConvertNum;
                    }
                }

            }
            catch
            {

            }
            return vConvertDecimal;
        }

        #endregion;

        #region ----- Convert Double Methods ----

        private decimal ConvertDouble(object pObject)
        {
            bool isConvert = false;
            decimal vConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    isConvert = pObject is double;
                    if (isConvert == true)
                    {
                        double isConvertDouble = (double)pObject;
                        vConvertDecimal = Convert.ToDecimal(isConvertDouble);
                    }
                }
            }
            catch
            {
            }

            return vConvertDecimal;
        }

        #endregion;

        #region ----- XL Loading -----

        public bool LoadXL(InfoSummit.Win.ControlAdv.ISDataCommand pCMD, int pStartRow)
        {
            string vMessage = string.Empty;

            
            mExcel_Upload.XLActiveSheet(1);
            mTotalROW = mExcel_Upload.CountROW + 1;

            bool isLoad = false;
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            DateTime vConvertDate = new DateTime();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            string vLOAN_NUM = string.Empty;

            int vADRow = 0;  
            try
            {
                for (int vRow = pStartRow; vRow < mTotalROW; vRow++)
                {
                    //KEY값에 해당하는 셀에 DATA가 있을 경우만 INSERT를 처리해야 하므로//
                    
                    vObject = mExcel_Upload.XLGetCell(vRow, 1);  //차입금번호.
                    vLOAN_NUM = iString.ISNull(vObject);

                    if (vLOAN_NUM != string.Empty)  //사원번호가 있을 경우만 처리.
                    {
                        pCMD.SetCommandParamValue("P_LOAN_NUM", vLOAN_NUM);

                        vObject = mExcel_Upload.XLGetCell(vRow, 2);
                        vConvertString = iString.ISNull(vObject);
                        vConvertString = vConvertString.Trim();
                        pCMD.SetCommandParamValue("P_LIMIT_LOAN_CODE", vConvertString);


                        vObject = mExcel_Upload.XLGetCell(vRow, 3);
                        vConvertString = iString.ISNull(vObject);
                        vConvertString = vConvertString.Trim();
                        pCMD.SetCommandParamValue("P_EXEC_NUM", vConvertString);
                       

                        vObject = mExcel_Upload.XLGetCell(vRow, 4);  
                        vConvertDate = ConvertDate(vObject); 
                        pCMD.SetCommandParamValue("P_PLAN_DATE", vConvertDate);

                        vObject = mExcel_Upload.XLGetCell(vRow, 5);
                        vConvertString = iString.ISNull(vObject);
                        vConvertString = vConvertString.Trim();
                        pCMD.SetCommandParamValue("P_LOAN_PLAN_TYPE", vConvertString);

                        vObject = mExcel_Upload.XLGetCell(vRow, 6);
                        vConvertString = iString.ISNull(vObject);
                        vConvertString = vConvertString.Trim();
                        pCMD.SetCommandParamValue("P_CURRENCY_CODE", vConvertString);

                        vObject = mExcel_Upload.XLGetCell(vRow, 7);
                        vConvertDecimal = iString.ISDecimaltoZero(vObject);
                        pCMD.SetCommandParamValue("P_AMOUNT", vConvertDecimal);

                        vObject = mExcel_Upload.XLGetCell(vRow, 8);
                        vConvertDate = ConvertDate(vObject);
                        pCMD.SetCommandParamValue("P_PLAN_DATE_FR", vConvertDate);

                        vObject = mExcel_Upload.XLGetCell(vRow, 9);
                        vConvertDate = ConvertDate(vObject);
                        pCMD.SetCommandParamValue("P_PLAN_DATE_TO", vConvertDate);

                        vObject = mExcel_Upload.XLGetCell(vRow, 10);
                        vConvertDecimal = iString.ISDecimaltoZero(vObject);
                        pCMD.SetCommandParamValue("P_INTEREST_RATE", vConvertDecimal);

                        vObject = mExcel_Upload.XLGetCell(vRow, 11);
                        vConvertDecimal = iString.ISDecimaltoZero(vObject);
                        pCMD.SetCommandParamValue("P_SPREAD_RATE", vConvertDecimal);
                        pCMD.ExecuteNonQuery();
                        vSTATUS = iString.ISNull(pCMD.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iString.ISNull(pCMD.GetCommandParamValue("O_MESSAGE"));
                        if(vSTATUS == "F")
                        {
                            isLoad = false;
                            if(vMESSAGE != string.Empty)
                            {
                                vMessage = string.Format("Excel Uploading Error : {0}", vMESSAGE);
                                mAppInterface.OnAppMessage(vMessage);
                                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                                System.Windows.Forms.Application.DoEvents();
                                return isLoad;
                            }
                        }

                    }
                    vADRow++; 

                    vMessage = string.Format("Excel Uploading : {0:D4}/{1:D4}", vRow, (mTotalROW - 1));
                    mAppInterface.OnAppMessage(vMessage);
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();
                }
                isLoad = true;
            }
            catch (System.Exception ex)
            {
                DisposeXL();

                mAppInterface.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return isLoad;
        }

        #endregion;
    }
}
