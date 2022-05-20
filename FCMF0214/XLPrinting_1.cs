using System;
using ISCommonUtil;

namespace FCMF0214
{
    public class XLPrinting_1
    {
        #region ----- Variables -----
        
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;

        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar1;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar2;

        private InfoSummit.Win.ControlAdv.ISGridAdvEx mMessageGrid;

        private XL.XLPrint mPrinting = null;

        // ��Ʈ�� ����.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "Source1";
        private string mSourceSheet2 = "Source2";

        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // ù ������ üũ.

        // �μ�� ���ο� �հ�.
        private int mCopyLineSUM = 0;

        // �μ� 1���� �ִ� �μ�����.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 45;
        private int mCopy_EndRow = 65;

        private int m1stLastRow = 59;       //ù�� ���� �μ� ����.

        private int mCurrentRow = 44;       //���� �μ�Ǵ� row ��ġ.
        private int mDefaultPageRow = 3;    //������ skip�� ����Ǵ� �⺻ PageRowCount �⺻��-������ġ.

        private decimal mDR_AMOUNT = 0; //�����հ�
        private decimal mCR_AMOUNT = 0; //�뺯�հ�

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar1
        {
            set
            {
                mProgressBar1 = value;
            }
        }

        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar2
        {
            set
            {
                mProgressBar2 = value;
            }
        }

        public InfoSummit.Win.ControlAdv.ISGridAdvEx MessageGridEx
        {
            set
            {
                mMessageGrid = value;
            }
        }

        public string OpenFileNameExcel
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting_1(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
        }

        #endregion;

        #region ----- XL File Open -----

        public bool XLFileOpen()
        {
            bool IsOpen = false;

            try
            {
                IsOpen = mPrinting.XLOpenFile(mXLOpenFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }

            return IsOpen;
        }

        #endregion;

        #region ----- Dispose -----

        public void Dispose()
        {
            mPrinting.XLOpenFileClose();
            mPrinting.XLClose();
        }

        #endregion;

        #region ----- MessageGrid Methods ----

        private void MessageGrid(string pMessage)
        {
            int vCountRow = mMessageGrid.RowCount;
            vCountRow = vCountRow + 1;
            mMessageGrid.RowCount = vCountRow;

            int vCurrentRow = vCountRow - 1;

            mMessageGrid.SetCellValue(vCurrentRow, 0, pMessage);

            mMessageGrid.CurrentCellMoveTo(vCurrentRow, 0);
            mMessageGrid.Focus();
            mMessageGrid.CurrentCellActivate(vCurrentRow, 0);
        }

        #endregion;

        #region ----- Convert DateTime Methods ----

        private object ConvertDateTime(object pObject)
        {
            object vObject = null;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vObject;
        }

        private object ConvertDate(object pObject)
        {
            object vObject = null;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vObject;
        }

        #endregion;

        #region ----- MaxIncrement Methods ----

        private int MaxIncrement(string pPathBase, string pSaveFileName)
        {
            int vMaxNumber = 0;
            System.IO.DirectoryInfo vFolder = new System.IO.DirectoryInfo(pPathBase);
            string vPattern = string.Format("{0}*", pSaveFileName);
            System.IO.FileInfo[] vFiles = vFolder.GetFiles(vPattern);

            foreach (System.IO.FileInfo vFile in vFiles)
            {
                string vFileNameExt = vFile.Name;
                int vCutStart = vFileNameExt.LastIndexOf(".");
                string vFileName = vFileNameExt.Substring(0, vCutStart);

                int vCutRight = 2;
                int vSkip = vFileName.Length - vCutRight;
                string vTextNumber = vFileName.Substring(vSkip, vCutRight);
                int vNumber = int.Parse(vTextNumber);

                if (vNumber > vMaxNumber)
                {
                    vMaxNumber = vNumber;
                }
            }

            return vMaxNumber;
        }

        #endregion;

        #region ----- Content Clear All Methods ----

        //private void XlAllContentClear(XL.XLPrint pPrinting)
        //{
        //    int vXLColumn01 = 2;
        //    int vXLColumn02 = 12;
        //    int vXLColumn03 = 18;
        //    int vXLColumn04 = 24;

        //    object vObject = null;
        //    int vPrintingLineMAX = mPrintingLineMAX + 2;

        //    pPrinting.XLActiveSheet("SourceTab1");

        //    for (int vXLine = mPositionPrintLineSTART; vXLine < vPrintingLineMAX; vXLine++)
        //    {
        //        pPrinting.XLSetCell(vXLine, vXLColumn01, vObject);
        //        pPrinting.XLSetCell(vXLine, vXLColumn02, vObject);
        //        pPrinting.XLSetCell(vXLine, vXLColumn03, vObject);
        //        pPrinting.XLSetCell(vXLine, vXLColumn04, vObject);
        //    }
        //}

        #endregion;

        #region ----- Line Clear All Methods ----

        //private void XlLineClear(int pPrintingLine)
        //{
        //    int vPrintHeaderColumnSTART = 1; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
        //    int vPrintHeaderColumnEND = 46;  //����Ǿ��� ��Ʈ�� ��, ���῭

        //    int vPrintingLineMAX = mPrintingLineMAX - 1;

        //    mPrinting.XLActiveSheet("SourceTab1");

        //    for (int vXLine = pPrintingLine; vXLine < vPrintingLineMAX; vXLine++)
        //    {
        //        mPrinting.XL_LineClear(vXLine, vPrintHeaderColumnSTART, vPrintHeaderColumnEND);
        //    }
        //}

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        {
            object vObject = null;

            try
            {
                string vColumnName_01 = "SLIP_NUM";           //��ǥ��ȣ[SLIP_NUM]
                string vColumnName_02 = "SLIP_DATE";          //��������[SLIP_DATE]
                string vColumnName_03 = "DEPT_NAME";          //���Ǻμ�[DEPT_NAME]
                string vColumnName_04 = "PERSON_NAME";        //������[PERSON_NAME]
                string vColumnName_06 = "SUB_REMARK";         //�ݾ�[SUB_REMARK]
                string vColumnName_07 = "REMARK";             //����[REMARK]
                string vColumnName_08 = "SUBSTANCE";          //����[SUBSTANCE]

                System.Drawing.Point vCellPoint01 = new System.Drawing.Point(7, 2);    //��ǥ��ȣ[SLIP_NUM]
                System.Drawing.Point vCellPoint02 = new System.Drawing.Point(9, 7);    //��������[SLIP_DATE]
                System.Drawing.Point vCellPoint03 = new System.Drawing.Point(11, 7);   //���Ǻμ�[DEPT_NAME]
                System.Drawing.Point vCellPoint04 = new System.Drawing.Point(13, 7);   //������[PERSON_NAME]
                System.Drawing.Point vCellPoint05 = new System.Drawing.Point(15, 7);   //����[�����]
                System.Drawing.Point vCellPoint06 = new System.Drawing.Point(17, 7);   //�ݾ�[SUB_REMARK]
                System.Drawing.Point vCellPoint07 = new System.Drawing.Point(19, 2);   //����[REMARK]
                System.Drawing.Point vCellPoint08 = new System.Drawing.Point(23, 2);   //����[SUBSTANCE]

                mPrinting.XLActiveSheet("SourceTab1"); //���� ���ڸ� �ֱ� ���� ��Ʈ ����

                //��ǥ��ȣ[SLIP_NUM]
                vObject = pAdapter.CurrentRow[vColumnName_01];
                if (vObject != System.DBNull.Value)
                {
                    vObject = string.Format("���ǹ�ȣ : {0}", vObject);
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vObject);
                }

                //��������[SLIP_DATE]
                vObject = pAdapter.CurrentRow[vColumnName_02];
                if (vObject != System.DBNull.Value)
                {
                    vObject = ConvertDate(vObject);
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, vObject);
                }

                //���Ǻμ�[DEPT_NAME]
                vObject = pAdapter.CurrentRow[vColumnName_03];
                if (vObject != System.DBNull.Value)
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, vObject);
                }

                //������[PERSON_NAME]
                vObject = pAdapter.CurrentRow[vColumnName_04];
                if (vObject != System.DBNull.Value)
                {
                    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, vObject);
                }

                //����[�����]
                vObject = "�����";
                mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, vObject);

                //�ݾ�[SUB_REMARK]
                vObject = pAdapter.CurrentRow[vColumnName_06];
                if (vObject != System.DBNull.Value)
                {
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, vObject);
                }

                //����[REMARK]
                vObject = pAdapter.CurrentRow[vColumnName_07];
                if (vObject != System.DBNull.Value)
                {
                    vObject = string.Format("��  �� : {0}", vObject);
                    mPrinting.XLSetCell(vCellPoint07.X, vCellPoint07.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint07.X, vCellPoint07.Y, vObject);
                }

                //����[SUBSTANCE]
                string vContent = string.Empty;
                vObject = pAdapter.CurrentRow[vColumnName_08];
                if (vObject != System.DBNull.Value)
                {
                    bool isConvert = vObject is string;
                    if (isConvert == true)
                    {
                        vContent = vObject as string;
                        bool isNull = string.IsNullOrEmpty(vContent.Trim());
                        if (isNull != true)
                        {
                            byte b_CR_Character = 0x0d; //CR
                            byte b_SP_Character = 0x20; //Space
                            char vCharOld = (char)b_CR_Character;
                            char vCharNew = (char)b_SP_Character;
                            vContent = vContent.Replace(vCharOld, vCharNew);
                        }
                    }
                    vObject = vContent;
                    mPrinting.XLSetCell(vCellPoint08.X, vCellPoint08.Y, vObject);
                }
                else
                {
                    vObject = null;
                    mPrinting.XLSetCell(vCellPoint08.X, vCellPoint08.Y, vObject);
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;

        #region ----- Line SLIP Methods ----

        #region ----- Array Set ----

        private void SetArray(out string[] pDBColumn, out int[] pXLColumn)
        {
            pDBColumn = new string[6];
            pXLColumn = new int[6];

            string vDBColumn01 = "ACCOUNT_CODE";
            string vDBColumn02 = "ACCOUNT_DESC";
            string vDBColumn03 = "DR_AMOUNT";
            string vDBColumn04 = "CR_AMOUNT";
            string vDBColumn05 = "M_REFERENCE";
            string vDBColumn06 = "REMARK";

            pDBColumn[0] = vDBColumn01;  //ACCOUNT_CODE
            pDBColumn[1] = vDBColumn02;  //ACCOUNT_DESC
            pDBColumn[2] = vDBColumn03;  //DR_AMOUNT
            pDBColumn[3] = vDBColumn04;  //CR_AMOUNT
            pDBColumn[4] = vDBColumn05;  //M_REFERENCE
            pDBColumn[5] = vDBColumn06;  //REMARK

            int vXLColumn01 = 2;         //ACCOUNT_CODE
            int vXLColumn02 = 2;         //ACCOUNT_DESC
            int vXLColumn03 = 14;        //DR_AMOUNT
            int vXLColumn04 = 20;        //CR_AMOUNT
            int vXLColumn05 = 26;        //M_REFERENCE
            int vXLColumn06 = 26;        //REMARK

            pXLColumn[0] = vXLColumn01;  //ACCOUNT_CODE
            pXLColumn[1] = vXLColumn02;  //ACCOUNT_DESC
            pXLColumn[2] = vXLColumn03;  //DR_AMOUNT
            pXLColumn[3] = vXLColumn04;  //CR_AMOUNT
            pXLColumn[4] = vXLColumn05;  //M_REFERENCE
            pXLColumn[5] = vXLColumn06;  //REMARK
        }

        #endregion;

        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {
            bool vIsConvert = false;
            pConvertString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is string;
                    if (vIsConvert == true)
                    {
                        pConvertString = pObject as string;
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }

            return vIsConvert;
        }

        private bool IsConvertNumber(string pStringNumber, out decimal pConvertDecimal)
        {
            bool vIsConvert = false;
            pConvertDecimal = 0m;

            try
            {
                if (pStringNumber != null)
                {
                    decimal vIsConvertNum = decimal.Parse(pStringNumber);
                    pConvertDecimal = vIsConvertNum;
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }

            return vIsConvert;
        }

        #endregion;

        #region ----- XlLine Methods -----

        private int XlLine(System.Data.DataRow pRow, int pPrintingLine, string[] pDBColumn, int[] pXLColumn)
        {
            int vXLine = pPrintingLine; //������ ������ ǥ�õǴ� �� ��ȣ

            string vColumnName1= string.Empty;
            int vXLColumnIndex = 0;

            string vConvertString1 = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert1 = false;

            try
            {
                //[ACCOUNT_CODE]
                vColumnName1 = pDBColumn[0];
                vXLColumnIndex = pXLColumn[0];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[DR_AMOUNT]
                vColumnName1 = pDBColumn[2];
                vXLColumnIndex = pXLColumn[2];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);

                    vConvertString1 = vConvertString1.Replace(",", "");
                    IsConvertNumber(vConvertString1, out vConvertDecimal);
                    mDR_AMOUNT = mDR_AMOUNT + vConvertDecimal;
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[CR_AMOUNT]
                vColumnName1 = pDBColumn[3];
                vXLColumnIndex = pXLColumn[3];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);

                    vConvertString1 = vConvertString1.Replace(",", "");
                    IsConvertNumber(vConvertString1, out vConvertDecimal);
                    mCR_AMOUNT = mCR_AMOUNT + vConvertDecimal;
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[M_REFERENCE]
                vColumnName1 = pDBColumn[4];
                vXLColumnIndex = pXLColumn[4];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------

                //[ACCOUNT_DESC]
                vColumnName1 = pDBColumn[1];
                vXLColumnIndex = pXLColumn[1];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------

                //[REMARK]
                vColumnName1 = pDBColumn[5];
                vXLColumnIndex = pXLColumn[5];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }


            pPrintingLine = vXLine;
            //IsNewPage(pPrintingLine);
            //if (mIsNewPage == true)
            //{
            //    pPrintingLine = mPositionPrintLineSTART;
            //}

            return pPrintingLine;
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            string[] vDBColumn;
            int[] vXLColumn;

            mDR_AMOUNT = 0;
            mCR_AMOUNT = 0;

            int vPrintingLine = mPositionPrintLineSTART;

            try
            {
                int vTotalRow = pData.OraSelectData.Rows.Count;
                if (vTotalRow > 0)
                {
                    mPageTotalNumber = vTotalRow / 5;
                    mPageTotalNumber = (vTotalRow % 5) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);

                    int vCountRow = 0;

                    SetArray(out vDBColumn, out vXLColumn);

                    foreach (System.Data.DataRow vRow in pData.OraSelectData.Rows)
                    {
                        vCountRow++;

                        vPrintingLine = XlLine(vRow, vPrintingLine, vDBColumn, vXLColumn);

                        if (vTotalRow == vCountRow)
                        {
                            if (mPositionPrintLineSTART != vPrintingLine) //66������ 1������ ��¹����� 2������ �غ� ������ �̸� ǥ�õ� ��Ʈ�� ���� Skip �ǵ��� �ϱ����� ��
                            {
                                //[��]
                                //mPrinting.XLSetCell(65, 2, "��");
                                string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
                                mPrinting.XLSetCell(60, 14, vDRAmount);
                                string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
                                mPrinting.XLSetCell(60, 20, vCRAmount);

                                XlLineClear(vPrintingLine);
                            }
                            else
                            {
                                //[��]
                                //mPrinting.XLSetCell(65, 2, "��");
                                string vDRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
                                mPrinting.XLSetCell(60, 14, vDRAmount);
                                string vCRAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
                                mPrinting.XLSetCell(60, 20, vCRAmount);
                            }

                            mCopySumPrintingLine = CopyAndPaste(mPrinting, mCopySumPrintingLine);
                            XlAllContentClear(mPrinting);
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                vPrintingLine = mPositionPrintLineSTART;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine)
        {
            if (mPrintingLineMAX < pPrintingLine)
            {
                mIsNewPage = true;
                mCopySumPrintingLine = CopyAndPaste(mPrinting, mCopySumPrintingLine);

                XlAllContentClear(mPrinting);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            int vPrintHeaderColumnSTART = 1; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
            int vPrintHeaderColumnEND = 46;  //����Ǿ��� ��Ʈ�� ��, ���῭

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPageNumber++; //������ ��ȣ
            //string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
            //mPrinting.XLActiveSheet("SourceTab1"); //�� �Լ��� ȣ�� ���� ������ �׸������� XL Sheet�� Insert ���� �ʴ´�.
            //mPrinting.XLSetCell(66, 21, vPageNumberText); //������ ��ȣ, XLcell[��, ��]

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            pPrinting.XLActiveSheet("SourceTab1");
            object vRangeSource = pPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mIncrementCopyMAX, vPrintHeaderColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet("Destination");
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            //mPrinting.XLPrinting(pPageSTART, pPageEND);
            mPrinting.XLPrintPreview();
        }

        public void PreView(int pPageStart, int pPageEnd)
        {
            mPrinting.XLPreviewPrinting(pPageStart, pPageEnd, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;
    }
}