﻿using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace FCMF0512
{
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        // 쉬트명 정의.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "SourceTab1";
        private string mSourceSheet2 = ""; 

        private int mPageTotalNumber = 0;
        private int mPageNumber = 0;
        
        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 1;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 0;
        private int mCopy_StartRow = 0;
        private int mCopy_EndCol = 0;
        private int mCopy_EndRow = 0;
        private int mPrintingLastRow = 0;       //최종 인쇄 라인.
        private int m2nd_PrintingLastRow = 0;   //2번째장 부터 최종 인쇄 라인.

        private int mCurrentRow = 0;       //현재 인쇄되는 row 위치.
        private int mDefaultPageRow = 5;    // 페이지 증가후 PageCount 기본값.

        private int mCountLinePrinting = 0; //엑셀 라인 Seq 
        private bool m2ndPage = false;

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
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

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface)
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

        private void XlAllContentClear(XL.XLPrint pPrinting)
        {
            object vObject = null;

            pPrinting.XLActiveSheet("SourceTab1");

            //int vStartRow = mPrintingLineSTART1;
            //int vStartCol = mCopyColumnSTART;
            //int vEndRow = mPrintingLineEND1 + 5;
            //int vEndCol = mCopyColumnEND - 1;

            //mPrinting.XLSetCell(vStartRow, vStartCol, vEndRow, vEndCol, vObject);
        
        }

        #endregion;

        #region ----- Line Clear All Methods ----

        private void XlLineClear(int pPrintingLine)
        {
            mPrinting.XLActiveSheet("SourceTab1");

            //int vStartRow = pPrintingLine + 1;
            //int vStartCol = mCopyColumnSTART + 1;
            //int vEndRow = mPrintingLineEND1 - 4;
            //int vEndCol = mCopyColumnEND - 1;

            //if (vStartRow > vEndRow)
            //{
            //    vStartRow = vEndRow; //시작하는 행이 계산후, 끝나는 행 보다 값이 커지므로, 끝나는 행 값을 줌
            //}

            //mPrinting.XL_LineClearInSide(vStartRow, vStartCol, vEndRow, vEndCol);
            //mPrinting.XL_LineClearInSide(vEndRow + 2, vStartCol, vEndRow, vEndCol);
        
        }

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pLOCAL_DATE)
        {
            object vObject = null;
            string vString = string.Empty;
            int vLine = 4;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_CODE"];
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 6, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 12, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 57, vString);

                //인쇄일시[PRINTED DATE]
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vString = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(34, 1, vString);


                ///////////////////////////////////////////////////////////////////////
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                //작성일자[SLIP_DATE]
                vObject = pAdapter.CurrentRow["GL_DATE"];
                if (vObject != null)
                {
                    vString = string.Format("{0:dd-MM-yyyy}", iDate.ISGetDate(vObject));
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 32, vString);

                vLine = vLine + 1;
                //전표번호[GL_NUM]
                vObject = pAdapter.CurrentRow["GL_NUM"];
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);                    
                }
                else
                {
                    vString = string.Empty;               
                }
                mPrinting.XLSetCell(vLine, 32, vString);

                vLine = 8;
                //작성부서명[DEPT_CODE DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_CODE"];
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);  
                }
                else
                {
                    vString = string.Empty;  
                }
                mPrinting.XLSetCell(vLine, 6, vString);

                //작성부서명[DEPT_NAME]
                vObject = pAdapter.CurrentRow["DEPT_NAME"];
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);  
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 12, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pAdapter.CurrentRow["PERSON_NAME"];
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);  
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 57, vString);

                //인쇄일시[PRINTED DATE]
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vString = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(34, 1, vString);
 
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        public void HeaderWrite_1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, object pLOCAL_DATE)
        {
            object vObject = null;
            string vString = string.Empty;
            int vLine = 4;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet2); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("DEPT_CODE"));
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 6, vString);

                //작성부서명[DEPT_NAME]
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("DEPT_NAME"));
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 12, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("PERSON_NAME"));
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 57, vString);

                //인쇄일시[PRINTED DATE]
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vString = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(34, 1, vString);


                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                //작성일자[SLIP_DATE]
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_DATE"));
                if (vObject != null)
                {
                    vString = string.Format("{0:dd-MM-yyyy}", iDate.ISGetDate(vObject));
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 32, vString);

                vLine = vLine + 1;
                //전표번호[GL_NUM]
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_NUM"));
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 32, vString);

                vLine = 8;
                //작성부서명[DEPT_CODE DEPT_NAME]
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("DEPT_CODE"));
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 6, vString);

                //작성부서명[DEPT_NAME]
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("DEPT_NAME"));
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 12, vString);

                //작성자 이름[PERSON_NAME]
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("PERSON_NAME"));
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vLine, 57, vString);

                //인쇄일시[PRINTED DATE]
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vString = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(34, 1, vString);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        public bool HeaderWrite(string pGL_DATE)
        {
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]
                mPrinting.XLSetCell(5, 8, pGL_DATE);
                //mPrinting.XLSetCell(8, 1, "보\r\n통\r\n예\r\n금");
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
                return false;
            }
            return true;
        }

        public bool HeaderWrite_130_1()
        {
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]               
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
                return false;
            }
            return true;
        }

        public bool HeaderWrite_140()
        {
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]                
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
                return false;
            }
            return true;
        }

        public bool HeaderWrite_120()
        {
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
                return false;
            }
            return true;
        }

        public bool HeaderWrite_210_1()
        {
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
                return false;
            }
            return true;
        }

        public bool HeaderWrite_LC()
        {
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
                return false;
            }
            return true;
        }
        
        public bool HeaderWrite_TRX()
        {
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME] 
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
                return false;
            }
            return true;
        }

        #endregion;

        #region ----- Line SLIP Methods ----

        #region ----- Array Set ----

        private void SetArray(out string[] pDBColumn, out int[] pXLColumn)
        {
            pDBColumn = new string[8];
            pXLColumn = new int[8];

            string vDBColumn01 = "ACCOUNT_CODE";
            string vDBColumn02 = "ACCOUNT_DESC";
            string vDBColumn03 = "DR_AMOUNT";
            string vDBColumn04 = "CR_AMOUNT";
            string vDBColumn05 = "M_REFERENCE";
            string vDBColumn06 = "REMARK";
            string vDBColumn07 = "CUSTOMER_DESC";
            string vDBColumn08 = "DEPT_DESC";

            pDBColumn[0] = vDBColumn01;  //ACCOUNT_CODE
            pDBColumn[1] = vDBColumn02;  //ACCOUNT_DESC
            pDBColumn[2] = vDBColumn03;  //DR_AMOUNT
            pDBColumn[3] = vDBColumn04;  //CR_AMOUNT
            pDBColumn[4] = vDBColumn05;  //M_REFERENCE
            pDBColumn[5] = vDBColumn06;  //REMARK
            pDBColumn[6] = vDBColumn07;  //CUSTOMER_DESC
            pDBColumn[7] = vDBColumn08;  //DEPT_DESC

            int vXLColumn01 = 3;         //ACCOUNT_CODE
            int vXLColumn02 = 3;         //ACCOUNT_DESC
            int vXLColumn03 = 12;        //DR_AMOUNT
            int vXLColumn04 = 18;        //CR_AMOUNT
            int vXLColumn05 = 24;        //M_REFERENCE
            int vXLColumn06 = 24;        //REMARK
            int vXLColumn07 = 24;        //CUSTOMER_DESC
            int vXLColumn08 = 40;        //DEPT_DESC

            pXLColumn[0] = vXLColumn01;  //ACCOUNT_CODE
            pXLColumn[1] = vXLColumn02;  //ACCOUNT_DESC
            pXLColumn[2] = vXLColumn03;  //DR_AMOUNT
            pXLColumn[3] = vXLColumn04;  //CR_AMOUNT
            pXLColumn[4] = vXLColumn05;  //M_REFERENCE
            pXLColumn[5] = vXLColumn06;  //REMARK
            pXLColumn[6] = vXLColumn07;  //CUSTOMER_DESC
            pXLColumn[7] = vXLColumn08;  //DEPT_DESC
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
                System.Windows.Forms.Application.DoEvents();
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
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        #endregion;

        #region ----- XlLine Methods -----

        private int XlLine_110(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString= string.Empty;   

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                //[거래은행]
                vObject = pRow["BANK_NAME"];  
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject); 
                }
                else
                {
                    vString = string.Empty; 
                }
                mPrinting.XLSetCell(vXLine, 3, vString);

                //[ACCOUNT_DESC]
                vObject = pRow["BANK_ACCOUNT_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vString);

                //[CURRENCY_CODE]
                vObject = pRow["CURRENCY_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vString);

                //[BEGIN_AMOUNT]
                vObject = pRow["BEGIN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vString);

                //[DR_AMOUNT]
                vObject = pRow["DR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vString);

                //[CR_AMOUNT]
                vObject = pRow["CR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vString);

                //[REMAIN_AMOUNT]
                vObject = pRow["REMAIN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 44, vString);

                //COLOR//
                if (iString.ISNull(pRow["COLOR_FLAG"]) == "C")
                {
                    //mPrinting.XLCellColorBrush(vXLine, 18, 44, System.Drawing.Color.Cornsilk);
                }
                else if (iString.ISNull(pRow["COLOR_FLAG"]) == "S")
                {
                    mPrinting.XLCellColorBrush(vXLine, 3, 44, System.Drawing.Color.Cornsilk);
                }
                else if (iString.ISNull(pRow["COLOR_FLAG"]) == "T")
                {
                    System.Drawing.Color vColor = new System.Drawing.Color();
                    vColor = System.Drawing.Color.FromArgb(216, 228, 188);
                    mPrinting.XLCellColorBrush(vXLine, 3, 44, vColor);
                }

                //-------------------------------------------------------------------
                vXLine= vXLine + 1;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            } 
            pPrintingLine = vXLine;

            return pPrintingLine;
         }
        
        private int XlLine_130_1(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString = string.Empty;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                //[발행처]
                vObject = pRow["BANK_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 3, vString);

                //[수탁및보관처
                vObject = pRow["KEEPER_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vString);

                //[만기월
                vObject = pRow["DUE_MONTH"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vString);

                //[BEGIN_AMOUNT]
                vObject = pRow["BEGIN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vString);

                //[DR_AMOUNT]
                vObject = pRow["DR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vString);

                //[CR_AMOUNT]
                vObject = pRow["CR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vString);

                //[REMAIN_AMOUNT]
                vObject = pRow["REMAIN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 44, vString);

                if (iString.ISNull(pRow["SUMMARY_FLAG"]) == "T")
                {
                    mPrinting.XLCellColorBrush(vXLine, 3, 44, System.Drawing.Color.Cornsilk);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            pPrintingLine = vXLine;

            return pPrintingLine;
        }

        private int XlLine_140(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString = string.Empty;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                //[발행처]
                vObject = pRow["VENDOR_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 3, vString);

                //[결제은행]
                vObject = pRow["BANK_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vString);

                //[만기일
                vObject = pRow["DUE_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vString);

                //[BEGIN_AMOUNT]
                vObject = pRow["BEGIN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vString);

                //[CR_AMOUNT]
                vObject = pRow["CR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vString);

                //[DR_AMOUNT]
                vObject = pRow["DR_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vString);

                //[REMAIN_AMOUNT]
                vObject = pRow["REMAIN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 44, vString);

                if (iString.ISNull(pRow["SUMMARY_FLAG"]) == "T")
                {
                    mPrinting.XLCellColorBrush(vXLine, 3, 44, System.Drawing.Color.Cornsilk);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            pPrintingLine = vXLine;

            return pPrintingLine;
        }

        private int XlLine_120(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString = string.Empty;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                //[거래은행]
                vObject = pRow["BANK_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 3, vString);

                //[계좌번호]
                vObject = pRow["BANK_ACCOUNT_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vString);

                //[금리]
                vObject = pRow["INTER_RATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vString);

                //[금액]
                vObject = pRow["REMAIN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vString);

                //[가입일]
                vObject = pRow["START_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vString);

                //[만기일]
                vObject = pRow["DUE_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vString);

                //[비고]
                vObject = pRow["REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vString);

                if (iString.ISNull(pRow["SUMMARY_FLAG"]) == "T")
                {
                    mPrinting.XLCellColorBrush(vXLine, 3, 44, System.Drawing.Color.Cornsilk);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            pPrintingLine = vXLine;

            return pPrintingLine;
        }

        private int XlLine_210_1(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString = string.Empty;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                //[거래은행]
                vObject = pRow["BANK_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 3, vString);

                //[계좌]
                vObject = pRow["BANK_ACCOUNT_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vString);

                //[금리]
                vObject = pRow["INTER_RATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vString);

                //[대출금액]
                vObject = pRow["REMAIN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vString);

                //[차입일]
                vObject = pRow["ISSUE_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vString);

                //[만기일]
                vObject = pRow["DUE_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vString);

                //[비고]
                vObject = pRow["REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vString);

                if (iString.ISNull(pRow["SUMMARY_FLAG"]) == "T")
                {
                    mPrinting.XLCellColorBrush(vXLine, 3, 44, System.Drawing.Color.Cornsilk);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            pPrintingLine = vXLine;

            return pPrintingLine;
        }

        private int XlLine_LC(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString = string.Empty;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                //[거래은행]
                vObject = pRow["BANK_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 3, vString);

                //[통화]
                vObject = pRow["CURRENCY_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 8, vString);

                //[약정금액]
                vObject = pRow["AMOUNT_LIMIT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 11, vString);

                //[잔여한도]
                vObject = pRow["REMAIN_LIMIT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vString);

                //[이월잔고]
                vObject = pRow["BEGIN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 23, vString);

                //[개설]
                vObject = pRow["OPEN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vString);

                //[만기/할인]
                vObject = pRow["CLOSE_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 35, vString);

                //[당일잔고]
                vObject = pRow["ENDING_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 41, vString);

                //[당월결재]
                vObject = pRow["C_ENDING_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 46, vString);

                if (iString.ISNull(pRow["SUMMARY_FLAG"]) == "T")
                {
                    mPrinting.XLCellColorBrush(vXLine, 3, 50, System.Drawing.Color.Cornsilk);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            pPrintingLine = vXLine;

            return pPrintingLine;
        }

        private int XlLine_TRX(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString = string.Empty;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                //[거래처]
                vObject = pRow["IN_BANK_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 3, vString);

                //[거래내역]
                vObject = pRow["IN_REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 5, vString);

                //[금액]
                vObject = pRow["IN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 22, vString);

                ////[비고]
                //vObject = pRow["IN_DESCRIPTION"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, 23, vString);

                ////[거래처]
                //vObject = pRow["OUT_BANK_NAME"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, 27, vString);

                //[거래내역]
                vObject = pRow["OUT_REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 28, vString);

                //[금액]
                vObject = pRow["OUT_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 45, vString);

                ////[비고]
                //vObject = pRow["OUT_DESCRIPTION"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, 47, vString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            pPrintingLine = vXLine;

            return pPrintingLine;
        }

        private int XlLine_TRX_SUMMARY(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString = string.Empty;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                ////[거래처]
                //vObject = pRow["IN_BANK_NAME"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, 3, vString);

                //[거래내역]
                vObject = pRow["IN_REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 5, vString);

                //[금액]
                vObject = pRow["IN_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 22, vString);

                ////[비고]
                //vObject = pRow["IN_DESCRIPTION"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, 23, vString);

                ////[거래처]
                //vObject = pRow["OUT_BANK_NAME"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, 27, vString);

                //[거래내역]
                vObject = pRow["OUT_REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 28, vString);

                //[금액]
                vObject = pRow["OUT_AMOUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 45, vString);

                ////[비고]
                //vObject = pRow["OUT_DESCRIPTION"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, 47, vString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            pPrintingLine = vXLine;

            return pPrintingLine;
        }

        #endregion;

        #region ----- Sum Write Methods -----

        private void SumWrite(int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber 인쇄//
            //int vPageNumber = 34;
            //int vLINE = 0;
            //for (int r = 1; r <= mPageNumber; r++)
            //{
            //    vLINE = vPageNumber * r;
            //    mPrinting.XLSetCell(vLINE, 29, string.Format("Page {0} of {1}", r, mPageNumber));

            //    if (r == mPageNumber)
            //    {
            //        //
            //    }
            //    else 
            //    {
            //        vLINE = vLINE - 1;
            //        mPrinting.XLSetCell(vLINE, 1, "");
            //    }
            //}

            ////[합계]
            //vLINE = vLINE - 1;
            //mPrinting.XLSetCell(vLINE, 1, "SUM");
            //string vAmount = string.Empty;

            //vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mCURR_DR_AMOUNT);
            //mPrinting.XLSetCell(vLINE, 31, vAmount);

            //vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mDR_AMOUNT);
            //mPrinting.XLSetCell(vLINE, 40, vAmount);

            //vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mCURR_CR_AMOUNT);
            //mPrinting.XLSetCell(vLINE, 49, vAmount);

            //vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mCR_AMOUNT);
            //mPrinting.XLSetCell(vLINE, 58, vAmount);

            //XlLineClear(pPrintingLine);
            
            
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite_110(XLPrinting pXLPrinting, string pGL_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;  

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 51;
            mDefaultPageRow = 6;

            mSourceSheet1 = "SourceTab1";
            mSourceSheet2 = "SourceTab1_1";

            mPrintingLastRow = 44;  //최종 인쇄 라인.
            m2nd_PrintingLastRow = 50;

            mCurrentRow = 7;
            int vPrintingLine = mCurrentRow;
            
            try
            {
                HeaderWrite(pGL_Date);

                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                { 
                    int vCountRow = 0;
                     
                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("보통예금 :: {0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_110(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 1;

                        if (vTotalRow == vCountRow)
                        {
                            mPrinting.XLCellMerge(8, 1, mCurrentRow -1, 2, false);
                            mPrinting.XLSetCell(8, 1, "보\r\n통\r\n예\r\n금");

                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, false);
                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, true);
                            mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCopyLineSUM, mCopy_EndCol);
                            mPrinting.XL_LineDraw_Top(mCurrentRow, mCopy_StartCol, mCopy_EndCol, 2);

                            mPageNumber = iString.ISNumtoZero(Math.Truncate(iString.ISDecimaltoZero(mCopyLineSUM - 1) / 51));
                            mCopyLineSUM = mCurrentRow;
                        }
                        else
                        {
                            IsNewPage_110(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                vPrintingLine = 0;

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
        
        public int LineWrite_111()
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 1;
            mDefaultPageRow = 6;

            mSourceSheet1 = "SourceTab2";
            mSourceSheet2 = "";

            mPrintingLastRow = 50;  //최종 인쇄 라인. 
            int vPrintingLine = mCurrentRow;

            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);
                mCurrentRow = mCurrentRow + 1;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
            mPageNumber = iString.ISNumtoZero(Math.Truncate(iString.ISDecimaltoZero(mCopyLineSUM - 1) / 51));
            return mPageNumber;
        }

        public int LineWrite_130_1(XLPrinting pXLPrinting, string pGL_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 51;
            mDefaultPageRow = 1;

            mSourceSheet1 = "SourceTab3";
            mSourceSheet2 = "";

            mPrintingLastRow = 49;  //최종 인쇄 라인. 
            int vPrintingStartLine = mCurrentRow;
            int vPrintingLine = 1;

            try
            {
                mCurrentRow++;  //실제 인쇄되는 row//
                HeaderWrite_130_1();

                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("전자채권 :: {0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_130_1(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 1;

                        if (vTotalRow == vCountRow)
                        {
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 1, 2, false);
                            mPrinting.XLSetCell(vPrintingStartLine, 1, "전\r\n자\r\n채\r\n권");

                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, false);
                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, true);
                            mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCopyLineSUM, mCopy_EndCol);
                            mPrinting.XL_LineDraw_Top(mCurrentRow, mCopy_StartCol, mCopy_EndCol, 2);

                            mPageNumber = iString.ISNumtoZero(Math.Truncate(iString.ISDecimaltoZero(mCopyLineSUM - 1) / 51));
                            mCopyLineSUM = mCurrentRow;
                        }
                        else
                        {
                            IsNewPage_110(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow;
                                vPrintingLine = mDefaultPageRow;

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

        public int LineWrite_140(XLPrinting pXLPrinting, string pGL_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 51;
            mDefaultPageRow = 1;

            mSourceSheet1 = "SourceTab4";
            mSourceSheet2 = "";

            mPrintingLastRow = 49;  //최종 인쇄 라인. 
            int vPrintingStartLine = mCurrentRow;
            int vPrintingLine = 1;

            try
            {
                mCurrentRow++;  //실제 인쇄되는 row//
                HeaderWrite_140();

                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("전자채무 :: {0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_140(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 1;

                        if (vTotalRow == vCountRow)
                        {
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 1, 2, false);
                            mPrinting.XLSetCell(vPrintingStartLine, 1, "전\r\n자\r\n채\r\n무");

                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, false);
                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, true);
                            mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCopyLineSUM, mCopy_EndCol);
                            mPrinting.XL_LineDraw_Top(mCurrentRow, mCopy_StartCol, mCopy_EndCol, 2);

                            mPageNumber = iString.ISNumtoZero(Math.Truncate(iString.ISDecimaltoZero(mCopyLineSUM - 1) / 51));
                            mCopyLineSUM = mCurrentRow;
                        }
                        else
                        {
                            IsNewPage_110(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow;
                                vPrintingLine = mDefaultPageRow + 1;

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

        public int LineWrite_120(XLPrinting pXLPrinting, string pGL_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 51;
            mDefaultPageRow = 1;

            mSourceSheet1 = "SourceTab5";
            mSourceSheet2 = "";

            mPrintingLastRow = 49;  //최종 인쇄 라인. 
            int vPrintingStartLine = mCurrentRow;
            int vPrintingLine = 1;

            //if (m2ndPage == false)
            //{
            //    mPrintingLastRow = 49;  //최종 인쇄 라인. 
            //    vPrintingStartLine = mCurrentRow;
            //    vPrintingLine = 1;

            //    mCopyLineSUM = (mPageNumber * 51) + 1;
            //    mCurrentRow = mCopyLineSUM;
            //    vPrintingStartLine = mCurrentRow;
            //    mCurrentRow = mCurrentRow + mDefaultPageRow;
            //    vPrintingStartLine = mCurrentRow;
            //    vPrintingLine = mDefaultPageRow;

            //    m2ndPage = true;
            //} 

            vPrintingStartLine = mCurrentRow;
            mCurrentRow = mCurrentRow + mDefaultPageRow;
            vPrintingLine = mDefaultPageRow;

            try
            {
                HeaderWrite_120();

                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("예적금 :: {0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_120(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 1;

                        if (vTotalRow == vCountRow)
                        {
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 1, 2, false);
                            mPrinting.XLSetCell(vPrintingStartLine, 1, "예\r\n적\r\n금");

                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, false);
                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, true);
                            mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCopyLineSUM, mCopy_EndCol);
                            mPrinting.XL_LineDraw_Top(mCurrentRow, mCopy_StartCol, mCopy_EndCol, 2);

                            mPageNumber = iString.ISNumtoZero(Math.Truncate(iString.ISDecimaltoZero(mCopyLineSUM - 1) / 51));
                            mCopyLineSUM = mCurrentRow;
                        }
                        else
                        {
                            IsNewPage_110(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + 2;
                                vPrintingLine = mDefaultPageRow + 1;

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
        
        public int LineWrite_210_1(XLPrinting pXLPrinting, string pGL_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 51;
            mDefaultPageRow = 1;

            mSourceSheet1 = "SourceTab6";
            mSourceSheet2 = "";

            mPrintingLastRow = 49;  //최종 인쇄 라인. 
            int vPrintingStartLine = mCurrentRow;
            int vPrintingLine = 1;
            //if (m2ndPage == false)
            //{

            //    mPrintingLastRow = 49;  //최종 인쇄 라인. 
            //    vPrintingStartLine = mCurrentRow;
            //    vPrintingLine = 1;

            //    mCopyLineSUM = (mPageNumber * 51) + 1;
            //    mCurrentRow = mCopyLineSUM;
            //    vPrintingStartLine = mCurrentRow;
            //    mCurrentRow = mCurrentRow + mDefaultPageRow;
            //    vPrintingStartLine = mCurrentRow;
            //    vPrintingLine = mDefaultPageRow;

            //    m2ndPage = true;
            //} 

            //mCurrentRow = mCurrentRow + mDefaultPageRow;
            vPrintingStartLine = mCurrentRow;
            vPrintingLine = mDefaultPageRow;

            try
            {
                mCurrentRow++;  //실제 인쇄되는 row//
                HeaderWrite_210_1();

                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("차입금 :: {0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_210_1(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 1;

                        if (vTotalRow == vCountRow)
                        {
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 1, 2, false);
                            mPrinting.XLSetCell(vPrintingStartLine, 1, "차\r\n입\r\n금");

                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, false);
                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, true);
                            mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCopyLineSUM, mCopy_EndCol);
                            mPrinting.XL_LineDraw_Top(mCurrentRow, mCopy_StartCol, mCopy_EndCol, 2);

                            mPageNumber = iString.ISNumtoZero(Math.Truncate(iString.ISDecimaltoZero(mCopyLineSUM - 1) / 51));
                            mCopyLineSUM = mCurrentRow;
                        }
                        else
                        {
                            IsNewPage_110(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + 2;
                                vPrintingLine = mDefaultPageRow + 1;

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


        public int LineWrite_LC(XLPrinting pXLPrinting, string pGL_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 51;
            mDefaultPageRow = 1;

            mSourceSheet1 = "SourceTab7";
            mSourceSheet2 = "";

            mPrintingLastRow = 47;  //최종 인쇄 라인. 
            int vPrintingStartLine = mCurrentRow;
            int vPrintingLine = 1;

            //if (m2ndPage == false)
            //{

            //    mPrintingLastRow = 49;  //최종 인쇄 라인. 
            //    vPrintingStartLine = mCurrentRow;
            //    vPrintingLine = 1;

            //    mCopyLineSUM = (mPageNumber * 51) + 1;
            //    mCurrentRow = mCopyLineSUM;
            //    vPrintingStartLine = mCurrentRow;
            //    mCurrentRow = mCurrentRow + mDefaultPageRow;
            //    vPrintingStartLine = mCurrentRow;
            //    vPrintingLine = mDefaultPageRow;

            //    m2ndPage = true;
            //} 

            //mCurrentRow = mCurrentRow + mDefaultPageRow;
            vPrintingStartLine = mCurrentRow;
            vPrintingLine = mDefaultPageRow;

            try
            {
                mCurrentRow++;  //실제 인쇄되는 row//
                HeaderWrite_LC();

                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("L/C :: {0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_LC(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 1;

                        if (vTotalRow == vCountRow)
                        {
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 1, 2, false);
                            mPrinting.XLSetCell(vPrintingStartLine, 1, "L/C");

                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, false);
                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, true);
                            mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCopyLineSUM, mCopy_EndCol);
                            mPrinting.XL_LineDraw_Top(mCurrentRow, mCopy_StartCol, mCopy_EndCol, 2);

                            mPageNumber = iString.ISNumtoZero(Math.Truncate(iString.ISDecimaltoZero(mCopyLineSUM - 1) / 51));
                            mCopyLineSUM = mCurrentRow;
                        }
                        else
                        {
                            IsNewPage_110(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + 2;
                                vPrintingLine = mDefaultPageRow + 1;

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

        public int LineWrite_TRX(XLPrinting pXLPrinting, string pGL_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            string vType = string.Empty;
            string vSUMMARY_FLAG = string.Empty;

            mIsNewPage = false;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;  
            mCopy_EndRow = 51;
            mDefaultPageRow = 3;

            mSourceSheet1 = "SourceTab8";
            mSourceSheet2 = "SourceTab8";

            mPrintingLastRow = 50;  //최종 인쇄 라인. 
            int vPrintingStartLine = mCurrentRow;
            int vPrintingLine = 1;

            //mCopyLineSUM = (mPageNumber * 51) + 1;
            //mCurrentRow = mCopyLineSUM;

            mCurrentRow = mCurrentRow + 1;
            mCopyLineSUM = mCurrentRow;
            vPrintingStartLine = mCurrentRow;
            mCurrentRow = mCurrentRow + mDefaultPageRow;
            vPrintingStartLine = mCurrentRow;
            vPrintingLine = mDefaultPageRow; 

            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;
                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("입/출금 거래 :: {0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_TRX(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 1;
                        vSUMMARY_FLAG = iString.ISNull(vRow["SUMMARY_FLAG"]);

                        if (vSUMMARY_FLAG == "S_CASH_CASH")
                        {                            
                            mPrinting.XL_LineDraw_Bottom(mCurrentRow -1, 3, 4, 1);
                            mPrinting.XL_LineClearRIGHT(mCurrentRow - 1, 4, 4);

                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 5, mCurrentRow - 1, 5, "C");
                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 28, mCurrentRow - 1, 28, "C");
                            mPrinting.XLCellColorBrush(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, System.Drawing.Color.FromArgb(248,244,230));
                        }
                        else if (vSUMMARY_FLAG == "S_CASH_EBILL")
                        {
                            mPrinting.XL_LineDraw_Bottom(mCurrentRow - 1, 3, 4, 1);
                            mPrinting.XL_LineClearRIGHT(mCurrentRow - 1, 4, 4);

                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 5, mCurrentRow - 1, 5, "C");
                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 28, mCurrentRow - 1, 28, "C");
                            mPrinting.XLCellColorBrush(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, System.Drawing.Color.FromArgb(248, 244, 230));
                        }
                        else if (vSUMMARY_FLAG == "S_CASH")
                        {
                            vType = "입\r\n/\r\n출\r\n금\r\n\r\n거\r\n래";
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 2, 2, false);
                            mPrinting.XL_LineDraw_Top(mCurrentRow, 1, 4, 1); 
                            mPrinting.XLSetCell(vPrintingStartLine, 1, vType);

                            for (int r = vPrintingStartLine; r < mCurrentRow - 1; r++)
                            {
                                mPrinting.XL_LineDraw_Right(r, 2, 2, 1);
                            } 

                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 5, mCurrentRow - 1, 5, "C");
                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 28, mCurrentRow - 1, 28, "C");
                            mPrinting.XLCellColorBrush(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, System.Drawing.Color.Cornsilk);

                            vPrintingStartLine = mCurrentRow;               
                        }
                        else if (vSUMMARY_FLAG == "S_MOVE_CASH")
                        {
                            mPrinting.XL_LineDraw_Bottom(mCurrentRow - 1, 3, 4, 1);
                            mPrinting.XL_LineClearRIGHT(mCurrentRow - 1, 4, 4);

                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 5, mCurrentRow - 1, 5, "C");
                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 28, mCurrentRow - 1, 28, "C");
                            mPrinting.XLCellColorBrush(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, System.Drawing.Color.FromArgb(248, 244, 230));
                        }
                        else if (vSUMMARY_FLAG == "S_MOVE_EBILL")
                        {
                            mPrinting.XL_LineDraw_Bottom(mCurrentRow - 1, 3, 4, 1);
                            mPrinting.XL_LineDraw_Bottom(mCurrentRow, 3, 4, 1);
                            mPrinting.XL_LineClearRIGHT(mCurrentRow - 1, 4, 4);

                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 5, mCurrentRow - 1, 5, "C");
                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 28, mCurrentRow - 1, 28, "C");
                            mPrinting.XLCellColorBrush(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, System.Drawing.Color.FromArgb(248, 244, 230));
                        } 
                        //else if (iString.ISNull(vRow["SUMMARY_FLAG"]) == "S_MOVE")
                        else if (vSUMMARY_FLAG == "S_MOVE")
                        {
                            vType = "대\r\n체\r\n거\r\n래";
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 2, 2, false);
                            mPrinting.XL_LineDraw_Top(mCurrentRow, 1, 4, 1);
                            mPrinting.XLSetCell(vPrintingStartLine, 1, vType);

                            for (int r = vPrintingStartLine; r < mCurrentRow - 1; r++)
                            {
                                mPrinting.XL_LineDraw_Right(r, 2, 2, 1);
                            } 

                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 5, mCurrentRow - 1, 5, "C");
                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 28, mCurrentRow - 1, 28, "C");
                            mPrinting.XLCellColorBrush(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, System.Drawing.Color.Cornsilk);
                            vPrintingStartLine = mCurrentRow; 
                        }
                        else if (vSUMMARY_FLAG  == "T")
                        {
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 1, 2, false);

                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 5, mCurrentRow - 1, 5, "C");
                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 28, mCurrentRow - 1, 28, "C");

                            mPrinting.XL_LineDraw_Top(mCurrentRow, mCopy_StartCol, mCopy_EndCol, 1);

                            System.Drawing.Color vColor = new System.Drawing.Color();
                            vColor = System.Drawing.Color.FromArgb(216,228,188);
                            mPrinting.XLCellColorBrush(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, vColor);
                        }

                        if (vTotalRow == vCountRow)
                        {
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 1, 1, false);

                               //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, false);
                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, true);
                            mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCopyLineSUM, mCopy_EndCol);
                            mPrinting.XL_LineDraw_Bottom(mCurrentRow -1, mCopy_StartCol, mCopy_EndCol, 2);
                        }
                        else
                        {
                            IsNewPage_TRX(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 1, 2, false);
                                mCurrentRow = mCurrentRow + mDefaultPageRow + (mCopy_EndRow - mPrintingLastRow);
                                vPrintingStartLine = mCurrentRow;
                                vPrintingLine = mDefaultPageRow;
                                mPrinting.XLSetCell(mCurrentRow, 1, vType);

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


        public int LineWrite_TRX_SUMMARY(XLPrinting pXLPrinting, string pGL_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            string vType = string.Empty;
            string vIN_OUT_TYPE = string.Empty;

            mIsNewPage = false;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 51;
            //mDefaultPageRow = 3;
            mDefaultPageRow = 0;

            mSourceSheet1 = "SourceTab8_1";
            mSourceSheet2 = "SourceTab8_1";

            mPrintingLastRow = 50;  //최종 인쇄 라인. 
            int vPrintingStartLine = mCurrentRow;
            int vPrintingLine = 1;

            mCopyLineSUM = mCurrentRow;  // (mPageNumber * 51) + 1;
            mCurrentRow = mCopyLineSUM;
            vPrintingStartLine = mCurrentRow;
            mCurrentRow = mCurrentRow + mDefaultPageRow;
            vPrintingStartLine = mCurrentRow;
            vPrintingLine = mDefaultPageRow;

            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);

                int vTotalRow = pData.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;
                    foreach (System.Data.DataRow vRow in pData.CurrentRows)
                    {
                        vCountRow++;

                        vMessage = string.Format("입/출금 거래내역 - 집계 :: {0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine_TRX_SUMMARY(vRow, mCurrentRow);
                        vPrintingLine = vPrintingLine + 1; 
                         
                        if (iString.ISNull(vRow["SUMMARY_FLAG"]) == "T")
                        {
                            vType = "집\r\n계";
                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 2, 4, false);
                            mPrinting.XLSetCell(vPrintingStartLine, 1, vType);

                            mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 2, 4, false);

                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 5, mCurrentRow - 1, 5, "C");
                            mPrinting.XLCellAlignmentHorizontal(mCurrentRow - 1, 28, mCurrentRow - 1, 28, "C");

                            mPrinting.XL_LineDraw_Top(mCurrentRow, mCopy_StartCol, mCopy_EndCol, 1);

                            System.Drawing.Color vColor = new System.Drawing.Color();
                            vColor = System.Drawing.Color.FromArgb(216, 228, 188);
                            mPrinting.XLCellColorBrush(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, vColor);
                        }

                        if (vTotalRow == vCountRow)
                        {      
                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, false);
                            //mPrinting.XLCellMerge(mCurrentRow, 1, mCopyLineSUM, mCopy_EndCol, true);
                            mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCopyLineSUM, mCopy_EndCol);
                            mPrinting.XL_LineDraw_Bottom(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, 2);

                            mPageNumber = iString.ISNumtoZero(Math.Ceiling(iString.ISDecimaltoZero(mCurrentRow - 1) / 51));
                            //mCopyLineSUM = mCurrentRow;
                        }
                        else
                        {
                            IsNewPage_TRX(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mPrinting.XLCellMerge(vPrintingStartLine, 1, mCurrentRow - 1, 2, false);
                                mCurrentRow = mCurrentRow + mDefaultPageRow + (mCopy_EndRow - mPrintingLastRow);
                                vPrintingStartLine = mCurrentRow;
                                vPrintingLine = mDefaultPageRow;
                                mPrinting.XLSetCell(mCurrentRow, 1, vType);

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


        public int LineWrite_PL(XLPrinting pXLPrinting, string pGL_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 50;
            mCopy_EndRow = 51;
            mDefaultPageRow = 3;

            mSourceSheet1 = "SourceTab9";
            mSourceSheet2 = "";

            mPrintingLastRow = 50;  //최종 인쇄 라인. 
            int vPrintingStartLine = mCurrentRow;
            int vPrintingLine = 1;

            mPrintingLastRow = 49;  //최종 인쇄 라인. 
            vPrintingStartLine = mCurrentRow;
            vPrintingLine = 1;

            //mCopyLineSUM = (mPageNumber * 51) + 1;
            //mCurrentRow = mCopyLineSUM;

            mCopyLineSUM = mCurrentRow + 1;
            mCurrentRow = mCopyLineSUM;
            vPrintingStartLine = mCurrentRow;
            mCurrentRow = mCurrentRow + mDefaultPageRow;
            vPrintingStartLine = mCurrentRow;
            vPrintingLine = mDefaultPageRow;

            try
            { 
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopyLineSUM);

                mPageNumber = iString.ISNumtoZero(Math.Ceiling(iString.ISDecimaltoZero(mCopyLineSUM -1) / 51));
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

        #region ----- Last Page Number Compute Methods ----

        private void ComputeLastPageNumber(int pTotalRow)
        {
            int vRow = 0;
            mPageTotalNumber = 1;

            if (pTotalRow > 12)
            {
                vRow = pTotalRow - 12;
                mPageTotalNumber = vRow / 18;
                mPageTotalNumber = (vRow % 18) == 0 ? (mPageTotalNumber + 1) : (mPageTotalNumber + 2);
            }
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage_110(int pPrintingLine)
        {
            if (mPageNumber == 1 && m2nd_PrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        private void IsNewPage_TRX(int pPrintingLine)
        {
            if (mPrintingLastRow == pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM);

                //XlAllContentClear(mPrinting);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        private void IsNewPage(int pPrintingLine)
        {
            if (mPrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2,  mCopyLineSUM);

                //XlAllContentClear(mPrinting);
            }
            else
            {
                mIsNewPage = false;
            }
            
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(XL.XLPrint pPrinting, string pSourceTab, int pCopySumPrintingLine)
        {
            mPageNumber++; //페이지 번호

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(pSourceTab); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceTab);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호            
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(vCopySumPrintingLine, mCopy_StartCol, vCopySumPrintingLine + mCopy_EndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            return vCopySumPrintingLine;
        }
         
        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }

        public void PreView(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}