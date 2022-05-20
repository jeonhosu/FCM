﻿using System;
using ISCommonUtil;

namespace FCMF0551
{
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        // 쉬트명 정의.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "SOURCE1";
        private string mSourceSheet2 = "SOURCE2";
        
        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // 첫 페이지 체크.
        
        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 10;
        private int mCopy_EndRow = 72;
        private int m1stLastRow = 69;  //첫장 최종 인쇄 라인.
        private int m2ndLastRow = 72;  //첫장이외 최종 인쇄 라인.

        private int mCurrentRow = 6;       //현재 인쇄되는 row 위치.
        private int mDefaultPageRow = 2;    // 페이지 증가후 PageCount 기본값.
         
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

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
            mMessageAdapter = pMessageAdapter;
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

        #region ----- MaxIncrement Methods ----

        private int MaxIncrement(string pPathBase, string pSaveFileName)
        {// 파일명 뒤에 일련번호 증가.
            int vMaxNumber = 0;
            System.IO.DirectoryInfo vFolder = new System.IO.DirectoryInfo(pPathBase);
            string vPattern = string.Format("{0}*", pSaveFileName);
            System.IO.FileInfo[] vFiles = vFolder.GetFiles(vPattern);

            foreach (System.IO.FileInfo vFile in vFiles)
            {
                string vFileNameExt = vFile.Name;
                int vCutStart = vFileNameExt.LastIndexOf(".");
                string vFileName = vFileNameExt.Substring(0, vCutStart);

                int vCutRight = 3;
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

        #region ----- SetArray Grid Index ----

        private void SetArray_Grid(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[9]; 
            // 그리드 or 아답터 위치. 
            pGDColumn[0] = pGrid.GetColumnToIndex("DP_CURRENCY_CODE");
            pGDColumn[1] = pGrid.GetColumnToIndex("DP_ACCOUNT_GROUP_DESC");
            pGDColumn[2] = pGrid.GetColumnToIndex("INCOME_REMARK");
            pGDColumn[3] = pGrid.GetColumnToIndex("INCOME_PLAN");
            pGDColumn[4] = pGrid.GetColumnToIndex("INCOME_RESULT");
            pGDColumn[5] = pGrid.GetColumnToIndex("EXPENSE_REMARK");
            pGDColumn[6] = pGrid.GetColumnToIndex("EXPENSE_PLAN");
            pGDColumn[7] = pGrid.GetColumnToIndex("EXPENSE_RESULT");
            pGDColumn[8] = pGrid.GetColumnToIndex("UNDER_LINE_FLAG"); 
        }

        private void SetArray_Grid_Sum(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[12];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("DP_ACCOUNT_GROUP_DESC");
            pGDColumn[1] = pGrid.GetColumnToIndex("DP_CURRENCY_CODE");
            pGDColumn[2] = pGrid.GetColumnToIndex("BANK_ACCOUNT_NAME");
            pGDColumn[3] = pGrid.GetColumnToIndex("OPEN_BALANCE");
            pGDColumn[4] = pGrid.GetColumnToIndex("INCOME_PLAN");
            pGDColumn[5] = pGrid.GetColumnToIndex("INCOME_RESULT");
            pGDColumn[6] = pGrid.GetColumnToIndex("EXPENSE_PLAN");
            pGDColumn[7] = pGrid.GetColumnToIndex("EXPENSE_RESULT");
            pGDColumn[8] = pGrid.GetColumnToIndex("BALANCE_PLAN");
            pGDColumn[9] = pGrid.GetColumnToIndex("BALANCE_RESULT");
            pGDColumn[10] = pGrid.GetColumnToIndex("UNDER_LINE_FLAG");
            pGDColumn[11] = pGrid.GetColumnToIndex("CELL_MERGE");  
        }

        #endregion;

        #region ----- Array Set 2 ----

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[12];
            pXLColumn = new int[12];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = 0;
            pGDColumn[1] = pGrid.GetColumnToIndex("DOCUMENT_TYPE_DESC");
            pGDColumn[2] = pGrid.GetColumnToIndex("ISSUER_NAME");
            pGDColumn[3] = pGrid.GetColumnToIndex("ISSUE_DATE");
            pGDColumn[4] = pGrid.GetColumnToIndex("SHIPPING_DATE");
            pGDColumn[5] = pGrid.GetColumnToIndex("CURRENCY_CODE");
            pGDColumn[6] = pGrid.GetColumnToIndex("EXCHANGE_RATE");
            pGDColumn[7] = pGrid.GetColumnToIndex("TOTAL_CURR_AMOUNT");
            pGDColumn[8] = pGrid.GetColumnToIndex("TOTAL_BASE_AMOUNT");
            pGDColumn[9] = pGrid.GetColumnToIndex("THIS_CURR_AMOUNT");
            pGDColumn[10] = pGrid.GetColumnToIndex("THIS_BASE_AMOUNT");
            pGDColumn[11] = 0;


            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 2;
            pXLColumn[1] = 4;
            pXLColumn[2] = 8;
            pXLColumn[3] = 11;
            pXLColumn[4] = 15;
            pXLColumn[5] = 19;
            pXLColumn[6] = 21;
            pXLColumn[7] = 25;
            pXLColumn[8] = 31;
            pXLColumn[9] = 37;
            pXLColumn[10] = 43;
            pXLColumn[11] = 49;
        }

        #endregion;

        #region ----- Array Set 2  : Adapter 적용시 ----

        //private void SetArray2(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        //{// 아답터의 table 값.
        //    pGDColumn = new int[10];
        //    pXLColumn = new int[10];

        //    pGDColumn[0] = pTable.Columns.IndexOf("PO_TYPE_NAME");
        //    pGDColumn[1] = pTable.Columns.IndexOf("DISPLAY_NAME");
        //    pGDColumn[2] = pTable.Columns.IndexOf("PO_DATE");
        //    pGDColumn[3] = pTable.Columns.IndexOf("PO_NO");
        //    pGDColumn[4] = pTable.Columns.IndexOf("SUPPLIER_SHORT_NAME");
        //    pGDColumn[5] = pTable.Columns.IndexOf("PRICE_TERM_NAME");
        //    pGDColumn[6] = pTable.Columns.IndexOf("PAYMENT_METHOD_NAME");
        //    pGDColumn[7] = pTable.Columns.IndexOf("PAYMENT_TERM_NAME");
        //    pGDColumn[8] = pTable.Columns.IndexOf("REMARK");
        //    pGDColumn[9] = pTable.Columns.IndexOf("STEP_DESCRIPTION");


        //    pXLColumn[0] = 9;   //PO_TYPE_NAME
        //    pXLColumn[1] = 25;  //DISPLAY_NAME
        //    pXLColumn[2] = 42;  //PO_DATE
        //    pXLColumn[3] = 54;  //PO_NO
        //    pXLColumn[4] = 9;   //SUPPLIER_SHORT_NAME
        //    pXLColumn[5] = 35;  //PRICE_TERM_NAME
        //    pXLColumn[6] = 14;  //PAYMENT_METHOD_NAME
        //    pXLColumn[7] = 41;  //PAYMENT_TERM_NAME
        //    pXLColumn[8] = 7;   //REMARK
        //    pXLColumn[9] = 49;  //금액
        //}

        #endregion;

        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {// 문자열 여부 체크 및 해당 값 리턴.
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

        private bool IsConvertNumber(object pObject, out decimal pConvertDecimal)
        {// 숫자 여부 체크 및 해당 값 리턴.
            bool vIsConvert = false;
            pConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is decimal;
                    if (vIsConvert == true)
                    {
                        decimal vIsConvertNum = (decimal)pObject;
                        pConvertDecimal = vIsConvertNum;
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

        private bool IsConvertDate(object pObject, out System.DateTime pConvertDateTimeShort)
        {// 날짜 여부 체크 및 해당 값 리턴.
            bool vIsConvert = false;
            pConvertDateTimeShort = new System.DateTime();

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is System.DateTime;
                    if (vIsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime;
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

        #endregion;

        #region ----- Excel Write -----

        #region ----- Header Write Method ----

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pPERIOD, object pISSUE_PERIOD, object pWRITE_DATE, object pWRITE_DATE_1)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;
            
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // 신고자 인적사항.
                vXLine = 5;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["CORP_NAME"]);

                vXLine = 5;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUMBER"]);

                vXLine = 6;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["PRESIDENT_NAME"]);

                vXLine = 6;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["ADDRESS"]);
                
                vXLine = 7;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["BUSINESS_TYPE"]);

                vXLine = 7;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["BUSINESS_ITEM"]);

                vXLine = 8;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, pISSUE_PERIOD);

                vXLine = 8;
                vXLColumn = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, pWRITE_DATE);

                //제출일
                vXLine = 29;
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, pWRITE_DATE_1);
                // 제출인
                vXLine = 30;
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["PRESIDENT_NAME"]);

                // 기간.
                vXLine = 3;
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, pPERIOD);

                mPrinting.XLActiveSheet(mSourceSheet2);
                vXLine = 3;
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, pPERIOD);

                vXLine = 4;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["CORP_NAME"]);

                vXLine = 4;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUMBER"]);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        public void HeaderWrite(string pDate)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // 신고자 인적사항.
                vXLine = 3;
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine, vXLColumn, pDate);                 
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Header - SUM Write Method ----

        public void Header_SUM(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        {// 헤더 인쇄.
            int vXLine = 0; //엑셀에 내용이 표시되는 행 번호
            int vXLColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mSourceSheet1);
                
                // 합계 - 건수
                vXLine = 11;
                vXLColumnIndex = 17;
                vObject = pAdapter.CurrentRow["TOTAL_COUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 합계 - 원화금액
                vXLine = 11;
                vXLColumnIndex = 26;
                vObject = pAdapter.CurrentRow["TOTAL_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 내국신용장 - 건수
                vXLine = 12;
                vXLColumnIndex = 17;
                vObject = pAdapter.CurrentRow["LC_COUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 내국신용장 - 원화금액
                vXLine = 12;
                vXLColumnIndex = 26;
                vObject = pAdapter.CurrentRow["LC_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 구매확인서 - 건수
                vXLine = 13;
                vXLColumnIndex = 17;
                vObject = pAdapter.CurrentRow["PC_COUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 구매확인서 - 원화금액
                vXLine = 13;
                vXLColumnIndex = 26;
                vObject = pAdapter.CurrentRow["PC_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Excel Write [Line] Method -----

        private int XLLine_Grid(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.

            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vGDColumnIndex = 0; 

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty; 
            DateTime vCONVERT_DATE = new DateTime();   
            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);
                
                //0 - currency 
                vGDColumnIndex = pGDColumn[0];  
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex); 
                vConvertString = string.Format("{0}", vObject);
                mPrinting.XLSetCell(vXLine, 1, vConvertString);

                //1-item
                vGDColumnIndex = pGDColumn[1];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0}", vObject);
                mPrinting.XLSetCell(vXLine, 2, vConvertString);

                //2-income remark 
                vGDColumnIndex = pGDColumn[2];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0}", vObject);
                mPrinting.XLSetCell(vXLine, 3, vConvertString);
                                
                //3-income plan
                vGDColumnIndex = pGDColumn[3]; 
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 5, vConvertString);

                //4-income result
                vGDColumnIndex = pGDColumn[4];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 6, vConvertString);

                //5-expense remark 
                vGDColumnIndex = pGDColumn[5];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0}", vObject);
                mPrinting.XLSetCell(vXLine, 7, vConvertString);

                //6-expense plan
                vGDColumnIndex = pGDColumn[6];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 9, vConvertString);

                //7-expense result
                vGDColumnIndex = pGDColumn[7];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 10, vConvertString);   
              
                //8-line border position
                vGDColumnIndex = pGDColumn[8];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0}", vObject);
                XLLine_BORDER_LINE(vConvertString, vXLine);
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            pXLine = vXLine; 
            return pXLine;
        }

        private int XLLine_Grid_Sum(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn_Sum)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.

            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vGDColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            DateTime vCONVERT_DATE = new DateTime();
            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                //0 - item 
                vGDColumnIndex = pGDColumn_Sum[0];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0}", vObject);
                mPrinting.XLSetCell(vXLine, 1, vConvertString);

                //1-curr
                vGDColumnIndex = pGDColumn_Sum[1];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0}", vObject);
                mPrinting.XLSetCell(vXLine, 2, vConvertString);

                //2-bank
                vGDColumnIndex = pGDColumn_Sum[2];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0}", vObject);
                mPrinting.XLSetCell(vXLine, 3, vConvertString);

                //3-previous
                vGDColumnIndex = pGDColumn_Sum[3];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 4, vConvertString);

                //4-income plan
                vGDColumnIndex = pGDColumn_Sum[4];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 5, vConvertString);

                //5-income result 
                vGDColumnIndex = pGDColumn_Sum[5];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 6, vConvertString);

                //6-expense plan
                vGDColumnIndex = pGDColumn_Sum[6];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 7, vConvertString);

                //7-expense result
                vGDColumnIndex = pGDColumn_Sum[7];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 8, vConvertString);

                //8-balance plan
                vGDColumnIndex = pGDColumn_Sum[8];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 9, vConvertString);

                //9-balance result
                vGDColumnIndex = pGDColumn_Sum[9];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0:###,###,###,###,###,###.##}", vObject);
                mPrinting.XLSetCell(vXLine, 10, vConvertString);

                //10-line border position
                vGDColumnIndex = pGDColumn_Sum[10];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertString = string.Format("{0}", vObject);
                XLLine_BORDER_LINE(vConvertString, vXLine);

                //11-CELL MERGE
                vGDColumnIndex = pGDColumn_Sum[11];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                int vEND_MERGE_CELL = iString.ISNumtoZero(vObject, 0);
                if (vEND_MERGE_CELL > 1)
                {
                    XLLine_LINE_MERGE(1, vEND_MERGE_CELL, vXLine);
                }
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            pXLine = vXLine;
            return pXLine;
        }

        private int XLLine_BORDER_LINE(string pSTART_BORDER_COL, int pXL_Line)
        {
            int vColumnStart = 1;
            int vColumnEnd = 10;

            mPrinting.XL_LineDraw_Left(pXL_Line, 1, 1, 1);
            for (int c = 1; c <= 10; c++)
            {
                mPrinting.XL_LineDraw_Right(pXL_Line, c, c, 1);
            }

            if (pSTART_BORDER_COL == "C3")
            {
                vColumnStart = 3;
            }
            else if (pSTART_BORDER_COL == "C3")
            {
                vColumnStart = 3;
            }
            else if (pSTART_BORDER_COL == "C2")
            {
                vColumnStart = 2;
            }
            else
            {
                vColumnStart = 1;
            }
            if(pSTART_BORDER_COL == "N")
            {

            }
            else
            {
                mPrinting.XL_LineDraw_Bottom(pXL_Line, vColumnStart, vColumnEnd, 1);
            }
            return pXL_Line;
        }

        private int XLLine_LINE_MERGE(int pSTART_MERGE_CELL, int pEND_MERGE_CELL, int pXL_Line)
        {
            mPrinting.XLCellMerge(pXL_Line, pSTART_MERGE_CELL, pXL_Line, pEND_MERGE_CELL, true);
             
            return pXL_Line;
        }

        #endregion;

        #region ----- TOTAL AMOUNT Write Method -----

        private int XLTOTAL_Line(int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumnIndex = 0;

            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                ////17 - 건수
                //vXLColumnIndex = 17;
                //IsConvert = IsConvertNumber(mTOT_COUNT, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                ////23 - 외화금액
                //vXLColumnIndex = 23;
                //IsConvert = IsConvertNumber(mTOT_CURR_AMOUNT, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0.00}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                ////31 - 원화금액
                //vXLColumnIndex = 31;
                //IsConvert = IsConvertNumber(mTOT_BASE_AMOUNT, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            pXLine = vXLine;

            return pXLine;
        }

        #endregion;

        #region ----- PageNumber Write Method -----

        private void  XLPageNumber(string pActiveSheet, object pPageNumber)
        {// 페이지수를 원본쉬트 복사하기 전에 원본쉬트에 기록하고 쉬트를 복사한다.
            
            int vXLRow = 32; //엑셀에 내용이 표시되는 행 번호
            int vXLCol = 40;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(pActiveSheet);
                mPrinting.XLSetCell(vXLRow, vXLCol, pPageNumber);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte MAIN Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_Sum, string pDate)
        {// 실제 호출되는 부분.
            string vMessage = string.Empty;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);

            int[] vGDColumn;
            int[] vGDColumn_Sum;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            try
            {
                // 실제인쇄되는 행수.
                //int vBy = 35;         
                vTotalRow = pGrid.RowCount;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                //// 총합계.
                //mTOT_COUNT = 0;
                //mTOT_CURR_AMOUNT = 0;
                //mTOT_BASE_AMOUNT = 0;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.

                #region ----- Header Write ----

                HeaderWrite(pDate);  // 헤더 인쇄.

                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                #endregion;

                #region ----- Line Write : Currency ----

                if (vTotalRow > 0)
                {
                    SetArray_Grid(pGrid, out vGDColumn);
                    
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XLLine_Grid(pGrid, vRow, mCurrentRow, vGDColumn); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow - 1)
                        {                            
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow;
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
                }

                #endregion;

                #region ----- Line Write : Item & Total Balance -----
                 
                if (pGrid_Sum.RowCount > 0)
                {
                    mCurrentRow = mCurrentRow + mDefaultPageRow;
                    vPageRowCount = vPageRowCount + mDefaultPageRow;
                    IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                    if (mIsNewPage == true)
                    {
                        mCurrentRow = mCurrentRow + mDefaultPageRow;
                        vPageRowCount = mDefaultPageRow;
                    }
                    else
                    {
                        // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                        mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow);

                        mCurrentRow = mCurrentRow + mDefaultPageRow;
                        vPageRowCount = vPageRowCount + mDefaultPageRow;
                    } 

                    vTotalRow = pGrid_Sum.RowCount;
                    SetArray_Grid_Sum(pGrid_Sum, out vGDColumn_Sum);

                    for (int vRow = 0; vRow < pGrid_Sum.RowCount; vRow++)
                    {
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XLLine_Grid_Sum(pGrid_Sum, vRow, mCurrentRow, vGDColumn_Sum); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow - 1)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow;
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
                }


                #endregion
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            if (mPageNumber == 0)
            {
                mPageNumber = 1;
            }
            return mPageNumber;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount)
        {
            if (mPageNumber == 1 && pPageRowCount >= m1stLastRow)
            { // 첫장 인쇄
                mIsNewPage = true;
                mCurrentRow = mCurrentRow + mCopy_EndRow - m1stLastRow;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow);
            }
            else if (pPageRowCount >= m2ndLastRow)
            { // 첫장 이외
                mIsNewPage = true;
                mCurrentRow = mCurrentRow + mCopy_EndRow - m2ndLastRow;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;

            // page수 표시.
            mPageNumber = mPageNumber + 1;
            XLPageNumber(pActiveSheet, mPageNumber);

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pActiveSheet);            
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol); 
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.

            return vPasteEndRow;


            //int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            //pPrinting.XLActiveSheet("SourceTab1");
            //object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            //pPrinting.XLActiveSheet("Destination");
            //object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            //pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.


            //mPageNumber++; //페이지 번호
            //// 페이지 번호 표시.
            ////string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
            ////int vRowSTART = vCopyPrintingRowEnd - 2;
            ////int vRowEND = vCopyPrintingRowEnd - 2;
            ////int vColumnSTART = 30;
            ////int vColumnEND = 33;
            ////mPrinting.XLCellMerge(vRowSTART, vColumnSTART, vRowEND, vColumnEND, false);
            ////mPrinting.XLSetCell(vRowSTART, vColumnSTART, vPageNumberText); //페이지 번호, XLcell[행, 열]

            //return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            //mPrinting.XLSave(vSaveFileName);

            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
