using System;
using ISCommonUtil;

namespace FCMF0519
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
        private string mSourceSheet1 = "Source1";
        private string mSourceSheet2 = "Source2";
        
        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // 첫 페이지 체크.
        
        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 0;
        private int mCopy_StartRow = 0;
        private int mCopy_EndCol = 0;
        private int mCopy_EndRow = 0;
        private int mPrintingLastRow = 0;  //최종 인쇄 라인.

        private int mCurrentRow = 0;       //현재 인쇄되는 row 위치.
        private int mDefaultPageRow = 0;    // 페이지 증가후 PageCount 기본값.

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

        #region ----- Line SLIP Methods ----

        #region ----- Array Set 1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDCol)
        {
            // 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDCol = new int[10];

            // 그리드 or 아답터 위치.
            pGDCol[0] = pGrid.GetColumnToIndex("ACCOUNT_DR_CR_NAME");       //구분
            pGDCol[1] = pGrid.GetColumnToIndex("GL_DATE");                  //전표일자
            pGDCol[2] = pGrid.GetColumnToIndex("REMARK");                   //적요
            pGDCol[3] = pGrid.GetColumnToIndex("CUSTOMER_NAME");            //거래처
            pGDCol[4] = pGrid.GetColumnToIndex("BANK_NAME");                //은행
            pGDCol[5] = pGrid.GetColumnToIndex("ORDINARY_AMOUNT");          //보통예금
            pGDCol[6] = pGrid.GetColumnToIndex("ORDINARY_CURR_AMOUNT");     //외화보통예금
            pGDCol[7] = pGrid.GetColumnToIndex("DEPOSIT_AMOUNT");           //정기적금
            pGDCol[8] = pGrid.GetColumnToIndex("CASH_AMOUNT");              //현금
            pGDCol[9] = pGrid.GetColumnToIndex("TOTAL_AMOUNT");             //총합계
        }

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDCol)
        {
            // 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDCol = new int[12];

            // 그리드 or 아답터 위치.
            pGDCol[0] = pGrid.GetColumnToIndex("ACCOUNT_CODE");
            pGDCol[1] = pGrid.GetColumnToIndex("ACCOUNT_NAME");
            pGDCol[2] = pGrid.GetColumnToIndex("ACCOUNT_DESC");
            pGDCol[3] = pGrid.GetColumnToIndex("DEPT_DESC");
            pGDCol[4] = pGrid.GetColumnToIndex("THIS_BUDGET_AMOUNT");
            pGDCol[5] = pGrid.GetColumnToIndex("THIS_SLIP_AMOUNT");
            pGDCol[6] = pGrid.GetColumnToIndex("THIS_GAP_AMOUNT");
            pGDCol[7] = pGrid.GetColumnToIndex("THIS_USE_RATE");
            pGDCol[8] = pGrid.GetColumnToIndex("BUDGET_AMOUNT");
            pGDCol[9] = pGrid.GetColumnToIndex("SLIP_AMOUNT");
            pGDCol[10] = pGrid.GetColumnToIndex("GAP_AMOUNT");
            pGDCol[11] = pGrid.GetColumnToIndex("USE_RATE");
        }

        #endregion;

        #region ----- Array Set 2 ----

        //private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        //{// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
        //    pGDColumn = new int[3];
        //    pXLColumn = new int[3];
        //    // 그리드 or 아답터 위치.
        //    pGDColumn[0] = pGrid.GetColumnToIndex("VAT_COUNT");
        //    pGDColumn[1] = pGrid.GetColumnToIndex("GL_AMOUNT");
        //    pGDColumn[2] = pGrid.GetColumnToIndex("VAT_AMOUNT");

        //    // 엑셀에 인쇄해야 할 위치.
        //    pXLColumn[0] = 20;
        //    pXLColumn[1] = 25;
        //    pXLColumn[2] = 30;
        //}

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

        #region ----- Header Write Method ----

        public void HeaderWrite_1(object pPeriod_Term)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);
            string vPrintDateTime = string.Format("[{0} {1}]", vPrintingDate, vPrintingTime);

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // 인쇄 기간.
                vXLine = 3;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, pPeriod_Term);

                //인쇄 일시.
                vXLine = 42;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintDateTime);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        public void HeaderWrite_2(object pBUDGET_YEAR, object pBUDGET_MONTH)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);
            string vPrintDateTime = string.Format("[{0} {1}]", vPrintingDate, vPrintingTime);

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // 년도
                vXLine = 3;
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine, vXLColumn, pBUDGET_YEAR);

                // 월
                vXLine = 4;
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine, vXLColumn, pBUDGET_MONTH);

                vXLine = 33;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintDateTime);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Header1 Write Method ----

        private void XLHeader1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int[] pGDColumn, int[] pXLColumn)
        {// 헤더 인쇄.
            int vXLine = 9; //엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                for (int i = 0; i <= pGrid.RowCount; i++)
                {
                    // 숫자형 예시.
                    vGDColumnIndex = pGDColumn[0];
                    vXLColumnIndex = pXLColumn[0];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
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

                    // 숫자형 예시.
                    vGDColumnIndex = pGDColumn[1];
                    vXLColumnIndex = pXLColumn[1];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
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

                    // 숫자형 예시.
                    vGDColumnIndex = pGDColumn[2];
                    vXLColumnIndex = pXLColumn[2];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
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
                    vXLine++;
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

        #region ----- Line Write Method -----

        private int XLLine_1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pXLine, int[] pGDCol)
        {
            // pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vGDColumn = 0;
            int vXLColumn = 0;
                        
            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            DateTime vCONVERT_DATE = new DateTime();
            bool IsConvert = false;

            try
            { 
                // 타겟쉬트 활성화
                mPrinting.XLActiveSheet(mTargetSheet);

                //구분
                vGDColumn = pGDCol[0];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //일자
                vGDColumn = pGDCol[1];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertDate(vObject, out vCONVERT_DATE);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vCONVERT_DATE.ToShortDateString());
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //적요
                vGDColumn = pGDCol[2];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                ////거래처
                //vGDColumn = pGDCol[3];
                //vObject = pGrid.GetCellValue(pRow, vGDColumn);
                //IsConvert = IsConvertString(vObject, out vConvertString);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0}", vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 26;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //은행
                vGDColumn = pGDCol[4];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //보통예금
                vGDColumn = pGDCol[5];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외화보통예금
                vGDColumn = pGDCol[6];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //현금
                vGDColumn = pGDCol[8];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 52;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //합계
                vGDColumn = pGDCol[9];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 59;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

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

        private int XLLine_2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pXLine, int[] pGDCol, bool pPrint_Flag)
        {
            // pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            // pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vGDColumn = 0;
            int vXLColumn = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            DateTime vCONVERT_DATE = new DateTime();
            bool IsConvert = false;

            try
            {
                // 타겟쉬트 활성화
                mPrinting.XLActiveSheet(mTargetSheet);

                //계정과목
                if (pPrint_Flag == true)
                {
                    //계정과목
                    vGDColumn = pGDCol[2];
                    vObject = pGrid.GetCellValue(pRow, vGDColumn);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vConvertString = string.Format("{0}", vObject);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    vXLColumn = 1;
                    mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                }

                //부서명
                vGDColumn = pGDCol[3];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //예산
                vGDColumn = pGDCol[4];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //실적
                vGDColumn = pGDCol[5];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //차이
                vGDColumn = pGDCol[6];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사용율
                vGDColumn = pGDCol[7];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal == 0)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###.00}", vConvertDecimal);
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 42;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //예산
                vGDColumn = pGDCol[8];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 47;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //실적
                vGDColumn = pGDCol[9];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 53;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //차이
                vGDColumn = pGDCol[10];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 59;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사용율
                vGDColumn = pGDCol[11];
                vObject = pGrid.GetCellValue(pRow, vGDColumn);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal == 0)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###.00}", vConvertDecimal);
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 65;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

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

        #region ----- TOTAL AMOUNT Write Method -----

        //private int XL_TOTAL_Line(int pXLine, int[] pXLColumn)
        //{// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
        //    int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

        //    int vXLColumnIndex = 0;

        //    string vConvertString = string.Empty;
        //    decimal vConvertDecimal = 0m;
        //    bool IsConvert = false;

        //    try
        //    { // 원본을 복사해서 타겟 에 복사해 넣음.(
        //        mPrinting.XLActiveSheet(mTargetSheet);

        //        //10 - 보증금
        //        vXLColumnIndex = pXLColumn[10];
        //        IsConvert = IsConvertNumber(mTOT_DEPOSIT_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //11 - 월임대료
        //        vXLColumnIndex = pXLColumn[11];
        //        IsConvert = IsConvertNumber(mTOT_MONTHLY_RENT_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //12 - 합계
        //        vXLColumnIndex = pXLColumn[12];
        //        IsConvert = IsConvertNumber(mTOT_LEASE_SUM_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //13 - 보증금이자
        //        vXLColumnIndex = pXLColumn[13];
        //        IsConvert = IsConvertNumber(mTOT_DEPOSIT_INTEREST_AMT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //14 - 월임대료(계)
        //        vXLColumnIndex = pXLColumn[14];
        //        IsConvert = IsConvertNumber(mTOT_MONTHLY_RENT_SUM_AMT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //-------------------------------------------------------------------
        //        vXLine = vXLine + 1;
        //        //-------------------------------------------------------------------
        //    }
        //    catch (System.Exception ex)
        //    {
        //        mMessageError = ex.Message;
        //        mAppInterface.OnAppMessageEvent(mMessageError);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    pXLine = vXLine;

        //    return pXLine;
        //}

        #endregion;

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite_1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {// 실제 호출되는 부분.
            mPageNumber = 0;
            string vMessage = string.Empty;
            
            int vTotalRow = 0;
            int vPageRowCount = 0;
            int[] vGDCol;

            // 인쇄 1장의 최대 인쇄정보.
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 65;
            mCopy_EndRow = 42;
            mPrintingLastRow = 41;  //최종 인쇄 라인.

            mCurrentRow = 6;       //현재 인쇄되는 row 위치.
            mDefaultPageRow = 5;    // 페이지 증가후 PageCount 기본값.
            
            try
            {
                // 실제인쇄되는 행수.
                //int vBy = 35;         
                vTotalRow = pGrid.RowCount;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.

                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                #region ----- Header Write ----
                
                //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                //XLHeader1(pGrid, vGDColumn, vXLColumn);  // 헤더 인쇄.

                #endregion;

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    SetArray1(pGrid, out vGDCol);
                    for(int vRow =0; vRow < pGrid.RowCount; vRow++)
                    {
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XLLine_1(pGrid, vRow, mCurrentRow, vGDCol); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow - 1)
                        {
                            //XL_TOTAL_Line(12, vXLColumn);
                        }
                        else
                        { 
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
                }
                #endregion;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }

        public int LineWrite_2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {// 실제 호출되는 부분.
            mPageNumber = 0;
            string vMessage = string.Empty;
            bool vPrint_Flag = false;
            string vACCOUNT_CODE = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vIDX_ACCOUNT_CODE = 0;
            int[] vGDCol;

            // 인쇄 1장의 최대 인쇄정보.
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 69;
            mCopy_EndRow = 33;
            mPrintingLastRow = 32;  //최종 인쇄 라인.

            mCurrentRow = 7;       //현재 인쇄되는 row 위치.
            mDefaultPageRow = 6;    // 페이지 증가후 PageCount 기본값.

            try
            {
                // 실제인쇄되는 행수.
                //int vBy = 35;         
                vTotalRow = pGrid.RowCount;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.

                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                #region ----- Header Write ----

                //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                //XLHeader1(pGrid, vGDColumn, vXLColumn);  // 헤더 인쇄.

                #endregion;

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    SetArray2(pGrid, out vGDCol);
                    vIDX_ACCOUNT_CODE = pGrid.GetColumnToIndex("ACCOUNT_CODE");
                    for (int vRow = 0; vRow < pGrid.RowCount; vRow++)
                    {
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        //계정코드 동일 여부 체크.
                        vPrint_Flag = true;
                        if (vACCOUNT_CODE == null || vACCOUNT_CODE == string.Empty || mIsNewPage == true)
                        {

                        }
                        else if (vACCOUNT_CODE != iString.ISNull(pGrid.GetCellValue(vRow, vIDX_ACCOUNT_CODE)))
                        {

                        }
                        else
                        {
                            vPrint_Flag = false;
                            mPrinting.XL_LineClearTOP(mCurrentRow, 1, 16);
                        }
                        vACCOUNT_CODE = iString.ISNull(pGrid.GetCellValue(vRow, vIDX_ACCOUNT_CODE));

                        mCurrentRow = XLLine_2(pGrid, vRow, mCurrentRow, vGDCol, vPrint_Flag); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow - 1)
                        {
                            //XL_TOTAL_Line(12, vXLColumn);
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
                }
                #endregion;
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

        private void IsNewPage(int pPageRowCount)
        {
            if (pPageRowCount == mPrintingLastRow)
            { // pPrintingLine : 현재 출력된 행.
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCurrentRow + 1);
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

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pActiveSheet);            
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol); 
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.


            mPageNumber++; //페이지 번호
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
            if (pSaveFileName == string.Empty)
            {
                return;
            }

            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
