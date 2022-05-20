using System;
using ISCommonUtil;

namespace FCMF0814
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
        private string mTargetSheet = "DESTINATION";
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
        private int mCopy_EndCol = 43;
        private int mCopy_EndRow = 55;
        private int mCopy_2nd_EndRow = 116;     //2번째 페이지는 해당있을 경우만 인쇄.

        private int m1stLastRow = 49;       //첫장 최종 인쇄 라인.
        //private int m1stCurrentRowAdd = 21;  


        private int mPrintingLastRow = 52;  //최종 인쇄 라인 다음 라인.

        private int mCurrentRow = 39;       //현재 인쇄되는 row 위치.
        //private int mDefaultPageRow = 12;   //페이지 skip후 적용되는 기본 PageRowCount 기본값-시작위치.
        //private int mCurrentRowAdd = 18;    //페이지 skip후 기본적으로 증가하는 현재 row 값.

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

        #region ----- Array Set 1 (총합계)----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[12];
            pXLColumn = new int[12];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("CUSTOMER_COUNT");
            pGDColumn[1] = pGrid.GetColumnToIndex("VAT_COUNT");
            pGDColumn[2] = pGrid.GetColumnToIndex("GL_AMOUNT_5");
            pGDColumn[3] = pGrid.GetColumnToIndex("GL_AMOUNT_4");
            pGDColumn[4] = pGrid.GetColumnToIndex("GL_AMOUNT_3");
            pGDColumn[5] = pGrid.GetColumnToIndex("GL_AMOUNT_2");
            pGDColumn[6] = pGrid.GetColumnToIndex("GL_AMOUNT_1");
            pGDColumn[7] = pGrid.GetColumnToIndex("VAT_AMOUNT_5");
            pGDColumn[8] = pGrid.GetColumnToIndex("VAT_AMOUNT_4");
            pGDColumn[9] = pGrid.GetColumnToIndex("VAT_AMOUNT_3");
            pGDColumn[10] = pGrid.GetColumnToIndex("VAT_AMOUNT_2");
            pGDColumn[11] = pGrid.GetColumnToIndex("VAT_AMOUNT_1");
                        
            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 19;
            pXLColumn[1] = 24;
            pXLColumn[2] = 28;
            pXLColumn[3] = 31;
            pXLColumn[4] = 34;
            pXLColumn[5] = 37;
            pXLColumn[6] = 40;
            pXLColumn[7] = 42;
            pXLColumn[8] = 45;
            pXLColumn[9] = 48;
            pXLColumn[10] = 51;
            pXLColumn[11] = 54;
        }

        #endregion;

        #region ----- Array Set 2 (명세) -----

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[15];
            pXLColumn = new int[15];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = 0;
            pGDColumn[1] = pGrid.GetColumnToIndex("TAX_REG_NO");
            pGDColumn[2] = pGrid.GetColumnToIndex("CUSTOMER_DESC");
            pGDColumn[3] = pGrid.GetColumnToIndex("VAT_COUNT");
            pGDColumn[4] = pGrid.GetColumnToIndex("GL_AMOUNT_5");
            pGDColumn[5] = pGrid.GetColumnToIndex("GL_AMOUNT_4");
            pGDColumn[6] = pGrid.GetColumnToIndex("GL_AMOUNT_3");
            pGDColumn[7] = pGrid.GetColumnToIndex("GL_AMOUNT_2");
            pGDColumn[8] = pGrid.GetColumnToIndex("GL_AMOUNT_1");
            pGDColumn[9] = pGrid.GetColumnToIndex("VAT_AMOUNT_5");
            pGDColumn[10] = pGrid.GetColumnToIndex("VAT_AMOUNT_4");
            pGDColumn[11] = pGrid.GetColumnToIndex("VAT_AMOUNT_3");
            pGDColumn[12] = pGrid.GetColumnToIndex("VAT_AMOUNT_2");
            pGDColumn[13] = pGrid.GetColumnToIndex("VAT_AMOUNT_1");
            pGDColumn[14] = 0;


            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 3;
            pXLColumn[1] = 6;
            pXLColumn[2] = 12;
            pXLColumn[3] = 22;
            pXLColumn[4] = 25;
            pXLColumn[5] = 28;
            pXLColumn[6] = 31;
            pXLColumn[7] = 34;
            pXLColumn[8] = 37;
            pXLColumn[9] = 39;
            pXLColumn[10] = 42;
            pXLColumn[11] = 45;
            pXLColumn[12] = 48;
            pXLColumn[13] = 51;
            pXLColumn[14] = 53;
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

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, object pPERIOD, object pISSUE_PERIOD)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;
            
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                vXLine = 6;
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, pISSUE_PERIOD);

                // 기간.
                vXLine = 6;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, pPERIOD);

                // 신고자 인적사항.
                vXLine = 7;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["CORPORATE_NAME"]);
                
                vXLColumn = 22;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["PRESIDENT_NAME"]);

                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_1"]);
                vXLColumn = 32;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_2"]);
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_3"]);
                vXLColumn = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_4"]);
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_5"]);
                vXLColumn = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_6"]);
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_7"]);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_8"]);
                vXLColumn = 41;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_9"]);
                vXLColumn = 42;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_10"]);

                vXLine = 8;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["LEGAL_NUM"]);

                vXLine = 9;
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["TEL_NUM"]);
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["TEL_NUM"]);
                
                vXLine = 10;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["ADDRESS"]);

                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["EMAIL"]);

                // 은행.
                vXLine = 42;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["BANK_NAME"]);
                // 계좌.
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["BANK_ACCOUNT_NUM"]);

                // 신고자.
                vXLine = 48;
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["CORPORATE_NAME"]);
                                
                //2페이지.
                vXLine = 58;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_1"]);
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_2"]);
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_3"]);
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_4"]);
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_5"]);
                vXLColumn = 16;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_6"]);
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_7"]);
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_8"]);
                vXLColumn = 19;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_9"]);
                vXLColumn = 20;
                mPrinting.XLSetCell(vXLine, vXLColumn, pAdapter.CurrentRow["VAT_NUM_10"]);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Header1 (합계) Write Method ----

        private void XLHeader1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int[] pGDColumn, int[] pXLColumn)
        {// 헤더 인쇄.
            int vXLine = 0; //엑셀에 내용이 표시되는 행 번호

            int vIDX_LINE_TYPE = pGrid.GetColumnToIndex("LINE_TYPE");
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

                for (int i = 0; i < pGrid.RowCount; i++)
                {
                    // 총합계 구분에 따라 인쇄 ROW 지정.
                    if ("TS" == iString.ISNull(pGrid.GetCellValue(i, vIDX_LINE_TYPE)))
                    {//총합계
                        vXLine = 19;
                    }
                    else if ("YC" == iString.ISNull(pGrid.GetCellValue(i, vIDX_LINE_TYPE)))
                    {//전자세금계산서외의  발급받은분 - 사업자발행분.
                        vXLine = 21;
                    }
                    else if ("YP" == iString.ISNull(pGrid.GetCellValue(i, vIDX_LINE_TYPE)))
                    {//전자세금계산서 발급받은분 - 주민등록번호발행분.
                        vXLine = 23;
                    }
                    else if ("YS" == iString.ISNull(pGrid.GetCellValue(i, vIDX_LINE_TYPE)))
                    {//전자세금계산서 발급받은분 - 소계.
                        vXLine = 25;
                    }
                    else if ("NC" == iString.ISNull(pGrid.GetCellValue(i, vIDX_LINE_TYPE)))
                    {//전자세금계산서외의  발급받은분 - 사업자발행분.
                        vXLine = 27;
                    }
                    else if ("NP" == iString.ISNull(pGrid.GetCellValue(i, vIDX_LINE_TYPE)))
                    {//전자세금계산서외의  발급받은분 - 주민등록번호발행분.
                        vXLine = 29;
                    }
                    else if ("NS" == iString.ISNull(pGrid.GetCellValue(i, vIDX_LINE_TYPE)))
                    {//전자세금계산서외의  발급받은분 - 소계.
                        vXLine = 31;
                    }                    
                    
                    //0 - 매입처수.
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
                    //1 -  매수.
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
                    //-- 공급가액 --//
                    //2 - 조
                    vGDColumnIndex = pGDColumn[2];
                    vXLColumnIndex = pXLColumn[2];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //3 - 십억
                    vGDColumnIndex = pGDColumn[3];
                    vXLColumnIndex = pXLColumn[3];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //4 - 백만
                    vGDColumnIndex = pGDColumn[4];
                    vXLColumnIndex = pXLColumn[4];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //5 - 천
                    vGDColumnIndex = pGDColumn[5];
                    vXLColumnIndex = pXLColumn[5];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //6 - 일
                    vGDColumnIndex = pGDColumn[6];
                    vXLColumnIndex = pXLColumn[6];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //-- 세액 --//
                    //7 - 조
                    vGDColumnIndex = pGDColumn[7];
                    vXLColumnIndex = pXLColumn[7];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //8 - 십억
                    vGDColumnIndex = pGDColumn[8];
                    vXLColumnIndex = pXLColumn[8];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //9 - 백만
                    vGDColumnIndex = pGDColumn[9];
                    vXLColumnIndex = pXLColumn[9];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //10 - 천
                    vGDColumnIndex = pGDColumn[10];
                    vXLColumnIndex = pXLColumn[10];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //11 - 일
                    vGDColumnIndex = pGDColumn[11];
                    vXLColumnIndex = pXLColumn[11];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

        #region ----- Excel Write [Line - 1st PAGE ] Method -----

        private void XLLine_1(InfoSummit.Win.ControlAdv.ISDataAdapter pData1)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = 21; //엑셀에 내용이 표시되는 행 번호
            int vXLColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            DateTime vCONVERT_DATE = new DateTime();
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);
                
                //(1)-금액
                vXLine = 14;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_TAX_INVOICE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(1)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_TAX_INVOICE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //(2)-금액
                vXLine = 15;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_TAX_BUYER_INVOICE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(2)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_TAX_BUYER_INVOICE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //(3)-금액
                vXLine = 16;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_TAX_CREDIT_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(3)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_TAX_CREDIT_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //(4)-금액
                vXLine = 17;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_TAX_ETC_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(4)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_TAX_ETC_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(5)-금액
                vXLine = 18;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_ZERO_INVOICE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(5)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_ZERO_INVOICE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(6)-금액
                vXLine = 19;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_ZERO_ETC_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(6)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_ZERO_ETC_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(7)-금액
                vXLine = 20;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_SCHEDULE_OMIT_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(7)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_SCHEDULE_OMIT_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(8)-금액
                vXLine = 21;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_BAD_DEBT_TAX_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(8)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_BAD_DEBT_TAX_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(9)-금액
                vXLine = 22;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["SALES_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(9)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["SALES_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(10)-금액
                vXLine = 23;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["P_TAX_INVOICE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(10)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["P_TAX_INVOICE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(11)-금액
                vXLine = 24;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["P_TAX_ASSET_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(11)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["P_TAX_ASSET_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(12)-금액
                vXLine = 25;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["P_SCHEDULE_OMIT_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(12)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["P_SCHEDULE_OMIT_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(13)-금액
                vXLine = 26;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["P_BUYER_INVOICE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(13)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["P_BUYER_INVOICE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(14)-금액
                vXLine = 27;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["P_ETC_DED_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(14)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["P_ETC_DED_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(15)-금액
                vXLine = 28;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["PURCHASE_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(15)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["PURCHASE_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(16)-금액
                vXLine = 29;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["P_NOT_DED_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(16)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["P_NOT_DED_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(17)-금액
                vXLine = 30;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["P_SUB_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(17)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["P_SUB_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(18)-금액
                vXLine = 31;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["CALCULATE_TAX_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(납부세액)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["CALCULATE_TAX_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(18)-금액
                vXLine = 32;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["R_ETC_DED_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(18)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["R_ETC_DED_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(19)-금액
                vXLine = 33;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["R_CREDIT_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(19)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["R_CREDIT_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(20)-금액
                vXLine = 34;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["REDUCE_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(20)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["REDUCE_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(21)-금액
                vXLine = 35;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["SCHEDULE_YET_REFUND_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(21)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["SCHEDULE_YET_REFUND_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(22)-금액
                vXLine = 36;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["SCHEDULE_NOTICE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(22)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["SCHEDULE_NOTICE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(23)-금액
                vXLine = 37;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["GOLD_BAR_BUYER_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(23)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["GOLD_BAR_BUYER_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(24)-금액
                vXLine = 38;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["TAX_ADDITION_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(24)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["TAX_ADDITION_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(25)-금액
                vXLine = 39;
                //(25)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["BALANCE_TAX_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //작성일자.
                vXLine = 47;
                vXLColumnIndex = 25;
                vObject = vObject = pData1.CurrentRow["WRITE_DATE"];
                IsConvert = IsConvertDate(vObject, out vCONVERT_DATE);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}년 {1:D2}월  {2:D2} 일", vCONVERT_DATE.Year, vCONVERT_DATE.Month, vCONVERT_DATE.Day);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;        // 2줄씩 증가.
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Excel Write [Line - 2nd PAGE ] Method -----

        private void XLLine_2(InfoSummit.Win.ControlAdv.ISDataAdapter pData1)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = 21; //엑셀에 내용이 표시되는 행 번호
            int vXLColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            //DateTime vCONVERT_DATE = new DateTime(); ;
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                //(31)-금액
                vXLine = 61;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["SS_TAX_INVOICE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(31)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["SS_TAX_INVOICE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //(32)-금액
                vXLine = 62;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["SS_TAX_ETC_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(32)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["SS_TAX_ETC_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //(33)-금액
                vXLine = 63;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["SS_ZERO_INVOICE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(33)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["SS_ZERO_INVOICE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //(34)-금액
                vXLine = 64;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["SS_ZERO_ETC_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(4)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["SS_ZERO_ETC_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(35)-금액
                vXLine = 65;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_SALES_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(35)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_SALES_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(36)-금액
                vXLine = 66;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["SP_TAX_INVOICE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(36)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["SP_TAX_INVOICE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(37)-금액
                vXLine = 67;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["SP_ETC_DED_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(37)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["SP_ETC_DED_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(38)-금액
                vXLine = 68;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["S_PURCHASE_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(38)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["S_PURCHASE_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(39)-금액
                vXLine = 71;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["E_CREDIT_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(39)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["E_CREDIT_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(40)-금액
                vXLine = 72;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["E_CREDIT_ASSET_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(10)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["E_CREDIT_ASSET_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(41)-금액
                vXLine = 73;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["E_DEEMED_IP_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(41)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["E_DEEMED_IP_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(42)-금액
                vXLine = 74;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["E_RECYCLE_IP_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(42)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["E_RECYCLE_IP_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(43)-금액
                vXLine = 75;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["E_GOLD_BAR_IP_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(43)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["E_GOLD_BAR_IP_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(44)-금액
                vXLine = 76;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["E_TAX_BUSINESS_IP_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(44)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["E_TAX_BUSINESS_IP_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(45)-금액
                vXLine = 77;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["E_STOCK_IP_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(45)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["E_STOCK_IP_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(46)-금액
                vXLine = 78;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["E_BAD_TAX_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(46)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["E_BAD_TAX_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(47)-금액
                vXLine = 79;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["ETC_DED_IP_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(47)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["ETC_DED_IP_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(48)-금액
                vXLine = 82;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["N_NOT_DED_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(48)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["N_NOT_DED_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(49)-금액
                vXLine = 83;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["N_COMMON_IP_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(49)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["N_COMMON_IP_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(50)-금액
                vXLine = 84;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["N_BAD_RECEIVE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(50)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["N_BAD_RECEIVE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(51)-금액
                vXLine = 85;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["NOT_DED_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(51)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["NOT_DED_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(52)-금액
                vXLine = 88;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["R_ETAX_REPORT_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(52)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["R_ETAX_REPORT_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(53)-금액
                vXLine = 89;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["R_ETAX_ISSUE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(53)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["R_ETAX_ISSUE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(54)-금액
                vXLine = 90;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["R_TAXI_TRANSPORT_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(54)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["R_TAXI_TRANSPORT_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(55)-금액
                vXLine = 91;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["R_CASH_BILL_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(55)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["R_CASH_BILL_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(56)-금액
                vXLine = 92;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["R_ETC_DED_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(56)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["R_ETC_DED_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(57)-금액
                vXLine = 93;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["REDUCE_DED_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(57)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["REDUCE_DED_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(58)-금액
                vXLine = 96;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_VAT_NUM_UNENROLL_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(58)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_VAT_NUM_UNENROLL_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(59)-금액
                vXLine = 97;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_TAX_INVOICE_DELAY_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(59)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_TAX_INVOICE_DELAY_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(60)-금액
                vXLine = 98;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_TAX_INVOICE_UNISSUE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(60)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_TAX_INVOICE_UNISSUE_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(61)-금액
                vXLine = 99;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_ETAX_UNSEND_IN_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(61)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_ETAX_UNSEND_IN_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(62)-금액
                vXLine = 100;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_ETAX_UNSEND_OVER_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(62)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_ETAX_UNSEND_OVER_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(63)-금액
                vXLine = 101;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_TAX_INV_SUM_BAD_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(63)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_TAX_INV_SUM_BAD_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(64)-금액
                vXLine = 102;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_REPORT_BAD_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(64)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_REPORT_BAD_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(65)-금액
                vXLine = 103;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_PAYMENT_BAD_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(65)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_PAYMENT_BAD_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(66)-금액
                vXLine = 104;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_ZERO_REPORT_BAD_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(66)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_ZERO_REPORT_BAD_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(67)-금액
                vXLine = 105;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["A_CASH_SALES_UNREPORT_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(67)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["A_CASH_SALES_UNREPORT_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(68)-금액
                vXLine = 106;
                vXLColumnIndex = 21;
                vObject = pData1.CurrentRow["TAX_ADDITION_SUM_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(68)-부가세
                vXLColumnIndex = 34;
                vObject = pData1.CurrentRow["TAX_ADDITION_SUM_VAT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(73)-금액
                vXLine = 114;
                vXLColumnIndex = 17;
                vObject = pData1.CurrentRow["BILL_ISSUE_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //(74)-금액
                vXLine = 115;
                vXLColumnIndex = 17;
                vObject = pData1.CurrentRow["BILL_RECEIPT_AMT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;        // 2줄씩 증가.
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Excel Write [Line - 3rd : 과세표준 명세서 ] Method -----

        public void XLLine_3(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.

            decimal vTAX_STANDARD_ID = 0;
            int vGrid_Col_IDX = 0;
            int vXLine = 47; //엑셀에 내용이 표시되는 행 번호
            int vXLColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            { //
                mPrinting.XLActiveSheet(mSourceSheet1);

                for (int nRow = 0; nRow < pGrid.RowCount; nRow++)
                {
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("TAX_STANDARD_ID");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    vTAX_STANDARD_ID = iString.ISDecimaltoZero(vObject);
                    if (vTAX_STANDARD_ID == -1)
                    {
                        vXLine = 52;
                    }
                    else
                    {
                        vXLine = vXLine + 1;
                    }
                    // 업태.
                    vXLColumnIndex = 4;
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("BUSINESS_TYPE");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (vTAX_STANDARD_ID == -1)
                    {
                    }
                    else if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    // 업종.
                    vXLColumnIndex = 8;
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("BUSINESS_ITEM");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    // 업종코드.
                    vXLColumnIndex = 13;
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("BUSINESS_ITEM_CODE_1");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    // 업종코드.
                    vXLColumnIndex = 14;
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("BUSINESS_ITEM_CODE_2");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    // 업종코드.
                    vXLColumnIndex = 15;
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("BUSINESS_ITEM_CODE_3");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    // 업종코드.
                    vXLColumnIndex = 16;
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("BUSINESS_ITEM_CODE_4");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    // 업종코드.
                    vXLColumnIndex = 17;
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("BUSINESS_ITEM_CODE_5");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    // 업종코드.
                    vXLColumnIndex = 18;
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("BUSINESS_ITEM_CODE_6");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    // 금액.
                    vXLColumnIndex = 19;
                    vGrid_Col_IDX = pGrid.GetColumnToIndex("TAX_AMOUNT");
                    vObject = pGrid.GetCellValue(nRow, vGrid_Col_IDX);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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


        #region ----- TOTAL AMOUNT Write Method -----

        private int XLTOTAL_Line(int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            //int vXLColumnIndex = 0;

            //string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;
            //bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                ////12-건수
                //vXLColumnIndex = 12;
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
                ////22-공급가액
                //vXLColumnIndex = 22;
                //IsConvert = IsConvertNumber(mTOT_GL_AMOUNT, out vConvertDecimal);
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
                ////34-세액
                //vXLColumnIndex = 34;
                //IsConvert = IsConvertNumber(mTOT_VAT_AMOUNT, out vConvertDecimal);
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
                vXLine = vXLine + 2;
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
            //int vXLRow = 51; //엑셀에 내용이 표시되는 행 번호
            //int vXLCol = 51;
            //if (iString.ISDecimaltoZero(pPageNumber) == 1)
            //{//첫장 적용
            //    vXLRow = 51; //엑셀에 내용이 표시되는 행 번호
            //    vXLCol = 52;
            //}
            //else
            //{//첫장 외.
            //    vXLRow = 54; //엑셀에 내용이 표시되는 행 번호
            //    vXLCol = 51;
            //}

            //try
            //{ // 원본을 복사해서 타겟 에 복사해 넣음.(
            //    mPrinting.XLActiveSheet(pActiveSheet);
            //    mPrinting.XLSetCell(vXLRow, vXLCol, pPageNumber);
            //}
            //catch (System.Exception ex)
            //{
            //    mMessageError = ex.Message;
            //    mAppInterface.OnAppMessageEvent(mMessageError);
            //    System.Windows.Forms.Application.DoEvents();
            //}
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte MAIN Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pData1, InfoSummit.Win.ControlAdv.ISDataAdapter pData2)
        {// 실제 호출되는 부분.
            string vMessage = string.Empty;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);

            //int[] vGDColumn;
            //int[] vXLColumn;
            int vTotalRow = 0;
            try
            {
                //// 실제인쇄되는 행수.
                if (pData2.OraSelectData.Rows.Count > 0)
                {
                    mPageNumber = 1;
                    mCopy_EndRow = mCopy_2nd_EndRow;
                }

                // 첫장 인쇄.
                vTotalRow = pData1.OraSelectData.Rows.Count;
                
                //vPageRowCount = mCurrentRow - 2;    //첫장에 대해서는 시작row부터 체크.
                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.                

                if (vTotalRow > 0)
                {
                    #region ----- Header Write ----
                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                    #endregion;
                    XLLine_1(pData1); // 현재 위치 인쇄 후 다음 인쇄행 리턴.

                    //#region ----- Line Write ----
                    //SetArray2(pGrid_Detail, out vGDColumn, out vXLColumn);
                    //for (int vRow = 0; vRow < vTotalRow; vRow++)
                    //{
                    //    vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                    //    mAppInterface.OnAppMessageEvent(vMessage);
                    //    System.Windows.Forms.Application.DoEvents();

                    //    mCurrentRow = XLLine(pGrid_Detail, vRow, mCurrentRow, vGDColumn, vXLColumn); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                    //    vPageRowCount = vPageRowCount + 2;

                    //    if (vRow == vTotalRow - 1)
                    //    {
                    //        // 마지막 데이터 이면 처리할 사항 기술
                    //        // 라인지운다 또는 합계를 표시한다 등 기술.
                    //    }
                    //    else
                    //    {
                    //        IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                    //        if (mIsNewPage == true)
                    //        {
                    //            if (mPageNumber <= 2)
                    //            {
                    //                mCurrentRow = mCurrentRow + m1stCurrentRowAdd;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                    //            }
                    //            else
                    //            {
                    //                mCurrentRow = mCurrentRow + mCurrentRowAdd;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                    //            }
                    //            vPageRowCount = mDefaultPageRow;
                    //        }
                    //    }
                    //}
                    //#endregion;
                }

                // 실제인쇄되는 행수.
                // 첫장 인쇄.
                vTotalRow = pData2.OraSelectData.Rows.Count;
                if (vTotalRow > 0)
                {
                    XLLine_2(pData2);
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

        private void IsNewPage(int pPageRowCount)
        {
            int iDefaultEndRow = 1;
            if (mPageNumber == 1)
            {
                if (pPageRowCount == m1stLastRow)
                { // pPrintingLine : 현재 출력된 행.
                    mIsNewPage = true;
                    iDefaultEndRow = mCopy_EndRow - m1stLastRow - 1;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow + iDefaultEndRow);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
            else
            {
                if (pPageRowCount == mPrintingLastRow)
                { // pPrintingLine : 현재 출력된 행.
                    mIsNewPage = true;
                    iDefaultEndRow = mCopy_EndRow - mPrintingLastRow - 1;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow + iDefaultEndRow);
                }
                else
                {
                    mIsNewPage = false;
                }
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
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;
    }
}
