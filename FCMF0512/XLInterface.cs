using System;

namespace FCMF0512
{
    interface XLInterface
    {
        #region ----- Property -----

        string ErrorMessage
        {
            get;
        }
        string OpenFileNameExcel
        {
            set;
        }
        int PrintingLineSTART
        {
            set;
            get;
        }
        int CopyLineSUM
        {
            set;
            get;
        }
        int PrintingLineFIRST
        {
            set;
            get;
        }

        #endregion;

        #region ----- Methods -----

        bool XLFileOpen();
        void Dispose();
        int LineWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx[] pGrid, string pDate);
        void Printing(int pPageSTART, int pPageEND);
        void Save(string pSaveFileName);

        #endregion;
    }
}
