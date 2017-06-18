
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//using System.Windows; // For MessageBoxImage
using System.Windows.Forms; // For MessageBox

namespace Utils
{
    public static class clsAppUtils
    {
        public static string AppTitle { get; set; }
    }

    public static class clsMessageUtils
    {
        public static readonly string sCrLf = Environment.NewLine; // "\r\n"

        public static void ShowErrorMsg(string sTxt)
        {
            MessageBox.Show(sTxt, clsAppUtils.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void ShowErrorMsg(Exception ex, 
            string sFunctionTitle = "", string sInfo = "",
            string sDetailedErrMsg = "", bool bCopyErrMsgClipboard = true)
        {
            //if (!Cursor.Current.Equals(Cursors.Default))
            //    Cursor.Current = Cursors.Default;

            var sMsg = "";
            if (!string.IsNullOrEmpty(sFunctionTitle))
                sMsg = "Function: " + sFunctionTitle;
            if (!string.IsNullOrEmpty(sInfo))
                sMsg += sCrLf + sInfo;
            if (!string.IsNullOrEmpty(sDetailedErrMsg))
                sMsg += sCrLf + sDetailedErrMsg;
            if (ex != null && !string.IsNullOrEmpty(ex.Message))
            {
                sMsg += sCrLf + ex.Message.Trim();
                if ((ex.InnerException != null))
                    sMsg += sCrLf + ex.InnerException.Message;
            }
            if (bCopyErrMsgClipboard)
            {
                bCopyToClipboard(sMsg);
                sMsg += sCrLf + "(this error message has been copied into the clipboard)";
            }

            ShowErrorMsg(sMsg);

        }
        
        public static void ShowErrorMsg(Exception ex, out string sFinalErrMsg,
            string sFunctionTitle = "", string sInfo = "",
            string sDetailedErrMsg = "", bool bCopyErrMsgClipboard = true)
        {
            //if (!Cursor.Current.Equals(Cursors.Default))
            //    Cursor.Current = Cursors.Default;

            sFinalErrMsg = "";

            var sMsg = "";
            if (!string.IsNullOrEmpty(sFunctionTitle))
                sMsg = "Function : " + sFunctionTitle;
            if (!string.IsNullOrEmpty(sInfo))
                sMsg += sCrLf + sInfo;
            if (!string.IsNullOrEmpty(sDetailedErrMsg))
                sMsg += sCrLf + sDetailedErrMsg;
            if (ex != null && !string.IsNullOrEmpty(ex.Message))
            {
                sMsg += sCrLf + ex.Message.Trim();
                if ((ex.InnerException != null))
                    sMsg += sCrLf + ex.InnerException.Message;
            }
            if (bCopyErrMsgClipboard)
            {
                bCopyToClipboard(sMsg);
                sMsg += sCrLf + "(this error message has been copied into the clipboard)";
            }

            sFinalErrMsg = sMsg;
            ShowErrorMsg(sMsg);

        }

        public static bool bCopyToClipboard(string sInfo)
        {
            // Copy text into Windows clipboard (until the application is closed)
            try
            {
                DataObject dataObj = new DataObject();
                dataObj.SetData(DataFormats.Text, sInfo);
                Clipboard.SetDataObject(dataObj);
                return true;
            }
            catch (Exception ex)
            {
                // The clipboard can be unavailable
                ShowErrorMsg(ex, "bCopyToClipboard", bCopyErrMsgClipboard:false);
                return false;
            }
        }
    }
}
