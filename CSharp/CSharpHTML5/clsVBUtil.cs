
using System;

namespace LibIAVB2
{
public static class VBUtil
{

    //public VBUtil()
    //{
    //}
    
    #region VBUtil

    public static string sTrimRecursif(string sTxt)
    {
        int iLong = sTxt.Length;
        if ((sTxt == null) || (iLong == 0))
        {
            return "";
        }
        char VB_t_char_L0 = sTxt[0];
        if (VB_t_char_L0 == ' ')
        {
            return sTrimRecursif(sTxt.Substring(1));
        }
        char VB_t_char_L1 = sTxt[iLong - 1];
        if (VB_t_char_L1 == ' ')
        {
            return sTrimRecursif(sTxt.Substring(0, iLong - 1));
        }
        return sTxt;
    }

    public static int iVBLen(string sTxt)
    {
        if (sTxt == null) return 0;
        return sTxt.Length;
    }

    public static int iVBVal(string sTxt)
    {
        int iRes;
        int.TryParse(sTxt, out iRes);
        return iRes;
    }

    public static string sVBLeft(string sTxt, int iLeft)
    {
        int iLong = sTxt.Length;
        if (iLeft > iLong)
        {
            return sTxt;
        }
        return sTxt.Substring(0, iLeft);
    }

    public static string sVBMid(string sTxt, int iDeb, int iLong)
    {
        if (((iLong + iDeb) - 1) > sTxt.Length)
        {
            return "";
        }
        return sTxt.Substring(iDeb - 1, iLong);
    }

    public static string sVBMid2(string sTxt, int iDeb)
    {
        return sTxt.Substring(iDeb - 1);
    }

    public static string sVBRight(string sTxt, int iRight)
    {
        int iLong = sTxt.Length;
        if (iRight > iLong)
        {
            return sTxt;
        }
        return sTxt.Substring(iLong - iRight, iRight);
    }

    public static string sVBTrim(string sTxt)
    {
        return sTrimRecursif(sTxt);
    }

    //public string VBChr(int CharCode)
    //{
    //    return Convert.ToChar(CharCode).ToString() ;
    //}

#endregion
	
}
}
