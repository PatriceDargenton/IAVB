
using System;
using System.Text; // Encoding

namespace IAVB
{
public static class clsVBUtil
{

public static int iVBLen(string sTxt)
{
    if (sTxt == null) return 0;
    return sTxt.Length;
}

public static int iVBVal(string sTxt, int iValDef = 0)
{
    int iRes;
    if (int.TryParse(sTxt, out iRes)) return iRes;
    return iValDef;
}

public static string sVBLeft(string sTxt, int iLeft)
{
    if (string.IsNullOrEmpty(sTxt)) return "";
    int iLong = sTxt.Length;
    if (iLeft > iLong)
    {
        return sTxt;
    }
    return sTxt.Substring(0, iLeft);
}

public static string sVBMid(string sTxt, int iDeb, int iLong)
{
    if (string.IsNullOrEmpty(sTxt)) return "";
    if (((iLong + iDeb) - 1) > sTxt.Length)
    {
        return "";
    }
    return sTxt.Substring(iDeb - 1, iLong);
}

public static string sVBMid2(string sTxt, int iDeb)
{
    if (string.IsNullOrEmpty(sTxt)) return "";
    return sTxt.Substring(iDeb - 1);
}

public static string sVBRight(string sTxt, int iRight)
{
    if (string.IsNullOrEmpty(sTxt)) return "";
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

private static string sTrimRecursif(string sTxt)
{
    if (string.IsNullOrEmpty(sTxt)) return "";
    int iLong = sTxt.Length;
    if ((sTxt == null) || (iLong == 0))
    {
        return "";
    }
    char c1 = sTxt[0];
    if (c1 == ' ')
    {
        return sTrimRecursif(sTxt.Substring(1));
    }
    char c2 = sTxt[iLong - 1];
    if (c2 == ' ')
    {
        return sTrimRecursif(sTxt.Substring(0, iLong - 1));
    }
    return sTxt;
}

}
}
