using EllieMae.Encompass.BusinessObjects.Loans;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace FormPlugin
{
    public class GeneratePDFForm
    {
        public string GeneratePdfForm(Loan loan, string inputFile,string xmlSettings,string scriptClassName)
        {


            string templateFilename = inputFile;
            string outputFilename = Environment.GetEnvironmentVariable("temp").ToString() + "\\" + Path.GetRandomFileName() + ".pdf";

            using (Stream inputPdf = new FileStream(templateFilename, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (Stream outputPdf = new FileStream(outputFilename, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var reader = new PdfReader(inputPdf);
                    var stamper = new PdfStamper(reader, outputPdf) { FormFlattening = true };

                    // Create a BaseFont representation of an internal font - Helvetica, 
                    // using the Latin code page for Windows
                    var bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252,true);
                   
                    var fieldData = getFieldData(scriptClassName, loan);

                    foreach (Plugin.Field lc in getFormFields(xmlSettings))
                    {
                        // This field will appear near the bottom left of the page (based on the Rectangle)
                        TextField tf = new TextField(stamper.Writer, new iTextSharp.text.Rectangle(lc.lx, lc.uy, lc.rx, lc.dy), lc.name) {Font = bf };
                        tf.Alignment = lc.align;
                        tf.Alignment = 6;
                        tf.FontSize = 8;
                        tf.SetExtraMargin(0f, 1.2f);
                        tf.Text = fieldData.Get(lc.name);
                        tf.Options = TextField.READ_ONLY;
                        tf.GetAppearance();

                        // Add the TextField to the PDF on page 1
                        var pageNumber = 1;

                        stamper.FormFlattening = true;
                        stamper.AcroFields.GenerateAppearances = true;
                        stamper.AddAnnotation(tf.GetTextField(), pageNumber);

                    }

                    stamper.Close();
                    reader.Close();


                }
                // Save the changes to file

            }
            return outputFilename;
        }
        public string GeneratePdfForm(Loan loan, string inputFile, string xmlSettings, string scriptClassName, BaseColor color)
        {


            string templateFilename = inputFile;
            string outputFilename = Environment.GetEnvironmentVariable("temp").ToString() + "\\" + Path.GetRandomFileName() + ".pdf";

            using (Stream inputPdf = new FileStream(templateFilename, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (Stream outputPdf = new FileStream(outputFilename, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var reader = new PdfReader(inputPdf);
                    var stamper = new PdfStamper(reader, outputPdf) { FormFlattening = true };

                    // Create a BaseFont representation of an internal font - Helvetica, 
                    // using the Latin code page for Windows
                   
                    iTextSharp.text.Font font = FontFactory.GetFont(@"c:\windows\fonts\timesi.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED,.08f, iTextSharp.text.Font.ITALIC, BaseColor.RED);
              
                    BaseFont bsf = font.BaseFont;
                 

                    var fieldData = getFieldData(scriptClassName, loan);

                    foreach (Plugin.Field lc in getFormFields(xmlSettings))
                    {
                        // This field will appear near the bottom left of the page (based on the Rectangle)

                        TextField tf = new TextField(stamper.Writer, new iTextSharp.text.Rectangle(lc.lx, lc.uy, lc.rx, lc.dy), lc.name);
                    
                        tf.Text = fieldData.Get(lc.name);
                        tf.Font = bsf;
                       
                        tf.Alignment = lc.align;
                        tf.Alignment = 6;
                    
                        tf.TextColor = BaseColor.RED;
                     
                        tf.SetExtraMargin(0f, 1.2f);
                       
                        tf.Options = TextField.READ_ONLY;
  
                        // Add the TextField to the PDF on page 1
                        var pageNumber = 1;
                      
                        stamper.FormFlattening = true;
                     
                        stamper.AddAnnotation(tf.GetTextField(), pageNumber);

                    }

                    stamper.Close();
                    reader.Close();


                }
                // Save the changes to file

            }
            return outputFilename;
        }

        private NameValueCollection getFieldData(string scriptName,Loan loan)
        {
            NameValueCollection fieldData = null;

            switch (scriptName)
            {
                case "_1008___TSUM_P1CLASS":
                        _1008___TSUM_P1CLASS scripter = new _1008___TSUM_P1CLASS();
                    fieldData = scripter.RunScript(loan);
                    break;
             
               
                default:
                    break;

            }
            return fieldData;
        }

        private ArrayList getFormFields(string xmlFile)
        {
            ArrayList pdfFieldList = new ArrayList();

            XmlReader reader = XmlReader.Create(xmlFile);
            while (reader.Read())
            {
                if ((reader.NodeType == XmlNodeType.Element) && (reader.Name == "Field"))
                {
                    Plugin.Field fi = new Plugin.Field();

                    fi.dy = float.Parse(reader.GetAttribute("dy"));
                    fi.lx = float.Parse(reader.GetAttribute("lx"));
                    fi.rx = float.Parse(reader.GetAttribute("rx"));
                    fi.uy = float.Parse(reader.GetAttribute("uy"));
                    fi.name = reader.GetAttribute("name");
                    fi.align = int.Parse(reader.GetAttribute("align"));
                    fi.fontSize = float.Parse(reader.GetAttribute("fontsize"));
                    pdfFieldList.Add(fi);
                }
            }

            return pdfFieldList;
        }

      
    }
}
 class JS
{
    private static readonly Regex newlineTailRegex = new Regex("(\\r\\n)+$");

    public static string GetStr(Loan loan, string dataFieldName)
    {

        string field = loan.Fields[dataFieldName].FormattedValue;
        return JS.newlineTailRegex.Replace(field, "");
    }

    public static double GetNum(Loan loan, string dataFieldName)
    {
        string simpleField = loan.Fields[dataFieldName].FormattedValue;
        return Jed.S2N(JS.newlineTailRegex.Replace(simpleField, ""));
    }

    public static string Dummy()
    {
        return "";
    }
}
 class Jed
{
    private static readonly Regex notNumberOrDot = new Regex("[^0-9.-]");
    private static string[] NumberName = new string[20]
    {
      "Zero",
      "One",
      "Two",
      "Three",
      "Four",
      "Five",
      "Six",
      "Seven",
      "Eight",
      "Nine",
      "Ten",
      "Eleven",
      "Twelve",
      "Thirteen",
      "Fourteen",
      "Fiftheen",
      "Sixteen",
      "Seventeen",
      "Eighteen",
      "Nineteen"
    };
    private static string[] NumberNamety = new string[10]
    {
      "dummy",
      "dummy",
      "Twenty",
      "Thirty",
      "Fourty",
      "Fifty",
      "Sixty",
      "Seventy",
      "Eighty",
      "Ninety"
    };
    private const byte DDMask = 7;
    private const byte DSMask = 8;
    private const byte CommaMask = 16;
    private const byte NIEMask = 32;
    private const byte ZFMask = 192;
    public const byte NoFormat = 0;
    public const byte NoDD = 0;
    public const byte OneDD = 1;
    public const byte TwoDD = 2;
    public const byte ThreeDD = 3;
    public const byte FourDD = 4;
    public const byte DS = 8;
    public const byte Comma = 16;
    public const byte NIE = 32;
    public const byte ZF1 = 0;
    public const byte ZF2 = 64;
    public const byte ZF3 = 128;
    public const byte ZF4 = 192;

    public static string NF(double inValue, byte flag, int padding)
    {
        if (((int)flag & 32) == 32)
            return Jed.PutNumberInEnglish(inValue);
        if (inValue == 0.0)
        {
            switch ((int)flag & 192)
            {
                case 128:
                    return "-0-";
                case 192:
                    return string.Empty;
                case 64:
                    return "0.00";
            }
        }
        NumberFormatInfo numberFormatInfo = new NumberFormatInfo();
        switch ((int)flag & 7)
        {
            case 0:
                numberFormatInfo.CurrencyDecimalDigits = 0;
                break;
            case 1:
                numberFormatInfo.CurrencyDecimalDigits = 1;
                break;
            case 2:
                numberFormatInfo.CurrencyDecimalDigits = 2;
                break;
            case 3:
                numberFormatInfo.CurrencyDecimalDigits = 3;
                break;
            case 4:
                numberFormatInfo.CurrencyDecimalDigits = 4;
                break;
        }
        numberFormatInfo.CurrencySymbol = ((int)flag & 8) != 8 ? string.Empty : "$";
        numberFormatInfo.CurrencyGroupSeparator = ((int)flag & 16) != 16 ? string.Empty : ",";
        numberFormatInfo.CurrencyNegativePattern = 12;
        return inValue.ToString("c", (IFormatProvider)numberFormatInfo);
    }

    public static string BF(bool bValue, string strValue)
    {
        if (bValue)
            return strValue;
        return "";
    }

    public static string BF(bool bValue, string trueValue, string falseValue)
    {
        if (bValue)
            return trueValue;
        return falseValue;
    }

    public static double BF(bool bValue, double trueValue, double falseValue)
    {
        if (bValue)
            return trueValue;
        return falseValue;
    }

    public static string Date()
    {
        DateTime now = DateTime.Now;
        return now.Month.ToString() + "/" + now.Day.ToString() + "/" + now.Year.ToString();
    }

    public static double S2N(string strValue)
    {
        if (strValue == null || strValue == string.Empty)
            return 0.0;
        strValue = Jed.notNumberOrDot.Replace(strValue, "");
        try
        {
            return Convert.ToDouble(strValue);
        }
        catch
        {
            return 0.0;
        }
    }

    public static string GetPhoneNo(string phoneNum)
    {
        if (phoneNum == null || phoneNum == string.Empty || phoneNum.Length <= 4)
            return string.Empty;
        return phoneNum.Substring(4);
    }

    public static string GetPhoneNoWithoutExt(string phoneNum)
    {
        if (phoneNum == null || phoneNum == string.Empty || phoneNum.Length <= 4)
            return string.Empty;
        if (phoneNum.Substring(4).Length > 8)
            return phoneNum.Substring(4, 8);
        return phoneNum.Substring(4);
    }

    public static string GetAreaCode(string phoneNum)
    {
        if (phoneNum == null || phoneNum == string.Empty || phoneNum.Length < 3)
            return string.Empty;
        return phoneNum.Substring(0, 3);
    }

    public static string GetPhoneExt(string phoneNum)
    {
        if (phoneNum == null || phoneNum == string.Empty || phoneNum.Length < 14)
            return string.Empty;
        return phoneNum.Substring(13);
    }

    public static string Min(string numStr)
    {
        if (numStr == null || numStr == string.Empty)
            return string.Empty;
        return Jed.S2N(numStr).ToString();
    }

    public static string Min(string numStr1, string numStr2)
    {
        if (numStr2 == null || numStr2 == string.Empty)
            return Jed.Min(numStr1);
        if (numStr1 == null || numStr1 == string.Empty)
            return Jed.Min(numStr2);
        return Math.Min(Jed.S2N(numStr1), Jed.S2N(numStr2)).ToString();
    }

    public static string Min(string numStr1, string numStr2, string numStr3)
    {
        if (numStr3 == null || numStr3 == string.Empty)
            return Jed.Min(numStr1, numStr2);
        if (numStr2 == null || numStr2 == string.Empty)
            return Jed.Min(numStr1, numStr3);
        if (numStr1 == null || numStr1 == string.Empty)
            return Jed.Min(numStr2, numStr3);
        return Math.Min(Math.Min(Jed.S2N(numStr1), Jed.S2N(numStr2)), Jed.S2N(numStr3)).ToString();
    }

    public static double Num()
    {
        return 0.0;
    }

    public static double Num(double dValue)
    {
        return dValue;
    }

    private static string PutSmallIntegerInEnglish(double dValue)
    {
        if (dValue == 0.0)
            return string.Empty;
        string str1 = string.Empty;
        int num = (int)dValue;
        int index1 = num / 100;
        if (index1 != 0)
            str1 = str1 + Jed.NumberName[index1] + " Hundred ";
        int index2 = num % 100;
        if (index2 == 0)
            return str1;
        if (index2 < 20)
            return str1 + Jed.NumberName[index2] + " ";
        int index3 = index2 / 10;
        int index4 = index2 % 10;
        string str2;
        if (index4 == 0)
            str2 = str1 + Jed.NumberNamety[index3] + " ";
        else
            str2 = str1 + Jed.NumberNamety[index3] + " " + Jed.NumberName[index4] + " ";
        return str2;
    }

    private static string PutNumberInEnglish(double dValue)
    {
        double num = Math.Floor(dValue);
        if (num == 0.0)
            return "Zero";
        string str1 = Jed.PutSmallIntegerInEnglish(Math.Floor(num / 1000000.0));
        Console.WriteLine(str1);
        string str2 = Jed.PutSmallIntegerInEnglish(Math.Floor(num % 1000000.0 / 1000.0));
        Console.WriteLine(str2);
        string str3 = Jed.PutSmallIntegerInEnglish(num % 1000.0);
        Console.WriteLine(str3);
        string str4 = string.Empty;
        if (str1 != string.Empty)
            str4 = str1 + "Million ";
        if (str2 != string.Empty)
            str4 = str4 + str2 + "Thousand ";
        return str4 + str3;
    }
}