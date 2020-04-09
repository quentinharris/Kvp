using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using System.Linq;

namespace VBALibrary
{
    internal static class KvpExtensions
    {
        //public static bool IsNumericDatatype(this object obj)
        //{
        //    switch (Type.GetTypeCode(obj.GetType()))
        //    {
        //        case TypeCode.Byte:
        //        case TypeCode.SByte:
        //        case TypeCode.UInt16:
        //        case TypeCode.UInt32:
        //        case TypeCode.UInt64:
        //        case TypeCode.Int16:
        //        case TypeCode.Int32:
        //        case TypeCode.Int64:
        //        case TypeCode.Decimal:
        //        case TypeCode.Double:
        //        case TypeCode.Single:
        //            return true;

        //        default:
        //            return false;
        //    }
        //}

        public static string ToSeparatedString(this ICollection myColl, string ipSeparator = ",")
        {
            StringBuilder mySb = new StringBuilder();
            foreach (dynamic myItem in myColl)
            {
                string myString;
                try
                {
                    myString = myItem.ToString();
                }
                catch
                {
                    throw new Exception("KvpExtensions:ToSeparatedString: myItem does not have a ToString()");
                }

                if (mySb.Length == 0)
                {
                    mySb.Append(myString);
                }
                else
                {
                    mySb.AppendFormat(ipSeparator + "{0}", myString);
                }
            }
            return mySb.ToString();
        }
    }
}