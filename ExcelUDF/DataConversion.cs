using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace ExcelUDF
{
    public partial class ExcelUDF
    {

        [ExcelFunction(Category = "数据转换", Description = "修约函数。四舍六入五成双。**Excel自定义函数**")]
        public static object Round5(
            [ExcelArgument(Description = "需要修约的数字")] double num,
            [ExcelArgument(Description = "小数保留位数")] int digits
         )
        {
            return Math.Round((decimal)num, digits);
        }


        [ExcelFunction(Category = "数据转换", Description = "经纬度转换：度分秒--->数字。**Excel自定义函数**")]
        public static object ZH度分秒到数字(
        [ExcelArgument(Description = "度分秒经纬度")] string degrees,
        [ExcelArgument(Description = "转换后保留的小数位数，默认为6")] int digits = 6
        )
        {
            const double num = 60;
            double digitalDegree = 0.0;
            int d = degrees.IndexOf('°');           //度的符号对应的 Unicode 代码为：00B0[1]（六十进制），显示为°。
            if (d < 0)
            {
                return digitalDegree;
            }
            string degree = degrees.Substring(0, d);
            digitalDegree += Convert.ToDouble(degree);

            int m = degrees.IndexOf('′');           //分的符号对应的 Unicode 代码为：2032[1]（六十进制），显示为′。
            if (m < 0)
            {
                return digitalDegree;
            }
            string minute = degrees.Substring(d + 1, m - d - 1);
            digitalDegree += ((Convert.ToDouble(minute)) / num);

            int s = degrees.IndexOf('″');           //秒的符号对应的 Unicode 代码为：2033[1]（六十进制），显示为″。
            if (s < 0)
            {
                return digitalDegree;
            }
            string second = degrees.Substring(m + 1, s - m - 1);
            digitalDegree += (Convert.ToDouble(second) / (num * num));

            return Math.Round(digitalDegree, digits);

        }

        [ExcelFunction(Category = "数据转换", Description = "经纬度转换：小数--->度分秒。**Excel自定义函数**")]
        public static object ZH数字到度分秒(
        [ExcelArgument(Description = "数字经纬度")] double digitalDegree,
        [ExcelArgument(Description = "转换后秒保留的小数位数，默认为0")] int digits = 0
        )
        {
            const double num = 60;
            int degree = (int)digitalDegree;
            double tmp = (digitalDegree - degree) * num;
            int minute = (int)tmp;
            double second = Math.Round((tmp - minute) * num, digits);
            string degrees = "" + degree + "°" + minute + "′" + second + "″";
            return degrees;

        }

        [ExcelFunction(Category = "数据转换", Description = "不同进制数间的转换。**Excel自定义函数**")]
        public static object ZH进制(
            [ExcelArgument(Description = "输入待转换的值")] string input,
            [ExcelArgument(Description = "输入值的进制数")] int fromType,
            [ExcelArgument(Description = "需要转换的进制数")] int toType
            )
        {

            string output = input;
            switch (fromType)
            {
                case 2:
                    output = ConvertGenericBinaryFromBinary(input, toType);
                    break;
                case 8:
                    output = ConvertGenericBinaryFromOctal(input, toType);
                    break;
                case 10:
                    output = ConvertGenericBinaryFromDecimal(input, toType);
                    break;
                case 16:
                    output = ConvertGenericBinaryFromHexadecimal(input, toType);
                    break;
                default:
                    break;
            }
            return output;

        }

        #region 进制转换中间函数
        /// <summary>
        /// 从二进制转换成其他进制
        /// </summary>
        /// <param name="input"></param>
        /// <param name="toType"></param>
        /// <returns></returns>
        private static string ConvertGenericBinaryFromBinary(string input, int toType)
        {
            switch (toType)
            {
                case 8:
                    //先转换成十进制然后转八进制
                    input = Convert.ToString(Convert.ToInt32(input, 2), 8);
                    break;
                case 10:
                    input = Convert.ToInt32(input, 2).ToString();
                    break;
                case 16:
                    input = Convert.ToString(Convert.ToInt32(input, 2), 16);
                    break;
                default:
                    break;
            }
            return input;
        }

        /// <summary>
        /// 从八进制转换成其他进制
        /// </summary>
        /// <param name="input"></param>
        /// <param name="toType"></param>
        /// <returns></returns>
        private static string ConvertGenericBinaryFromOctal(string input, int toType)
        {
            switch (toType)
            {
                case 2:
                    input = Convert.ToString(Convert.ToInt32(input, 8), 2);
                    break;
                case 10:
                    input = Convert.ToInt32(input, 8).ToString();
                    break;
                case 16:
                    input = Convert.ToString(Convert.ToInt32(input, 8), 16);
                    break;
                default:
                    break;
            }
            return input;
        }

        /// <summary>
        /// 从十进制转换成其他进制
        /// </summary>
        /// <param name="input"></param>
        /// <param name="toType"></param>
        /// <returns></returns>
        private static string ConvertGenericBinaryFromDecimal(string input, int toType)
        {
            string output = "";
            int intInput = Convert.ToInt32(input);
            switch (toType)
            {
                case 2:
                    output = Convert.ToString(intInput, 2);
                    break;
                case 8:
                    output = Convert.ToString(intInput, 8);
                    break;
                case 16:
                    output = Convert.ToString(intInput, 16);
                    break;
                default:
                    output = input;
                    break;
            }
            return output;
        }

        /// <summary>
        /// 从十六进制转换成其他进制
        /// </summary>
        /// <param name="input"></param>
        /// <param name="toType"></param>
        /// <returns></returns>
        private static string ConvertGenericBinaryFromHexadecimal(string input, int toType)
        {
            switch (toType)
            {
                case 2:
                    return Convert.ToString(Convert.ToInt32(input, 16), 2);

                case 8:
                    return Convert.ToString(Convert.ToInt32(input, 16), 8);

                case 10:
                    return Convert.ToInt32(input, 16).ToString();

                default:
                    return string.Empty;
            }
        }
        #endregion 进制转换中间函数

        [ExcelFunction(Category = "数据转换", Description = "不同颜色间表示方法间的转换。**Excel自定义函数**")]
        public static object ZH颜色RGB到Html(
                    [ExcelArgument(Description = "输入R值，范围0-255")] object inputR,
                    [ExcelArgument(Description = "输入G值，范围0-255")] object inputG,
                    [ExcelArgument(Description = "输入B值，范围0-255")] object inputB
                     )
        {
            try
            {
                int R = Convert.ToInt32(inputR.ToString().Trim());
                int G = Convert.ToInt32(inputG.ToString().Trim());
                int B = Convert.ToInt32(inputB.ToString().Trim());

                return ColorTranslator.ToHtml(Color.FromArgb(255, R, G, B));

            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }



        }

        [ExcelFunction(Category = "数据转换", Description = "不同颜色间表示方法间的转换。**Excel自定义函数**")]
        public static object ZH颜色RGB到Ole(
            [ExcelArgument(Description = "输入R值，范围0-255")] object inputR,
            [ExcelArgument(Description = "输入G值，范围0-255")] object inputG,
            [ExcelArgument(Description = "输入B值，范围0-255")] object inputB
             )
        {
            try
            {
                int R = Convert.ToInt32(inputR.ToString().Trim());
                int G = Convert.ToInt32(inputG.ToString().Trim());
                int B = Convert.ToInt32(inputB.ToString().Trim());

                return ColorTranslator.ToOle(Color.FromArgb(255, R, G, B));
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "数据转换", Description = "不同颜色间表示方法间的转换。**Excel自定义函数**")]
        public static object ZH颜色Ole到RGB(
                [ExcelArgument(Description = "输入Ole值，OFFICE软件的Color属性")] int inputOle
                     )
        {
            try
            {
               Color color= ColorTranslator.FromOle(inputOle);
                return $"{color.R},{color.G},{color.B}"; 
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "数据转换", Description = "不同颜色间表示方法间的转换。**Excel自定义函数**")]
        public static object ZH颜色Ole到Html(
        [ExcelArgument(Description = "输入Ole值，OFFICE软件的Color属性")] int inputOle
             )
        {
            try
            {
                Color color = ColorTranslator.FromOle(inputOle);
                
                return ColorTranslator.ToHtml(Color.FromArgb(255,color.R,color.G,color.B));
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "数据转换", Description = "不同颜色间表示方法间的转换。**Excel自定义函数**")]
        public static object ZH颜色Html到RGB(
                [ExcelArgument(Description = "输入网页Html格式颜色值，由#开头")] string inputHtmlColor
                         )
        {
            try
            {
                Color color = ColorTranslator.FromHtml(inputHtmlColor);
                return $"{color.R},{color.G},{color.B}";
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "数据转换", Description = "不同颜色间表示方法间的转换。**Excel自定义函数**")]
        public static object ZH颜色Html到Ole(
        [ExcelArgument(Description = "输入网页Html格式颜色值，由#开头")] string inputHtmlColor
                 )
        {
            try
            {
                Color color = ColorTranslator.FromHtml(inputHtmlColor);
                return ColorTranslator.ToOle(color);
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "数据转换", Description = "Unix timestamp转普通日期。**Excel自定义函数**")]
        public static object ZH时间戳到日期(
           [ExcelArgument(Description = "输入UnixTimestamp")] Int64 inputUnixTimestamp)

        {
            if (inputUnixTimestamp.ToString().Length == 10)
            {
                inputUnixTimestamp = inputUnixTimestamp * 1000;
            }
            System.DateTime time = System.DateTime.MinValue;
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
            time = startTime.AddMilliseconds(inputUnixTimestamp);
            Common.ChangeNumberFormat("yyyy-mm-dd hh:mm:ss");
            return time;

        }

        [ExcelFunction(Category = "数据转换", Description = "普通日期转Unix timestamp。**Excel自定义函数**")]
        public static object ZH日期到时间戳(
            [ExcelArgument(Description = "输入UnixTimestamp")] DateTime inputDateTime,
             [ExcelArgument(Description = "是否精确到秒，TRUE为秒，FALSE为毫秒")] bool isSecond
            )
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1, 0, 0, 0, 0));
            //intResult = (time- startTime).TotalMilliseconds;
            long unixTime = (inputDateTime.Ticks - startTime.Ticks) / 10000;            //除10000调整为13位
            Common.ChangeNumberFormat("0");
            return isSecond ? unixTime / 1000 : unixTime;
        }

        [ExcelFunction(Category = "数据转换", Description = "数字转万为单位。**Excel自定义函数**")]
        public static object ZH数字到万(
           [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
           [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits,
           [ExcelArgument(Description = "是否需要带上万字样的数字格式")] bool isNumberFormatWan
           )
        {
            double ratio = 0.0001;
            if (isNumberFormatWan)
            {
                string numString = num_digits is ExcelMissing ? (new string('0', 2)) : (new string('0', Convert.ToInt32(num_digits)));
                Common.ChangeNumberFormat($"#,##0.{numString}万;-#,##0.{numString}万");
            }
            return ConvertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "数据转换", Description = "海里转千米。**Excel自定义函数**")]
        public static object ZH海里到千米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 1.852;
            return ConvertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "数据转换", Description = "英里转千米。**Excel自定义函数**")]
        public static object ZH英里到千米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 1.6093;
            return ConvertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "数据转换", Description = "米转英寸。**Excel自定义函数**")]
        public static object ZH米到英寸(
               [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
               [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 39.37;
            return ConvertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "数据转换", Description = "英寸转米。**Excel自定义函数**")]
        public static object ZH英寸到米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 0.0254;
            return ConvertByRatio(ratio, inputNumber, num_digits);
        }

        private static object ConvertByRatio(double ratio, double inputNumber, object num_digits)
        {
            if (num_digits is ExcelMissing)
            {
                return ratio * inputNumber;
            }
            else
            {
                return Math.Round((decimal)(ratio * inputNumber), Convert.ToInt32(num_digits), MidpointRounding.AwayFromZero);
            }
        }
    }

}
