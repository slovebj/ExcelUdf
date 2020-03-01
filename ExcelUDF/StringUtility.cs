using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;
using IExcel = Microsoft.Office.Interop.Excel;
namespace ExcelUDF
{
    public partial class ExcelUDF
    {


        [ExcelFunction(Category = "文本处理_提取替换", Description = "提取指定字符。**Excel自定义函数**")]
        public static object WB提取指定字符(
            [ExcelArgument(Description = "待查找提取的字符串")] string inputString,
            [ExcelArgument(Description = "查找用于提取的指定字符，多个字符用逗号隔开")] string matchString,
            [ExcelArgument(Description = "当提取多个结果时，结果之间的间隔符")] string splitStr
        )
        {

            var patterns = matchString.Split(new char[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Where(s => s.Length > 1).ToList();
            string patSingle = string.Join("", matchString.Split(new char[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Where(s => s.Length == 1));
            patterns.Add($"[{patSingle}]+");
            var pattern = string.Join("|", patterns);


            return RegMatchValue(inputString, pattern, splitStr);
        }
        [ExcelFunction(Category = "文本处理_提取替换", Description = "提取中文。**Excel自定义函数**")]
        public static object WB提取中文(
            [ExcelArgument(Description = "待查找提取的字符串")] string inputString,
            [ExcelArgument(Description = "当提取多个结果时，结果之间的间隔符")] string splitStr
                 )

        {
            string pattern = "[\u4e00-\u9fa5]+";
            return RegMatchValue(inputString, pattern, splitStr);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "提取数字。**Excel自定义函数**")]
        public static object WB提取数字(
                [ExcelArgument(Description = "待查找提取的字符串")] string inputString,
                [ExcelArgument(Description = "当提取多个结果时，结果之间的间隔符")] string splitStr
         )

        {
            string pattern = @"[0-9][0-9,]*\.[0-9]+|[0-9][0-9,]*";
            return RegMatchValue(inputString, pattern, splitStr);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "提取英文字母。**Excel自定义函数**")]
        public static object WB提取英文(
        [ExcelArgument(Description = "待查找提取的字符串")] string inputString,
        [ExcelArgument(Description = "当提取多个结果时，结果之间的间隔符")] string splitStr
                )

        {
            string pattern = @"[a-zA-Z]+";
            return RegMatchValue(inputString, pattern, splitStr);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "替换英文。**Excel自定义函数**")]
        public static object WB替换英文(
                [ExcelArgument(Description = "待查找替换的字符串")] string inputString,
                [ExcelArgument(Description = "查找到的字符串替换为此字符串，不输入默认替换为空")] string replaceString
        )

        {
            string pattern = @"[a-zA-Z]+";
            return RegReplaceValue(inputString, pattern, replaceString);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "替换中文。**Excel自定义函数**")]
        public static object WB替换中文(
        [ExcelArgument(Description = "待查找替换的字符串")] string inputString,
        [ExcelArgument(Description = "查找到的字符串替换为此字符串，不输入默认替换为空")] string replaceString
                )
        {
            string pattern = @"[\u4e00-\u9fa5]+";
            return RegReplaceValue(inputString, pattern, replaceString);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "替换中文。**Excel自定义函数**")]
        public static object WB替换数字(
            [ExcelArgument(Description = "待查找替换的字符串")] string inputString,
            [ExcelArgument(Description = "查找到的字符串替换为此字符串，不输入默认替换为空")] string replaceString
             )
        {
            string pattern = @"[0-9][0-9,]*\.[0-9]+|[0-9][0-9,]*";
            return RegReplaceValue(inputString, pattern, replaceString);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "替换指定字符。**Excel自定义函数**")]
        public static object WB替换指定字符(
            [ExcelArgument(Description = "待查找替换的字符串")] string inputString,
            [ExcelArgument(Description = "查找用于替换的指定字符，多个字符用逗号隔开")] string matchString,
            [ExcelArgument(Description = "查找到的字符串替换为此字符串，不输入默认替换为空")] string replaceString
                )
        {

            var patterns = matchString.Split(new char[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Where(s => s.Length > 1).ToList();
            string patSingle = string.Join("", matchString.Split(new char[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Where(s => s.Length == 1));
            patterns.Add($"[{patSingle}]+");
            var pattern = string.Join("|", patterns);

            return RegReplaceValue(inputString, pattern, replaceString);
        }

        private static string RegReplaceValue(string inputString, string pattern, string replaceString)
        {
            RegexOptions options = RegexOptions.Multiline;
            return Regex.Replace(inputString, pattern, replaceString, options);
        }

        private static string RegMatchValue(string inputString, string pattern, string splitStr)
        {
            RegexOptions options = RegexOptions.Multiline;
            MatchCollection matches = Regex.Matches(inputString, pattern, options);
            return string.Join(splitStr, matches.Cast<Match>().Where(s => !string.IsNullOrEmpty(s.Value)));
        }

        [ExcelFunction(Category = "文本处理", Description = "字符串去除重复字符。**Excel自定义函数**")]
        public static object WB去重(
            [ExcelArgument(Description = "待去重的字符串")] string inputString
                )

        {
            return string.Join("", inputString.ToArray().Distinct());
        }

        [ExcelFunction(Category = "文本处理", Description = "字符串反转。**Excel自定义函数**")]
        public static object WB反转(
            [ExcelArgument(Description = "待反转的字符串")] string inputString
        )

        {
            return string.Join("", inputString.ToArray().Reverse());
        }

        [ExcelFunction(Category = "文本处理", Description = "字符排序。**Excel自定义函数**")]
        public static object WB排序(
            [ExcelArgument(Description = "待排序的字符串")] string inputString,
            [ExcelArgument(Description = "是否降序排列，默认为升序，TRUE为降序，FALSE为升序")] bool isDesc
            )

        {
            if (isDesc)
            {
                return string.Join("", inputString.ToArray().OrderByDescending(s => s));
            }
            else
            {
                return string.Join("", inputString.ToArray().OrderBy(s => s));
            }

        }

        [ExcelFunction(Category = "文本处理", Description = "字符串分解为单个字符存放到多个单元格。对于小数金额，将抛弃小数点，并将小数部分用0补足两位，主要用于财务用途。**Excel自定义函数**")]
        public static object WB分解展开(
            [ExcelArgument(Description = "需要分解的原始字符串")] string inputString,
            [ExcelArgument(Description = "拆分的总行/列数。列数少于字符串长度可能会有部分字符被抛弃")] int colNum,
            [ExcelArgument(Description = "少于指定总行/列数是否前面以字符填补。默认为空格")] string padStr = " ",
            [ExcelArgument(Description = "输入H为按行横向展开，输入L为按列纵向展开,默认为H")] string optAlignHorL = "H"
            )
        {
            if (inputString.Contains("."))//含小数点
            {
                inputString = inputString.Split('.')[0] + inputString.Split('.')[1].PadRight(2, '0');
            }
            inputString = inputString.PadLeft(colNum, Convert.ToChar(padStr.ToString().Trim().Substring(0, 1)));
            inputString = inputString.Substring(inputString.Length - colNum, colNum);

            return Common.ReturnDataArray(inputString.ToArray().Select(s => s.ToString().Replace(" ", "")).ToArray(), optAlignHorL);
        }

        [ExcelFunction(Category = "文本处理", Description = "根据传入的分割符，对字符串进行分割操作，返回多值。**Excel自定义函数**")]
        public static object WB分割展开(
            [ExcelArgument(Description = "待分割的字符串")] string inputString,
            [ExcelArgument(Description = "分隔符，可以引用多个连续单元格或以英文逗号分隔的一个字符串")] object delimiter,
            [ExcelArgument(Description = "返回多值时输入H为按行横向展开，输入L为按列纵向展开")] string optAlignHorL
            )
        {
            var splitList = Common.GetSplitStringList(delimiter);
            var result = inputString.Split(splitList.ToArray(), StringSplitOptions.RemoveEmptyEntries);
            return Common.ReturnDataArray(result, optAlignHorL);
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "根据传入的分割符，对字符串进行分割操作，提取第N个值。**Excel自定义函数**")]
        public static object WB分割提取(
            [ExcelArgument(Description = "待分割的字符串")] string inputString,
            [ExcelArgument(Description = "分隔符，可以引用多个连续单元格或以英文逗号分隔的一个字符串")] object delimiter,
            [ExcelArgument(Description = "提取分割后的第N个值，从1开始计数")] int returnNum
             )
        {
            var splitList = Common.GetSplitStringList(delimiter);
            var result = inputString.Split(splitList.ToArray(), StringSplitOptions.RemoveEmptyEntries);
            if (returnNum <= result.Length && returnNum > 0)
            {
                return result[returnNum - 1];
            }
            else
            {
                return "";
            }
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "从字符串两端清除特定字符。**Excel自定义函数**")]
        public static object WB两端清除(
            [ExcelArgument(Description = "待两端清除内容的字符串")] string inputString,
            [ExcelArgument(Description = "输入需要清除的字符，可以引用多个连续单元格或以英文逗号分隔的单个字符串")] object trimValues
            )
        {
            var trimList = Common.GetSplitStringList(trimValues);
            return inputString.Trim(trimList.Select(s => Convert.ToChar(s.Trim().Substring(0, 1))).ToArray());
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "从字符串前端清除特定字符。**Excel自定义函数**")]
        public static object WB前端清除(
            [ExcelArgument(Description = "待前端清除内容的字符串")] string inputString,
            [ExcelArgument(Description = "输入需要清除的字符，可以引用多个连续单元格或以英文逗号分隔的单个字符串")] object trimValues
                 )
        {

            var trimList = Common.GetSplitStringList(trimValues);
            return inputString.TrimStart(trimList.Select(s => Convert.ToChar(s.Trim().Substring(0, 1))).ToArray());
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "从字符串末端清除特定字符。**Excel自定义函数**")]
        public static object WB末端清除(
            [ExcelArgument(Description = "待末端清除内容的字符串")] string inputString,
            [ExcelArgument(Description = "输入需要清除的字符串，可以引用多个连续单元格或以英文逗号分隔的一个ASCII字符串")] object trimValues
                )
        {
            var trimList = Common.GetSplitStringList(trimValues);
            return inputString.TrimEnd(trimList.Select(s => Convert.ToChar(s.Trim().Substring(0, 1))).ToArray());
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "前端填充指定字符，使字符串达到指定长度。**Excel自定义函数**")]
        public static object WB前端填充(
            [ExcelArgument(Description = "待填充的字符串")] string inputString,
            [ExcelArgument(Description = "前端填充的单个字符")] object padStr,
            [ExcelArgument(Description = "填充后字符串的总字符数，数字小于待填充的字符串长度时返回原字符串")] int strLen
         )

        {
            return inputString.PadLeft(strLen, Convert.ToChar(padStr.ToString().Trim().Substring(0, 1)));

        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "末端填充指定字符，使字符串达到指定长度。**Excel自定义函数**")]
        public static object WB末端填充(
            [ExcelArgument(Description = "待填充的字符串")] string inputString,
            [ExcelArgument(Description = "末端填充的单个字符")] object padStr,
            [ExcelArgument(Description = "填充后字符串的总字符数，数字小于待填充的字符串长度时返回原字符串")] int strLen
                 )
        {
            return inputString.PadRight(strLen, Convert.ToChar(padStr.ToString().Trim().Substring(0, 1)));
        }
        [ExcelFunction(Category = "文本处理",IsThreadSafe =true, Description = "字符串拼接函数。通过第三个可选参数，可以给每个需要拼接的字符串首尾附加上特定的符号或文字。**Excel自定义函数**")]
        public static string WB附加拼接(
            [ExcelArgument(Description = "要拼接的字符串区域")] object StringJoinRange,
            [ExcelArgument(Description = "分隔符，如：,*+等")] string delimiter,
            [ExcelArgument(Description = "附加在每个拼接字符串前端的符号或文字，如双引号、书名号等。可选")] string strSurround1,
            [ExcelArgument(Description = "附加在每个拼接字符串末端的符号或文字，如双引号、书名号等。可选")] string strSurround2)
        {

            List<object> valuesArr = new List<object>();
            Common.AddValueToList(StringJoinRange, ref valuesArr);
            return string.Join(delimiter, valuesArr.Select(s=> strSurround1+s.ToString()+ strSurround2));
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "字符串拼接函数。查找区域符合特定条件后，对应的拼接字符串区域用指定分隔符进行拼接。相对自带TEXTJOIN函数增加了条件查找功能。类似SUMIF、COUNTIF。**Excel自定义函数**")]
        public static object WB条件拼接(
                [ExcelArgument(Description = "查找区域，多列引用请使用FZGetMultiColRange函数输入")] object[,] lookupRange,
                [ExcelArgument(Description = "条件，用于验证查找区域是否符合")] string criteria,
                [ExcelArgument(Description = "拼接字符串区域，多列引用请使用FZGetMultiColRange函数输入")] object[,] StringJoinRange,
                [ExcelArgument(Description = "分隔符，如：,*+等")] string delimiter,
                [ExcelArgument(Description = "是否精确匹配，默认为否。TRUE为查找区域值与条件字符串完全相等，FALSE为包含即可")] bool isExactMatch = false)
        {
            //输入不规范，不是单列
            if (lookupRange.GetLength(0) != StringJoinRange.GetLength(0) || lookupRange.GetLength(1) != 1 || StringJoinRange.GetLength(1) != 1)
            {
                return ExcelError.ExcelErrorNA;
            }

            List<string> list = new List<string>();

            for (int i = 0; i < lookupRange.GetLength(0); i++)
            {
                if (isExactMatch)
                {
                    if (lookupRange.GetValue(i, 0).ToString() == criteria)
                    {
                        list.Add(StringJoinRange.GetValue(i, 0).ToString());
                    }

                }
                else
                {
                    if (lookupRange.GetValue(i, 0).ToString().Contains(criteria))
                    {
                        list.Add(StringJoinRange.GetValue(i, 0).ToString());
                    }
                }
            }

            return string.Join(delimiter, list);

        }

    }
}
