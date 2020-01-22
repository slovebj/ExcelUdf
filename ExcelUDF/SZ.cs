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
        [ExcelFunction(Category = "水质函数", Description = "计算加标回收率。**Excel自定义函数**")]
        public static object SZ加标回收(
            [ExcelArgument(Description = "加标前的水样浓度")] double c0,
            [ExcelArgument(Description = "加标后的水样浓度")] double c1,
            [ExcelArgument(Description = "标准溶液的浓度")] double cs,
            [ExcelArgument(Description = "加入的标准溶液体积")] double vs,
            [ExcelArgument(Description = "加标后与水样的总体积")] double v1
            )
        {
            return Math.Round((decimal)((c1 * v1 - c0*(v1 - vs))/cs/vs*100),1)+"%";
        }

        [ExcelFunction(Category = "水质函数", Description = "检测结果整理，小于检出限的用<检出限值表示，大于检出限的保留需要的小数位数。**Excel自定义函数**")]
        public static object SZ检测结果整理(
            [ExcelArgument(Description = "水样检测结果")] double c0,
            [ExcelArgument(Description = "检出限")] double dl,
            [ExcelArgument(Description = "大于检出限时的保留位数")] int digit
            )
        {
            if (c0<dl)
            {
                return "<" + dl;
            }else
            {
                return Math.Round((decimal)c0, digit);
            }
        }

    }

}
