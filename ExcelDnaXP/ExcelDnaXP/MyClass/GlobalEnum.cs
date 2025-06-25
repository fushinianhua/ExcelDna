using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Radiant.MyClass
{
    internal class GlobalEnum
    {
        // 考勤类型枚举
        public enum AttendanceType
        {
            事假,
            病假,
            產假,
            返鄉假,
            調休,
            喪假,
            工傷假,
            婚假,
            陪產假,
            年休假,
            春節獎勵假,
            育兒假,
            產檢假,
            節育假,
            公司因素假,
            曠工,
            退TC,
            不加班,
            正常離職,
            離職人數自離
        }

        public enum FillStyle
        {
            不填充,
            单元格填充,
            图片大小填充
        }

        // 填充方向枚举
        public enum FillDirection
        {
            水平方向填充,
            垂直方向填充
        }

        public enum 属性枚举
        {
            改变大小和位置,
            不改变大小和位置,
            改变位置不改变大小
        }

        public class ImageInsertOptions
        {
            /// <summary>图片路径  </summary>
            public string[] SelectedImagePaths { get; set; }

            /// <summary>填充样式枚举</summary>
            public FillStyle FillStyle { get; set; }

            /// <summary> 填充方向枚举</summary>
            public FillDirection FillDirection { get; set; }

            /// <summary>左边距(像素)</summary>
            public double LeftPadding { get; set; }

            /// <summary>上边距(像素)</summary>
            public double TopPadding { get; set; }

            /// <summary>自定义宽度(像素)</summary>
            public double CustomWidth { get; set; }

            /// <summary>自定义高度(像素)</summary>
            public double CustomHeight { get; set; }

            /// <summary>是否链接到文件</summary>
            public bool LinkToFile { get; set; }

            public 属性枚举 图片属性 { get; set; }
            public bool 嵌入单元格 { get; set; }
            public bool 保持纵横比 { get; set; }
        }

        public class 图片尺寸
        {
            public double 宽度 { get; set; }
            public double 高度 { get; set; }
            public double 上边距 { get; set; }
            public double 左边距 { get; set; }
            public double 下边距 { get; set; } = 0;
            public double 右边距 { get; set; } = 0;
            public double 旋转角度 { get; set; } = 0;
            public double 缩放比例 { get; set; } = 0.55;
        }

        /// <summary>
        /// Excel单元格转换
        /// </summary>
        public static class ExcelUnitConverter
        {
            // 默认DPI（大多数显示器）
            private const double DefaultDpi = 96;

            // 字符平均宽度（像素）
            private const double CharWidth = 7;

            // 列宽额外像素（边距）
            private const double ColumnPadding = 2;

            // 列宽转像素
            public static double ColumnWidthToPixels(double columnWidth)
            {
                return columnWidth * CharWidth + ColumnPadding;
            }

            // 像素转列宽
            public static double PixelsToColumnWidth(double pixels)
            {
                return (pixels - ColumnPadding) / CharWidth;
            }

            // 行高转像素
            public static double RowHeightToPixels(double rowHeight, double dpi = DefaultDpi)
            {
                return rowHeight * (dpi / 72);
            }

            // 像素转行高
            public static double PixelsToRowHeight(double pixels, double dpi = DefaultDpi)
            {
                return pixels * (72 / dpi);
            }
        }
    }
}