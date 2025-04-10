using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZXing;

namespace ExcelDnaXP.MyCalss
{
    public class 公用
    {
        public enum BarType
        {
            /// <summary>由美国韦林公司开发，具有较高的纠错能力和编码密度，常用于航空登机牌、票务等场景。</summary>
            AZTEC,

            /// <summary>常用于图书馆、血库和航空快递包裹等场景。它能够对数字 0 - 9、特定字母（A - D）和一些特殊字符进行编码。</summary>
            CODABAR,

            /// <summary>可编码数字、大写字母和部分特殊字符，广泛应用于工业、医疗和物流领域.</summary>
            CODE_39,

            /// <summary>是对 Code 39 的改进，编码密度更高，能表示更多字符。</summary>
            CODE_93,

            /// <summary>能编码全部 128 个 ASCII 字符，编码密度高，常用于物流、仓储和生产制造等行业</summary>
            CODE_128,

            /// <summary>呈正方形或长方形，对印刷质量要求相对较低，可在小面积上编码，常用于电子元器件、医疗器械等产品标识和追溯。</summary>
            DATA_MATRIX,

            /// <summary>由 8 位数字构成，一般用于小尺寸商品，如小零食、小饰品等</summary>
            EAN_8,

            /// <summary>国际通用的商品条码，由 13 位数字组成，广泛用于零售商品标识</summary>
            EAN_13,

            /// <summary>通常用于物流行业，标识运输和配送中的商品</summary>
            ITF,

            /// <summary>主要用于包裹和邮政服务，可存储大量信息。</summary>
            MAXICODE,

            /// <summary>堆叠式二维条码，可编码大量数据，纠错能力强，常用于身份证、护照等证件信息存储</summary>
            PDF_417,

            /// <summary>具有高密度、大容量、纠错能力强等特点，可编码文字、网址、名片信息等，广泛应用于移动支付、广告宣传等领域</summary>
            QR_CODE,

            /// <summary>用于标识商品的不同包装形式，常用于零售行业</summary>
            RSS_14,

            /// <summary>RSS - 14 的扩展版本，能编码更多信息</summary>
            RSS_EXPANDED,

            /// <summary>主要在美国和加拿大使用，由 12 位数字组成，用于零售商品</summary>
            UPC_A,

            /// <summary>是 UPC - A 的压缩版本，用于小包装商品.</summary>
            UPC_E,

            /// <summary>作为 UPC 或 EAN 条码的扩展，用于提供额外信息，像商品的变体、重量等</summary>
            UPC_EAN_EXTENSION
        }
    }
}