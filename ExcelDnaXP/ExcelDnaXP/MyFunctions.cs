using ExcelDna.Integration;

namespace xpzy
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "返回包含问候语和传入姓名的字符串")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }
    }
}