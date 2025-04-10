using ExcelDna.Integration;

namespace xpzy
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

        public static string love(string password)
        {
            return "yq";
        }
    }
}