using System.Reflection;
using System.IO;
using System.Drawing;

namespace Radiant
{
    internal class ResourceHelper
    {
        public static Bitmap GetEmbeddedResourceBitmap(string resourceName, string resourcedir = ".")
        {
            Bitmap image = null;
            //属性-生成操作-嵌入的资源，asm.GetManifestResourceStream("项目命名空间.资源文件所在文件夹名.资源文件名");
            Assembly assm = Assembly.GetExecutingAssembly();
            string extension = Path.GetExtension(resourceName).ToLower();                 //扩展名

            if (!resourcedir.StartsWith(".")) resourcedir = "." + resourcedir;
            if (!resourcedir.EndsWith(".")) resourcedir = resourcedir + ".";
            string sourcename = typeof(ResourceHelper).Assembly.GetName().Name + ".RibbonResources" + resourcedir + resourceName;

            using (Stream ressourceStream = assm.GetManifestResourceStream(sourcename))
            {
                switch (extension)
                {
                    //http://blogs.msdn.com/b/jensenh/archive/2006/11/27/ribbonx-image-faq.aspx
                    case ".ico":
                        image = new Icon(ressourceStream).ToBitmap();
                        break;

                    case ".png":
                    case ".jpg":
                    case ".bmp":
                    default:
                        image = new Bitmap(ressourceStream);
                        image.MakeTransparent();
                        break;
                }
            }
            return image;
        }

        //获取资源文本文件，文件要在属性-生成操作-嵌入资源
        internal static string GetResourceText(string resourceName, string resourcedir = ".")
        {
            string text = string.Empty;

            Assembly assm = Assembly.GetExecutingAssembly();

            if (!resourcedir.StartsWith(".")) resourcedir = "." + resourcedir;
            if (!resourcedir.EndsWith(".")) resourcedir = resourcedir + ".";
            string sourcename = typeof(ResourceHelper).Assembly.GetName().Name + ".MyRibbon" + resourcedir + resourceName;

            using (Stream streamText = assm.GetManifestResourceStream(sourcename))
            {
                using (StreamReader reader = new StreamReader(streamText))
                {
                    text = reader.ReadToEnd();
                }
            }
            return text;
        }
    }
}