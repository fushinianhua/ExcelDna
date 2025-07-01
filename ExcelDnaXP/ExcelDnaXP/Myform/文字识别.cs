using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using Tesseract;

namespace Radiant.Myform
{
    public partial class 文字识别 : Form
    {
        private string imagePath = "";
        private string selectedLanguage = "chi_sim";
        private EngineMode selectEngineMode = EngineMode.TesseractOnly;
        private TesseractEngine engine;
        private Bitmap selectImage = null;

        public 文字识别()
        {
            InitializeComponent();
            Comlange.Items.AddRange(new[] { "自动检测", "中文 (chi_sim)", "英文 (eng)", "日文 (jpn)" });
            Comlange.SelectedIndex = 1; // 默认选中中文
            ComEngineMode.Items.AddRange(new[] { "TesseractOnly", "LstmOnly", "TesseractAndLstm", "Defaul" });
            ComEngineMode.SelectedIndex = 1;
        }

        private void 文字识别_Load(object sender, EventArgs e)
        {
        }

        private void 文字识别_FormClosing(object sender, FormClosingEventArgs e)
        {
            engine?.Dispose();
            selectImage?.Dispose();
        }

        public string TesseractOCR(Bitmap image)
        {
            if (engine == null || engine.IsDisposed)
            {
                InitializeEngine();
                if (engine == null) return "OCR引擎初始化失败";
            }

            try
            {
                using (var page = engine.Process(image))
                {
                    string text = page.GetText();
                    float confidence = page.GetMeanConfidence() * 100;
                    Console.WriteLine($"识别置信度: {confidence:F2}%");
                    return text.Replace("\n", " ").Replace("  ", " ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"OCR识别错误: {ex.Message}");
                return "识别失败";
            }
        }

        private void InitializeEngine()
        {
            try
            {
                if (!Directory.Exists("./tessdata"))
                {
                    MessageBox.Show("找不到OCR训练数据，请确保tessdata文件夹在程序目录下。");
                    return;
                }

                engine = new TesseractEngine("./tessdata", selectedLanguage, selectEngineMode);
                // engine.SetVariable("tessedit_char_whitelist", "汉字");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化OCR引擎失败: {ex.Message}");
            }
        }

        private void Comlange_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (Comlange.SelectedIndex)
            {
                case 0: selectedLanguage = ""; break;
                case 1: selectedLanguage = "chi_sim"; break;
                case 2: selectedLanguage = "eng"; break;
                case 3: selectedLanguage = "jpn"; break;
            }

            if (engine != null)
            {
                engine.Dispose();
                InitializeEngine();
            }
        }

        private void ComEngineMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                switch (ComEngineMode.SelectedIndex)
                {
                    case 0: selectEngineMode = EngineMode.TesseractOnly; break;
                    case 1: selectEngineMode = EngineMode.LstmOnly; break;
                    case 2: selectEngineMode = EngineMode.TesseractAndLstm; break;
                    case 3: selectEngineMode = EngineMode.Default; break;
                }

                if (engine != null)
                {
                    engine.Dispose();
                    InitializeEngine();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp";
                    openFileDialog.Multiselect = false;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // 释放旧图片资源
                        if (selectImage != null)
                        {
                            selectImage.Dispose();
                            selectImage = null;
                        }

                        // 加载新图片
                        selectImage = (Bitmap)Image.FromFile(openFileDialog.FileName);
                        pictureBox1.Image = selectImage;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载图片失败: {ex.Message}");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (selectImage == null)
                {
                    MessageBox.Show("请先选择一张图片");
                    return;
                }

                // 创建图片副本进行识别，保留原始图片
                using (var imageCopy = new Bitmap(selectImage))
                {
                    Bitmap bitmap = ConvertToGrayscale(imageCopy);
                    pictureBox1.Image.Dispose();
                    pictureBox1.Image = bitmap;
                    string result = TesseractOCR(bitmap);
                    //textBox1.Text = result;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理图片失败: {ex.Message}");
            }
        }

        private Bitmap ConvertToGrayscale(Bitmap source)
        {
            Bitmap result = new Bitmap(source.Width, source.Height);
            using (Graphics g = Graphics.FromImage(result))
            {
                ColorMatrix colorMatrix = new ColorMatrix(
                    new float[][]
                    {
                new float[] {0.299f, 0.299f, 0.299f, 0, 0},
                new float[] {0.587f, 0.587f, 0.587f, 0, 0},
                new float[] {0.114f, 0.114f, 0.114f, 0, 0},
                new float[] {0, 0, 0, 1, 0},
                new float[] {0, 0, 0, 0, 1}
                    });
                ImageAttributes attributes = new ImageAttributes();
                attributes.SetColorMatrix(colorMatrix);
                g.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height),
                            0, 0, source.Width, source.Height, GraphicsUnit.Pixel, attributes);
            }
            return result;
        }
    }
}