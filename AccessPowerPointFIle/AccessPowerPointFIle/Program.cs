using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AccessPowerPointFIle.PowerPointAccessSamle;

namespace AccessPowerPointFIle
{
    class Program
    {
        static void Main(string[] args)
        {

            // ファイル読み込みサンプル
            // https://msdn.microsoft.com/ja-jp/library/office/gg278331.aspx
            // GetSlide.getSlideInfo();

            // ファイル書き換えられるか？
            // https://msdn.microsoft.com/ja-jp/library/office/gg278331.aspx
            minorChangeSlides.changeSlideInfo();

            // ファイル作成サンプル
            // MakePresentationFile.CreatePresentation(@"C:\Users\Tetsutaro Yamada\Desktop\PresentationFromFilename.pptx");
        }
    }
}
