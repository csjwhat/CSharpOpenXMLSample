using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace AccessPowerPointFIle.PowerPointAccessSamle
{
    class minorChangeSlides
    {

        public static void changeSlideInfo()
        {
            Console.Write("Please enter a presentation file name without extension: ");
            string fileName = Console.ReadLine();
            string file = @"C:\Users\Tetsutaro Yamada\Desktop\学習\" + fileName + ".pptx";
            int numberOfSlides = CountSlides(file);
            System.Console.WriteLine("Number of slides = {0}", numberOfSlides);
            string slideText;
            for (int i = 0; i < numberOfSlides; i++)
            {
                GetSlideIdAndText(out slideText, file, i);
                System.Console.WriteLine("Slide #{0} contains: {1}", i + 1, slideText);
                changeSlideIdAndText(file, i);
            }
            System.Console.ReadKey();
        }

        public static int CountSlides(string presentationFile)
        {
            // Open the presentation as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                // Pass the presentation to the next CountSlides method
                // and return the slide count.
                return CountSlides(presentationDocument);
            }
        }

        // Count the slides in the presentation.
        public static int CountSlides(PresentationDocument presentationDocument)
        {
            // Check for a null document object.
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            int slidesCount = 0;

            // Get the presentation part of document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            // Get the slide count from the SlideParts.
            if (presentationPart != null)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }
            // Return the slide count to the previous method.
            return slidesCount;
        }

        /// <summary>
        /// ファイル名とスライド番号に対応するテキストをすべて取得して結合して返却する。
        /// </summary>
        /// <param name="sldText">スライドに設定されたString</param>
        /// <param name="docName">ファイル名</param>
        /// <param name="index">スライド番号</param>
        public static void GetSlideIdAndText(out string sldText, string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                // スライド番号から、.NETで扱うためのスライドIDを取得する。
                string relId = (slideIds[index] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                // Build a StringBuilder object.
                StringBuilder paragraphText = new StringBuilder();

                // Get the inner text of the slide:
                // slide.Slide（current element's）がもつ DocumentFormat.OpenXml.Drawing.Text型のXML上の要素をすべて取得する。
                IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
                foreach (A.Text text in texts)
                {
                    paragraphText.Append(text.Text).Append("   ");
                    Console.WriteLine("-------------------------------------");
                    Console.WriteLine("Parent:" + text.LocalName);
                    Console.WriteLine("Parent:" + text.Parent.ToString());
                }
                sldText = paragraphText.ToString();
            }
        }

        public static void changeSlideIdAndText(string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                // スライド番号から、.NETで扱うためのスライドIDを取得する。
                string relId = (slideIds[index] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                // スライド番号からスライドのオブジェクトを取得
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                // Get the inner text of the slide:
                // slide.Slide（current element's）がもつ DocumentFormat.OpenXml.Drawing.Text型のXML上の要素をすべて取得する。
                IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
                foreach (A.Text text in texts)
                {
                    text.Text = text.Text + " Chage";
                    Console.WriteLine("-------------------------------------");
                    Console.WriteLine("Change後テキスト:" + text.Text);
                }
            }
        }

    }
}
