using System;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Microsoft.Office.Core;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Elise.Sources;

namespace PowerpointAutomation
{
    class PowerPointAutomation
    {
        static Application pptApplication;
        static Bible Bible = new Bible(BibleVersions.KJV);

        static Presentation InitializePresentation(string path, bool createNew = true)
        {
            Console.WriteLine($"Initializing presentation {path}...");
            pptApplication = new Application();
            Presentation pptPresentation;

            if (createNew)
                pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoFalse);
            else
                pptPresentation = pptApplication.Presentations.Open(path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            Console.WriteLine($"Creating slide transitions...");
            pptPresentation.SlideMaster.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            pptPresentation.SlideMaster.SlideShowTransition.Duration = 3.5F;
            pptPresentation.SlideMaster.SlideShowTransition.AdvanceOnTime = MsoTriState.msoCTrue;
            pptPresentation.SlideMaster.SlideShowTransition.AdvanceTime = 8.0F;
            return pptPresentation;
        }

        static CustomLayout CustomLayout(ref Presentation presentation, string pictureFile)
        {
            Console.WriteLine($"Creating custom layout with background {pictureFile}...");
            int ct = presentation.SlideMaster.CustomLayouts.Count + 1;
            presentation.SlideMaster.CustomLayouts.Add(ct);
            presentation.SlideMaster.CustomLayouts[ct].FollowMasterBackground = MsoTriState.msoFalse;
            presentation.SlideMaster.CustomLayouts[ct].Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 200, 200, 200, 200);

            if (!string.IsNullOrEmpty(pictureFile))
                presentation.SlideMaster.CustomLayouts[ct].Background.Fill.UserPicture(pictureFile);
            else
                presentation.SlideMaster.CustomLayouts[ct].Background.Fill.BackColor.RGB = 0x0000000;

            return presentation.SlideMaster.CustomLayouts[ct];
        }

        static Microsoft.Office.Interop.PowerPoint.Shape AddAnimatedText(ref Slide slide, string text, int shapeIndex=1)
        {
            if (!string.IsNullOrEmpty(text))
                Console.WriteLine($"Adding animated textrange {text.Substring(0, 10)}...");

            var objText = slide.Shapes[shapeIndex].TextFrame.TextRange;

            slide.Shapes[shapeIndex].TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            slide.Shapes[shapeIndex].TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            slide.Shapes[shapeIndex].TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            slide.Shapes[shapeIndex].TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1F;
            slide.Shapes[shapeIndex].TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;

            objText.Text = text;
            objText.Font.Name = "Georgia";
            objText.Font.Size = 32;
            objText.Font.Color.RGB = 0xFFFFFF;
            objText.Font.Shadow = MsoTriState.msoTrue;

            Console.WriteLine($"Setting up textrange animations...");
            slide.Shapes[shapeIndex].AnimationSettings.EntryEffect = PpEntryEffect.ppEffectFade;
            slide.Shapes[shapeIndex].AnimationSettings.TextUnitEffect = PpTextUnitEffect.ppAnimateByCharacter;
            //slide.Shapes[1].AnimationSettings.TextLevelEffect = PpTextLevelEffect.ppAnimateByFirstLevel;
            //slide.Shapes[1].AnimationSettings.AdvanceMode = PpAdvanceMode.ppAdvanceOnTime;
            //slide.Shapes[1].AnimationSettings.AdvanceTime = 3.0F;
            // slide.Shapes[1].AnimationSettings.AnimateBackground = MsoTriState.msoCTrue;
            slide.Shapes[shapeIndex].AnimationSettings.Animate = MsoTriState.msoCTrue;
            return slide.Shapes[shapeIndex];
        }

        static Slide AddImageSlide(ref Presentation presentation, string pictureFile, string text)
        {
            Console.WriteLine($"Adding new image slide...");
            var customLayout = CustomLayout(ref presentation, pictureFile);
            var slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, customLayout);
            AddAnimatedText(ref slide, text);
            return slide;
        }

        static void Save(ref Presentation presentation, string fileName)
        {
            Console.WriteLine($"Saving presentation to {fileName}...");
            presentation.SaveAs(fileName, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
        }

        static string RandomImage(string directory)
        {
            var random = new Random();
            var images = Directory.GetFiles(directory);
            var index = random.Next(0, images.Count());
            return new FileInfo(images[index]).FullName;
        }

        static string GetText(Bible bible, string verse)
        {
            Reference reference;
            if (bible.TryParseReference(verse, out reference))
                return bible[reference.ToString()];

            return string.Empty;
        }

        static string KJVText(string verse) => GetText(new Bible(), verse);
        //{
        //    Console.WriteLine($"Get text of {verse}...");
        //    Reference reference;

        //    if (Bible.TryParseReference(verse, out reference))
        //        return $"{Bible[reference.ToString()]}\n\t\t- {verse}";

        //    return string.Empty;
        //}

        static string GreekText(string verse) => GetText(new GreekNT(), verse);
        static string HebrewText(string verse) => GetText(new Tanach(), verse);


        public static void Main()
        {
            const string ImageDir = @"C:\Users\mikes\Dropbox\Pictures\Wallpapers";
            var presentation = InitializePresentation(@"C:\users\mikes\dropbox\Test.pptx");

            //var OT = new Tanach(OTSources.WLC);
            //var NT = new GreekNT(NTSources.BYZ);

            var mpPath = @"C:\Users\mikes\Dropbox\TYC\MP\mp.json";
            Console.WriteLine($"Loading prophecies from {mpPath}...");

            using (var file = File.OpenText(mpPath))
            {
                var serializer = new JsonSerializer();
                var prophecies = (Dictionary<string, Dictionary<string, string[]>>)serializer.Deserialize(file, typeof(Dictionary<string, Dictionary<string, string[]>>));
                                    
                foreach (var prophecy in prophecies.Take(10))
                {
                    var OTRef = prophecy.Key;
                    var otText = KJVText(OTRef);

                    if (!string.IsNullOrEmpty(otText))
                    {
                        var slide = AddImageSlide(ref presentation, RandomImage(ImageDir), otText);
                        //AddAnimatedText(ref slide, HebrewText(OTRef), 2);
                    }

                    var subprophecies = prophecy.Value;

                    foreach (var subprophecy in subprophecies)
                    {
                        var name = subprophecy.Key;
                        AddImageSlide(ref presentation, "", name);
                        var NTRefs = subprophecy.Value;

                        foreach (var NTRef in NTRefs)
                        {
                            var kjvText = KJVText(NTRef);

                            if (!string.IsNullOrEmpty(kjvText))
                            {
                                var slide = AddImageSlide(ref presentation, RandomImage(ImageDir), kjvText);
                                //AddAnimatedText(ref slide, GreekText(NTRef), 2);
                            }
                        }
                    }
                }
            }

            Save(ref presentation, @"C:\users\mikes\dropbox\Test.pptx");
            pptApplication.Visible = MsoTriState.msoCTrue;
            // pptPresentation.Close();
            // pptApplication.Quit();
        }
    }
}