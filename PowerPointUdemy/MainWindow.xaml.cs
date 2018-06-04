using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using NAudio.Lame;
using NAudio.Wave;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPointUdemy.YoutubeModels;

namespace PowerPointUdemy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        string openFileDialogStartFolder = "E:\\PowerPoint";
        string speechFilesOutputFolder = @"E:\PowerPoint\Read";


        private bool _IsBusy;
        private string _imagesFolderPath;
        private string _soundFilePath;
        private string _textFilePath;
        private string _outputFolderPath;
        private bool _generateVideoAutomatic;
        private bool _embeddedSpeechSupport;

        List<float>[] textBoxLocationArray = null;

        public string VideoId { get; set; }

        public bool IsBusy
        {
            get { return _IsBusy; }
            set
            {
                if (value != null || value != _IsBusy) _IsBusy = value;
                OnPropertyChanged("IsBusy");
            }
        }
        public string imagesFolderPath
        {
            get { return _imagesFolderPath; }
            set
            {
                if (value != null || value != _imagesFolderPath) _imagesFolderPath = value;
                OnPropertyChanged("imagesFolderPath");
            }
        }
        public string soundFilePath
        {
            get { return _soundFilePath; }
            set
            {
                if (value != null || value != _soundFilePath)
                    _soundFilePath = value;
                OnPropertyChanged("soundFilePath");
            }
        }
        public string textFilePath
        {
            get { return _textFilePath; }
            set
            {
                if (value != null || value != _textFilePath) _textFilePath = value;
                OnPropertyChanged("textFilePath");
            }
        }
        public string outputFolderPath
        {
            get { return _outputFolderPath; }
            set
            {
                if (value != null || value != _outputFolderPath) _outputFolderPath = value;
                OnPropertyChanged("outputFolderPath");
            }
        }

        public bool generateVideoAutomatic
        {
            get { return _generateVideoAutomatic; }
            set
            {
                if (value != null || value != _generateVideoAutomatic) _generateVideoAutomatic = value;
                OnPropertyChanged("generateVideoAutomatic");
            }
        }
        public bool embeddedSpeechSupport
        {
            get { return _embeddedSpeechSupport; }
            set
            {
                if (value != null || value != _embeddedSpeechSupport) _embeddedSpeechSupport = value;
                OnPropertyChanged("embeddedSpeechSupport");
            }
        }

        string pptTest = string.Empty;
        string pptTestVideo = string.Empty;

        private string[] images;

        private string nextPageToken = string.Empty;

        private Microsoft.Office.Interop.PowerPoint.Application pptApplication;
        private Microsoft.Office.Interop.PowerPoint.Slides slides;
        private Microsoft.Office.Interop.PowerPoint._Slide slide;
        private Microsoft.Office.Interop.PowerPoint.TextRange objText;
        private Presentation pptPresentation;
        private int shapeId;
        private bool isLastSlide = false;

        private string comment = string.Empty;
        private string author = string.Empty;
        private string likes = "0";
        private string image = string.Empty;
        private bool isNewPresentation = true;
        private Random rnd;

        private List<YoutubeModels.Comments.Item> totalComments;

        private List<int> _CommentsCount;

        public int NumberOfCommentsWanted { get; set; }


        public List<int> CommentsCount
        {
            get { return _CommentsCount; }
            set
            {
                if (value != null || value != _CommentsCount) _CommentsCount = value;
                OnPropertyChanged("CommentsCount");
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            CommentsCount = new List<int>();
            for (int i = 1; i < 100; i++)
                CommentsCount.Add(i);
           
            rnd = new Random();
            this.DataContext = this;
            textBoxLocationArray = new List<float>[3];

        }

        private async void OnGenerateClicked(object sender, RoutedEventArgs e)
        {
            this.totalComments = new List<YoutubeModels.Comments.Item>();

            this.IsBusy = true;
            pptTest = $"{this.outputFolderPath}\\test.pptx";
            pptTestVideo = $"{outputFolderPath}\\test.mp4";


            if (string.IsNullOrEmpty(imagesFolderPath) || string.IsNullOrEmpty(soundFilePath)
                || string.IsNullOrEmpty(textFilePath) || string.IsNullOrEmpty(outputFolderPath))
            {
                this.IsBusy = false;
                MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show
                    ("You have to select Images folder path, Sound file, text file and output folder path",
                    "Please correct the errors before you can generate the PPT",
                    MessageBoxButton.OK, MessageBoxImage.Error);

                return;
            }

            try
            {
                await Task.Run(() =>
                {
                    var urlTopLevelCommWithReplies = Constants.urlTopLevelCommentPrefix + this.VideoId + Constants.urlTopLevelCommentsPostfix;

                    Stream streamTopLevelCommWithReplies = GetStreamFromUrl(urlTopLevelCommWithReplies);
                    using (streamTopLevelCommWithReplies)
                    {
                        var resTopLevelCommWithReplies = GetRootObjFromStream<YoutubeModels.Comments.Rootobject>(streamTopLevelCommWithReplies);

                        this.nextPageToken = resTopLevelCommWithReplies.nextPageToken;
                        this.totalComments.AddRange(resTopLevelCommWithReplies.items.ToList());

                        var date1 = resTopLevelCommWithReplies.items.LastOrDefault().snippet.topLevelComment.snippet.publishedAt;

                        var date2 = DateTime.Now;

                        //while (!string.IsNullOrEmpty(this.nextPageToken))
                        //{
                        //    //if (date2 - date1.Date >= TimeSpan.FromDays(1))
                        //    //    break;

                        //    urlTopLevelCommWithReplies = Constants.urlTopLevelCommentPrefix + this.VideoId +
                        //        $"&part=snippet&order=relevance&maxResults=100&pageToken={this.nextPageToken}&key={Constants.key}";

                        //    Stream streamTopLevelCommWithRepliesMorePages = GetStreamFromUrl(urlTopLevelCommWithReplies);
                        //    using (streamTopLevelCommWithRepliesMorePages)
                        //    {
                        //        resTopLevelCommWithReplies = GetRootObjFromStream<YoutubeModels.Comments.Rootobject>(streamTopLevelCommWithRepliesMorePages);

                        //        this.nextPageToken = resTopLevelCommWithReplies.nextPageToken;

                        //        this.totalComments.AddRange(resTopLevelCommWithReplies.items.ToList());
                        //        date1 = resTopLevelCommWithReplies.items.LastOrDefault().snippet.topLevelComment.snippet.publishedAt.Date;
                        //        date2 = DateTime.UtcNow.Date;
                        //    }
                        //}


                        // Comments ordered by time.
                        // https://www.googleapis.com/youtube/v3/commentThreads?part=snippet&maxResults=5&order=time&videoId=xO8Cz-9qKTI&key={YOUR_API_KEY}

                        //var top5LikedCommentsToday = Constants.totalComments
                        //    .Where(c => (DateTime.UtcNow.Date - c.snippet.topLevelComment.snippet.publishedAt.Date <= TimeSpan.FromDays(1)))
                        //    .OrderByDescending(c => c.snippet.topLevelComment.snippet.likeCount).Take(5);

                        //var top5LikedCommentsToday = this.totalComments.OrderByDescending(c => c.snippet.topLevelComment.snippet.likeCount).Take(5);

                        var top5LikedCommentsToday = this.totalComments.Take(NumberOfCommentsWanted);

                        if (File.Exists(@"E:\PowerPoint\Comments.txt"))
                        {
                            File.WriteAllText(@"E:\PowerPoint\Comments.txt", String.Empty);
                        }

                        foreach (var commTopLevel in top5LikedCommentsToday)
                        {
                            using (StreamWriter sw = new StreamWriter(textFilePath, true))
                            {
                                sw.WriteLine($"Author:{commTopLevel.snippet.topLevelComment.snippet.authorDisplayName}");
                                sw.WriteLine($"CommentTopLevel:{commTopLevel.snippet.topLevelComment.snippet.textOriginal}");
                                sw.WriteLine($"Likes:{commTopLevel.snippet.topLevelComment.snippet.likeCount}");
                            }
                        }
                    }

                    images = Directory.GetFiles(imagesFolderPath, "*.*");


                    Constants.textBox1Width = 500;
                    Constants.textBox1Height = 100;

                    if (File.Exists(pptTest) && !this.IsFileinUse(new FileInfo(pptTest)))
                    {
                        File.Delete(pptTest);
                    }

                    if (File.Exists(pptTestVideo) && !this.IsFileinUse(new FileInfo(pptTestVideo)))
                    {
                        File.Delete(pptTestVideo);
                    }

                    Directory.CreateDirectory(speechFilesOutputFolder);

                    Array.ForEach(Directory.GetFiles(@"E:\PowerPoint\Read"),
                         delegate (string path) { File.Delete(path); });

                    string[] lines = File.ReadAllLines(textFilePath);

                    int j = 0;
                    int k = 0;
                    for (int i = 0; i < lines.Length; i = i + 3)
                    {

                        if (lines[i].Substring(0, 7) == "Author:")
                        {
                            this.author = lines[i].Substring(7, lines[i].Length - 7);
                            this.comment = lines[i + 1].Substring(16, lines[i + 1].Length - 16);
                            this.likes = lines[i + 2].Substring(6, lines[i + 2].Length - 6);

                            if (embeddedSpeechSupport)
                                ReadAndSaveToMP3(this.comment, ++j);
                        }

                        if (i == lines.Length - 3)
                            this.isLastSlide = true;

                        ReadTextAndImagesToMakePPT(this.author, this.comment, int.Parse(this.likes), ++k);
                    }

                    ReduceBackGroundMusicVolume();
                });

                BestCommentsTextBox.Text = File.ReadAllText(textFilePath);

            }
            catch (Exception ex)
            {
                this.IsBusy = false;
                MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show
                                ("Error",
                                "Some error occurred, please restart the app",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            this.IsBusy = false;

            if (generateVideoAutomatic)
                CreateVideoFromPPT();

            this.pptPresentation.Close();
            this.pptApplication.Quit();
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var button = sender as RadioButton;
            var str = button.Name.ToString();

            switch (str)
            {
                case "serial":
                    SelectImageDelegate = this.SelectImage;
                    break;
                case "random":
                    SelectImageDelegate = this.SelectRandomImage;
                    break;
                default:
                    break;
            }
        }

        private static string GetRandomAudioFile(string location)
        {
            var rand = new Random();
            var files = Directory.GetFiles(location, "*.mp3");
            return files[rand.Next(files.Length)];
        }

        private static T GetRootObjFromStream<T>(Stream stream) where T : new()
        {
            StreamReader reader = new StreamReader(stream, Encoding.UTF8);
            String responseString = reader.ReadToEnd();

            var root = new T();
            var resVideo = root.FillWithJson<T>(responseString);
            return resVideo;
        }

        private static Stream GetStreamFromUrl(string urlTopLevelComm)
        {
            WebRequest requestComm = WebRequest.Create(urlTopLevelComm);
            WebResponse responseComm = requestComm.GetResponse();
            Stream streamComm = responseComm.GetResponseStream();
            return streamComm;
        }

        private void ReduceBackGroundMusicVolume()
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in this.pptApplication.ActivePresentation.Slides)
            {
                var slideShapes = slide.Shapes;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slideShapes)
                {
                    if (shape.Type == MsoShapeType.msoMedia &&
                        shape.MediaType == Microsoft.Office.Interop.PowerPoint.PpMediaType.ppMediaTypeSound)
                    {

                        // MessageBox.Show(shape.MediaFormat.Volume.ToString());
                        if (shape.Name == "musicBackground")
                            shape.MediaFormat.Volume = 0.2f;
                        else
                            shape.MediaFormat.Volume = 1.0f;
                    }
                }
            }

            this.pptPresentation.SaveAs(Constants.pptTestFile,
                    Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
                    MsoTriState.msoTrue);
        }

        public static void ConvertWavStreamToMp3File(ref MemoryStream ms, string savetofilename)
        {
            //rewind to beginning of stream
            ms.Seek(0, SeekOrigin.Begin);

            using (var retMs = new MemoryStream())
            using (var rdr = new WaveFileReader(ms))
            using (var wtr = new LameMP3FileWriter(savetofilename, rdr.WaveFormat, LAMEPreset.VBR_90))
            {
                rdr.CopyTo(wtr);
            }
        }

        private static void ReadAndSaveToMP3(string comment, int j)
        {
            using (SpeechSynthesizer reader = new SpeechSynthesizer())
            {
                reader.SetOutputToDefaultAudioDevice();
                //set some settings
                reader.Volume = 100;
                reader.Rate = 0; //medium

                //save to memory stream
                MemoryStream ms = new MemoryStream();
                reader.SetOutputToWaveStream(ms);

                //do speaking
                reader.Speak(comment);

                //now convert to mp3 using LameEncoder or shell out to audiograbber
                ConvertWavStreamToMp3File(ref ms, $"E:\\PowerPoint\\Read\\{j}.mp3");
            }
        }

        private async Task CreateVideoFromPPT()
        {
            //var app = new Microsoft.Office.Interop.PowerPoint.Application();

            //var pres = app.Presentations;
            //var file = pres.Open(@"E:\PresentationWIthMacro.pptm", MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

            //file.SaveCopyAs(@"E:\presentation1.mp4", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsMP4, MsoTriState.msoTrue);

            // ------------------
            if (!string.IsNullOrEmpty(outputFolderPath))
            {
                bool validName = false;
                string nFile = "test";
                Microsoft.Office.Interop.PowerPoint.Application objApp;
                Microsoft.Office.Interop.PowerPoint.Presentation objPres;
                objApp = new Microsoft.Office.Interop.PowerPoint.Application();
                objApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                objApp.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized;
                objPres = objApp.Presentations.Open(this.pptTest, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoTrue);

                try
                {
                    if (!nFile.Contains(".mp4"))
                    {
                        nFile += ".mp4";
                    }
                    objPres.SaveAs(System.IO.Path.Combine(outputFolderPath, nFile),
                        Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsMP4,
                        MsoTriState.msoTrue);
                    // Wait for creation of video file
                    while (objApp.ActivePresentation.CreateVideoStatus ==
                        Microsoft.Office.Interop.PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusInProgress
                        ||
                        objApp.ActivePresentation.CreateVideoStatus ==
                        Microsoft.Office.Interop.PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusQueued)
                    {
                        // Application.DoEvents();
                        // System.Threading.Thread.Sleep(500);
                        await Task.Delay(500);
                    }

                    objPres.Close();
                    objApp.Quit();
                    // Release COM Objects
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objPres);
                    objPres = null;
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objApp);
                    objApp = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }
            else
            {
                // MessageBox.Show("Please select a file PowerPoint file to convert");
            }
        }

        private void SelectTextBoxLocation(string pictureFileName)
        {
            var rndTextBoxLocation = rnd.Next(0, 3);
            string actualFileNameString = pictureFileName.Split('\\').LastOrDefault();
            var startString = actualFileNameString.Substring(0, 2);

            switch (startString)
            {
                case "11":
                case "21":
                case "31":
                    Constants.leftLocationTextBox1 = 20;
                    Constants.topLocationTextBox1 = 20;
                    break;
                case "12":
                case "22":
                case "32":
                    Constants.leftLocationTextBox1 = this.pptPresentation.PageSetup.SlideWidth / 2 - Constants.textBox1Width / 2;
                    Constants.topLocationTextBox1 = 20;
                    break;
                case "13":
                case "23":
                case "33":
                    Constants.leftLocationTextBox1 = this.pptPresentation.PageSetup.SlideWidth - (Constants.textBox1Width + 20);
                    Constants.topLocationTextBox1 = 20;
                    break;
                default:
                    Constants.leftLocationTextBox1 = textBoxLocationArray[rndTextBoxLocation][0];
                    Constants.topLocationTextBox1 = textBoxLocationArray[rndTextBoxLocation][1];
                    break;
            }
        }

        public Func<int, string> SelectImageDelegate;

        private string SelectRandomImage(int i)
        {
            // return Image.FromFile(images[rnd.Next(1, 9)]);
            int rndNext = 0;
            try
            {
                rndNext = this.rnd.Next(1, 8);

                return images[rndNext];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Random number selected for selecting an image: {rndNext}");
                return string.Empty;
            }
        }

        private string SelectImage(int selectImage)
        {
            return images[selectImage];
        }

        private void ReadTextAndImagesToMakePPT(string author, string comment, int likes, int slideNumber)
        {
            string pictureFileName = SelectImageDelegate(slideNumber - 1);

            if (this.isNewPresentation)
            {
                this.isNewPresentation = false;
                this.pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
                // Create the Presentation File
                this.pptPresentation = this.pptApplication.Presentations.Add(MsoTriState.msoTrue);

                textBoxLocationArray[0] = new List<float> { 20, 20 };
                textBoxLocationArray[1] = new List<float> { this.pptPresentation.PageSetup.SlideWidth / 2 - Constants.textBox1Width / 2, 20 };
                textBoxLocationArray[2] = new List<float> { this.pptPresentation.PageSetup.SlideWidth - (Constants.textBox1Width + 20), 20 };

            }

            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout =
            this.pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            // Create new Slide
            this.slides = this.pptPresentation.Slides;

            this.slide = this.slides.AddSlide(1, customLayout);
            // slide.TimeLine.InteractiveSequences.

            this.slide.Layout = PpSlideLayout.ppLayoutBlank;

            // Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];
            var picShape = this.slide.Shapes.AddPicture(pictureFileName,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0,
                this.pptPresentation.PageSetup.SlideWidth,
               this.pptPresentation.PageSetup.SlideHeight);

            // picShape.AnimationSettings.AnimateBackground = MsoTriState.msoCTrue;

            SelectTextBoxLocation(pictureFileName);

            for (int i = 1; i < 2; i++)
            {
                CreateTextBoxesAndSetProperties(author, comment, likes, i);
            }

            if (this.isLastSlide)
            {
                var sh = this.slide.Shapes.AddMediaObject2(soundFilePath, MsoTriState.msoFalse,
                    MsoTriState.msoTrue, 250, 10);


                sh.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue;
                sh.AnimationSettings.PlaySettings.HideWhileNotPlaying = MsoTriState.msoTrue;
                // sh.AnimationSettings.PlaySettings.PauseAnimation = MsoTriState.msoFalse;      

                sh.AnimationSettings.PlaySettings.StopAfterSlides = 999;
                sh.AnimationSettings.PlaySettings.LoopUntilStopped = MsoTriState.msoCTrue;

                sh.Name = "musicBackground";
                // sh.MediaFormat.Volume = 0.1f;
            }

            if (embeddedSpeechSupport)
            {
                var she = this.slide.Shapes.AddMediaObject2($"E:\\PowerPoint\\Read\\{slideNumber}.mp3", MsoTriState.msoFalse,
    MsoTriState.msoTrue, 350, 10);


                she.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue;
                she.AnimationSettings.PlaySettings.HideWhileNotPlaying = MsoTriState.msoTrue;
                // sh.AnimationSettings.PlaySettings.PauseAnimation = MsoTriState.msoFalse;      

                she.AnimationSettings.PlaySettings.StopAfterSlides = 1;
            }
            picShape.AnimationSettings.AnimateBackground = MsoTriState.msoCTrue;

            this.pptPresentation.SaveAs(Constants.pptTestFile,
                Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
                MsoTriState.msoTrue);
        }

        private void CreateTextBoxesAndSetProperties(string author, string comment, int likes, int textBoxNumber)
        {
            switch (textBoxNumber)
            {
                case 1:
                    var she = this.slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                    Constants.leftLocationTextBox1, Constants.topLocationTextBox1,
                    Constants.textBox1Width, Constants.textBox1Height);
                    break;
                default:
                    break;
            }

            // Add title
            this.objText = this.slide.Shapes[textBoxNumber + 1].TextFrame.TextRange;
            this.slide.Shapes[textBoxNumber + 1].Fill.BackColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
            this.slide.Shapes[textBoxNumber + 1].Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 0, 0).ToArgb();
            this.slide.Shapes[textBoxNumber + 1].Fill.Transparency = 0.65F;

            var rndEffect = this.rnd.Next(0, Constants.animationForText.Length);
            this.slide.Shapes[textBoxNumber + 1].AnimationSettings.EntryEffect = (PpEntryEffect)Constants.animationForText[rndEffect]; // PpEntryEffect.ppEffectAppear;
            this.slide.Shapes[textBoxNumber + 1].AnimationSettings.Animate = MsoTriState.msoTrue;

            this.slide.Shapes[textBoxNumber + 1].TextFrame.WordWrap = MsoTriState.msoTrue;
            this.objText.Font.Name = Constants.fontsArray[this.rnd.Next(0, 6)];

            this.objText.Font.Size = this.rnd.Next(18, 26);
            this.objText.Text = $"Author: {author}{Environment.NewLine}Comment:{comment}{Environment.NewLine}Likes:{likes}";
            this.objText.Font.Bold = MsoTriState.msoTrue;
        }

        private void OnSelectTextFileClicked(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();

            openFileDialog1.InitialDirectory = openFileDialogStartFolder;
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                textFilePath = openFileDialog1.FileName;
        }

        private void OnSelectAudioFileClicked(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();

            openFileDialog1.InitialDirectory = openFileDialogStartFolder;
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                soundFilePath = openFileDialog1.FileName;
        }

        private void OnSelectOutputFolderClicked(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = openFileDialogStartFolder;
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                outputFolderPath = dialog.FileName;
            }
        }

        private void OnSelectImagesFolderClicked(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = openFileDialogStartFolder;
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                imagesFolderPath = dialog.FileName;
            }
        }

        protected virtual bool IsFileinUse(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }

        private void OnPropertyChanged(string prop)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(prop));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public static class Extensions
    {
        public static T FillWithJson<T>(this T o, string data)
        {
            try
            {
                var ds = JsonConvert.DeserializeObject<T>(data);
                return (T)ds;
            }
            catch (Exception ex)
            {
                Console.WriteLine("JSON Deserialize Object Exception: {0}", ex.Message);
                return default(T);
            }
        }

        public static string GetJson<T>(this T o)
        {
            string s = "";
            try
            {
                s = JsonConvert.SerializeObject(o, Newtonsoft.Json.Formatting.Indented);
            }
            catch (Exception ex)
            {
                Console.WriteLine("JSON Serialize Object Exception: {0}", ex.Message);
                s = "Exception";
            }
            return s;
        }
    }
}
