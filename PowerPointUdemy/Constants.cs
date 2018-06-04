using Microsoft.Office.Interop.PowerPoint;
using PowerPointUdemy.YoutubeModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointUdemy
{
    public static class Constants
    {
        public static string key = "PUT YOUR KEY HERE";
        public static string urlTopLevelCommentPrefix = "https://www.googleapis.com/youtube/v3/commentThreads?videoId=";
        public static string urlTopLevelCommentsPostfix = $"&part=snippet&order=relevance&maxResults=100&key={key}";
        public static string urlTopLevelCommentsPostfixWithToken = string.Empty;

        //public static string pathCommentsTextFile = @"E:\PowerPoint\Comments.txt";


        #region PPT

        //public static string videoGenerationPath = @"E:\PowerPoint\";
        public static string pptTestFile = @"E:\PowerPoint\test.pptx";
        //public static string pptTestMusic = @"E:\PowerPoint\test.mp4";
        //public static string commentsFilePath = @"E:\PowerPoint\Comments.txt";
         
        public static float leftLocationTextBox1 = 0.0F;
        public static float topLocationTextBox1 = 0.0F;
        public static int textBox1Height = 0;
        public static int textBox1Width = 0;

        public static string soundFileLocation = string.Empty;
        public static string[] fontsArray = new string[]
       {
        "Helvetica",
        "Garamond",
        "Harlow Solid Italic",
        "AR BERKLEY",
        "Rockwell",
        "Keania One"
        };

        //public static string imagesFolderPath = @"E:\PowerPoint\Images";

        #endregion

        public static int[] animationForText = new int[] { 3844, 1793,
            // 3849,
            // 3331, 3334, 3330, //3332, 3845,
        2820, 2817,
        2817,
        2819,
        2818,
        //3857,
        //3858,
        //3859,
        //3860,
        //3861,
        2305,
        2306,
        3350,
        3349,
        3345,
        3346,
        3347,
        3348,
        3356,
         3335,
         3336,
         3329,
         3331,
         3330,
         3333,
         3334,
         2820,
        };
    }
}
