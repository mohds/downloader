using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using OpenPop.Pop3;
using OpenPop.Mime;
using System.IO;
using System.Text.RegularExpressions;
using YoutubeExtractor;

namespace downloader
{
    public partial class Form1 : Form
    {
        String message_text_file = "message_text.txt";
        String links_text_file = "links.txt";

        public Form1()
        {
            InitializeComponent();
            get_last_message();
        }
        void get_last_message()
        {
            var client = new Pop3Client();
            client.Connect("pop.gmail.com", 995, true);
            client.Authenticate("send2sharer@gmail.com", "manipulate");
            var count = client.GetMessageCount();
            OpenPop.Mime.Message message = client.GetMessage(count);
            Console.WriteLine(message.Headers.Subject);

            FindPlainTextInMessage(message);

            message_parser(message_text_file);

            StreamReader reader = File.OpenText(links_text_file);
            String link = "";
            while ((link = reader.ReadLine()) != null)
            {
                if (link.Contains("youtube") || link.Contains("youtu.be"))
                {
                    download_youtube_video(link);
                }
            }

        }
        void FindPlainTextInMessage(OpenPop.Mime.Message message)
        {
            MessagePart plainText = message.FindFirstPlainTextVersion();
            if (plainText != null)
            {
                // Save the plain text to a file, database or anything you like
                plainText.Save(new FileInfo(message_text_file));
            }
        }
        void message_parser(String file_name)
        {
            try
            {
                File.Delete(links_text_file);
            }
            catch
            {
                Console.WriteLine("Creating New links file.");
            }

            StreamReader reader = File.OpenText(file_name);
            String line = "";
            while ((line = reader.ReadLine()) != null)
            {
                var linkParser = new Regex(@"\b(?:https?://|www\.)\S+\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                var rawString = line;
                foreach (Match m in linkParser.Matches(rawString))
                {
                    string link = m.Value;
                    using (StreamWriter writer = File.AppendText(links_text_file))
                    {
                        writer.WriteLine(link);
                    }
                }
            }
        }
        void download_youtube_video(String youtube_link)
        {
            /*
             * Get the available video formats.
             * We'll work with them in the video and audio download examples.
             */
            IEnumerable<VideoInfo> videoInfos = DownloadUrlResolver.GetDownloadUrls(youtube_link);
            /*
             * Select the first .mp4 video with 360p resolution
             */
            VideoInfo video = videoInfos.First(info => info.VideoType == VideoType.Mp4 && info.Resolution == 360);

            /*
             * If the video has a decrypted signature, decipher it
             */
            if (video.RequiresDecryption)
            {
                DownloadUrlResolver.DecryptDownloadUrl(video);
            }

            /*
             * Create the video downloader.
             * The first argument is the video to download.
             * The second argument is the path to save the video file.
             */
            var videoDownloader = new VideoDownloader(video, Path.Combine("M:/Downloads", video.Title + video.VideoExtension));

            // Register the ProgressChanged event and print the current progress
            //videoDownloader.DownloadProgressChanged += (sender, args) => Console.WriteLine(args.ProgressPercentage);

            /*
             * Execute the video downloader.
             * For GUI applications note, that this method runs synchronously.
             */
            videoDownloader.Execute();
        }
    }
}
