using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
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
        String downloads_folder = "C:/downloads";
        String destination_folder = @"\\ccoder2\MediaCoder\Sharer IT";
        List<OpenPop.Mime.Message> messages_currently_processing = new List<OpenPop.Mime.Message>();
        BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        BackgroundWorker backgroundWorker2 = new BackgroundWorker();

        public Form1()
        {
            InitializeComponent();
            
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker1.WorkerReportsProgress = true;

        }
        void process_last_message()
        {
            var client = new Pop3Client();
            client.Connect("pop.gmail.com", 995, true);
            client.Authenticate("send2sharer@gmail.com", "manipulate");
            var count = client.GetMessageCount();

            if (count > 0)
            {
                OpenPop.Mime.Message message = client.GetMessage(count);

                if (!messages_currently_processing.Contains(message))
                {
                    // add message to currently_processing list
                    messages_currently_processing.Add(message);

                    Console.WriteLine(message.Headers.Subject);

                    FindPlainTextInMessage(message);

                    message_parser(message_text_file);

                    StreamReader reader = File.OpenText(links_text_file);
                    String link = "";
                    while ((link = reader.ReadLine()) != null)
                    {
                        if (link.Contains("youtube") || link.Contains("youtu.be"))
                        {
                            BackgroundWorker worker = new BackgroundWorker();
                            worker.DoWork += new DoWorkEventHandler(backgroundWorker2_DoWork);
                            worker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker2_ProgressChanged);
                            worker.WorkerReportsProgress = true;
                            worker.RunWorkerAsync(link);
                        }
                        Thread.Sleep(1000);
                    }

                    // delete message to not process again
                    try
                    {
                        client.DeleteMessage(count);
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                    client.Disconnect();

                    // remove from currently processing messages
                    messages_currently_processing.Remove(message);
                }
            }
            Thread.Sleep(1000);
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
        String download_youtube_video(String youtube_link)
        {
            
            String download_path = "";
            String file_name = "";
            String video_title = "";
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
            file_name = video.Title + video.VideoExtension;
            
            video_title = video.Title;
            video_title = video_title.Substring(0, video_title.Length / 2);
            
            video_title = video_title.Replace(":", "");
            video_title = video_title.Replace(@"\", "");
            video_title = video_title.Replace("/", "");
            video_title = video_title.Replace('"', ' ').Trim();

            download_path = Path.Combine(downloads_folder, video_title + video.VideoExtension);
            file_name = video_title + video.VideoExtension;
            var videoDownloader = new VideoDownloader(video, download_path);

            // Register the ProgressChanged event and print the current progress
            videoDownloader.DownloadProgressChanged += (sender, args) => Console.WriteLine(args.ProgressPercentage);
            
            /*
             * Execute the video downloader.
             * For GUI applications note, that this method runs synchronously.
             */
            try
            {
                videoDownloader.Execute();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());

                //try again
                try
                {
                    videoDownloader.Execute();
                }
                catch
                {

                }
            }
            // move file to sharer
            String source_path = downloads_folder + "/" + file_name;
            String destination_path = destination_folder + "\\" + file_name;
            try
            {
                File.Move(source_path, destination_path);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not move file. Error:\n" + ex.ToString());
            }

            return file_name;
        }        
        private int get_number_of_chars_in_string(String line, char char_to_count)
        {
            int count = 0;
            for(int i = 0; i < line.Count(); i++)
            {
                if(line[i] == char_to_count)
                {
                    count++;
                }
            }
            return count;
        }
        private void start_processing()
        {
            while (true)
            {
                process_last_message();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            start_processing();
        }
        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            // label1.Text = e.ProgressPercentage.ToString();
        }
        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            String link = e.Argument.ToString();
            download_youtube_video(link);
        }
        private void backgroundWorker2_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            // label1.Text = e.ToString();
        }
    }
}
