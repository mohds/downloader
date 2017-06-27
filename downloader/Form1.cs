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
using System.Net;

namespace downloader
{
    public partial class Form1 : Form
    {
        String message_text_file = "message_text.txt";
        String links_text_file = "links.txt";
        String downloads_folder = "C:/downloads";
        String destination_folder = @"\\ccoder2\MediaCoder\Sharer IT";
        List<OpenPop.Mime.Message> messages_currently_processing = new List<OpenPop.Mime.Message>();
        List<String> downloaded_videos = new List<string>();
        List<String> failed_links = new List<string>();
        String sender = "";
        BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        BackgroundWorker backgroundWorker2 = new BackgroundWorker(); // leave it be
        BackgroundWorker backgroundWorker3 = new BackgroundWorker();
        List<String> workers_busy = new List<string>(); // a list in which each worker states that they are busy
        List<Label> labels = new List<Label>(); // to keep track of label positions
        List<Label> label_statuses = new List<Label>(); // status manipulation
        List<String> labels_to_add = new List<string>();
        int status_updated = 0; // changes to the index which needs an update

        List<String> allowed_senders = new List<string>();

        public Form1()
        {
            InitializeComponent();

            populate_allowed_senders();

            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker1.WorkerReportsProgress = true;

            backgroundWorker3.DoWork += backgroundWorker3_DoWork;
            backgroundWorker3.ProgressChanged += backgroundWorker3_ProgressChanged;
            backgroundWorker3.WorkerReportsProgress = true;
            
            labels.Add(label1);
        }
        private void populate_allowed_senders()
        {
            allowed_senders.Add("m.salloum@almayadeen.net");
            allowed_senders.Add("z.hteit@almayadeen.net");
            allowed_senders.Add("s.souwaidan@almayadeen.net");
            allowed_senders.Add("w.elahmar@almayadeen.net");
            allowed_senders.Add("M.Awad@almayadeen.net");
            allowed_senders.Add("ma.saad@almayadeen.net");
            allowed_senders.Add("j.krayem@almayadeen.net");
            allowed_senders.Add("h.chreim@almayadeen.net");
            allowed_senders.Add("a.hammoud@almayadeen.net");
            allowed_senders.Add("a.kobeissi@almayadeen.net");
            allowed_senders.Add("r.dagher@almayadeen.net");
        }
        void process_last_message()
        {
            var client = new Pop3Client();
            try
            {
                client.Connect("pop.gmail.com", 995, true);
                client.Authenticate("send2sharer@gmail.com", "manipulate");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.WriteLine("Failed to authenticate with server");
                return;
            }
            var count = client.GetMessageCount();

            if (count > 0)
            {
                downloaded_videos.Clear();
                failed_links.Clear();
                workers_busy.Clear();

                OpenPop.Mime.Message message = client.GetMessage(count);

                if (!messages_currently_processing.Contains(message))
                {
                    // grab sender
                    sender = message.Headers.From.Address;

                    // security
                    if (!(allowed_senders.Contains(sender)))
                    {
                        // delete message to not process again
                        try
                        {
                            client.DeleteMessage(count);
                            client.Disconnect();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                        Console.WriteLine("No access for " + sender);
                        return;
                    }

                    // add message to currently_processing list
                    messages_currently_processing.Add(message);

                    Console.WriteLine(message.Headers.Subject + " From: " + sender);

                    FindPlainTextInMessage(message);

                    if (message_parser(message_text_file)) // returns true on success
                    { 

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
                                workers_busy.Add("busy");
                            }
                            Thread.Sleep(1000);
                        }
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
                while(workers_busy.Count > 0)
                {
                    // halt while workers work
                }
                // all workers done
                send_email(sender);
            }
            //client.Disconnect();
            //client.Dispose();                       
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
        bool message_parser(String file_name)
        {
            bool return_me = false;
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
                        return_me = true;
                    }
                }
            }
            return return_me;
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
            IEnumerable<VideoInfo> videoInfos;
            try
            {
                videoInfos = DownloadUrlResolver.GetDownloadUrls(youtube_link);

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.WriteLine("GetDownloadUrls() failed, trying again...");
                videoInfos = DownloadUrlResolver.GetDownloadUrls(youtube_link);
                
            }
            
            /*
             * Select the first .mp4 video with 360p resolution
             */
            VideoInfo video = videoInfos.First(info => info.VideoType == VideoType.Mp4 && info.Resolution == 360);

            double video_size = extract_video_size(video.DownloadUrl);
            // convert to MB
            video_size = video_size / 1000000;

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
            var charsToRemove = new string[] { "@", ",", ".", ";", "'", @"\", "/", "?", "|", "}", "{", "[", "]", "<", ">", "$", "%", "^", "*", "!", "#", "=", "+", " ", "é", "ë", "ç", ":", "\"" };
            foreach (var c in charsToRemove)
            {
                video_title = video_title.Replace(c, " ");
            }

            download_path = Path.Combine(downloads_folder, video_title + video.VideoExtension);
            file_name = video_title + video.VideoExtension;
            var videoDownloader = new VideoDownloader(video, download_path);

            // Register the ProgressChanged event and print the current progress
            // videoDownloader.DownloadProgressChanged += (sender, args) => Console.WriteLine(args.ProgressPercentage);

            /*
             * Execute the video downloader.
             * For GUI applications note, that this method runs synchronously.
             */

            // index of status for later use
            int index_status = 0;

            try
            {
                Console.WriteLine("Started downloading: " + video_title);
                add_row(video_title, "Started");
                index_status = label_statuses.Count - 1;

                videoDownloader.DownloadProgressChanged += (sender, args) => {
                    //label_statuses[index_status].SafeInvoke(d => d.Text = args.ProgressPercentage.ToString());
                    String percentage_downloaded_str = args.ProgressPercentage.ToString();
                    double percentage_downloaded = Convert.ToDouble(percentage_downloaded_str);
                    
                    percentage_downloaded = percentage_downloaded / 100;
                    double size_downloaded = percentage_downloaded * video_size;
                    video_size = Math.Round(video_size, 2);
                    size_downloaded = Math.Round(size_downloaded, 2);
                    String display_me = size_downloaded.ToString() + "MB / " + video_size.ToString() + "MB";
                    label_statuses[index_status].SafeInvoke(d => d.Text = display_me);
                };

                videoDownloader.Execute();                
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());

                // try again
                try
                {
                    Console.WriteLine("Started downloading: " + video_title);
                    add_row(video_title, "Started");
                    index_status = label_statuses.Count - 1;
                    videoDownloader.Execute();
                }
                catch
                {
                    // exit function
                    return file_name;
                }
            }

            // to check title-status harmony
            //label_statuses[index_status].SafeInvoke(d => d.Text = video_title);
            //label_statuses[index_status].Text = index_status.ToString();
            //this.SafeInvoke(d => d.Refresh());

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

            // add name to downloaded list
            downloaded_videos.Add(video_title);

            // update interface
            // label_statuses[index_status].Text = "Done";
            label_statuses[index_status].SafeInvoke(d => d.Text = "Done");
            this.SafeInvoke(d => d.Refresh());
            // status_updated = index_status; // it is only required to change it from 0 (so far)

            return file_name;
        }
        private int extract_video_size(String video_url)
        {
            int video_size = 1;
            String[] sections;
            sections = video_url.Split('&');
            foreach(var section in sections)
            {
                if (section.Contains("clen="))
                {
                    String video_size_str = section.Substring(section.IndexOf('=') + 1, (section.Length - 1) - section.IndexOf('='));
                    video_size = Int32.Parse(video_size_str);
                    break;
                }
            }
            return video_size;
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
            button1.Enabled = false;
            backgroundWorker3.RunWorkerAsync();
            
        }
        private void add_row(String title, String status)
        {
            Label label_title = new Label();
            Label label_status = new Label();
            label_title.Width = 100;
            label_status.Width = 300;
            int workers_count = workers_busy.Count + 1;
            label_title.Text = title;
            label_status.Text = status;
            int y_of_last_label = labels[labels.Count - 1].Location.Y;
            label_title.Left = label1.Location.X;
            label_title.Top = y_of_last_label + 30;
            label_status.Left = label2.Location.X;
            label_status.Top = y_of_last_label + 30;
            //this.Controls.Add(label_status);
            //this.Controls.Add(label_title);
            this.SafeInvoke(d => d.Controls.Add(label_status));
            this.SafeInvoke(d => d.Controls.Add(label_title));
            this.SafeInvoke(d => d.Refresh());
            labels.Add(label_title);

            // keeping labels within reach
            label_statuses.Add(label_status);
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
            try
            {
                download_youtube_video(link);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not download: " + link);
                failed_links.Add(link);
            }
            if (workers_busy.Count > 0)
            {
                workers_busy.Remove(workers_busy[0]);
            }
        }
        private void backgroundWorker2_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {

        }
        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            /*
            while (true)
            {
                // interface handling
                if (labels_to_add.Count > 0)
                {
                    int index = labels_to_add.Count - 1;
                    add_row(labels_to_add[index], label_statuses[index].Text);
                    labels_to_add.Remove(labels_to_add[index]);
                }
                if (status_updated != 0)
                {
                    this.SafeInvoke(d => d.Refresh());
                    status_updated = 0;
                }
                Thread.Sleep(100);
            }*/
        }
        private void backgroundWorker3_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {


        }
        private void send_email(String to)
        {
            var client = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new NetworkCredential("send2sharer@gmail.com", "manipulate"),
                EnableSsl = true
            };
            String message_body = "";
            message_body += "On sharer:\n";
            foreach(var title in downloaded_videos)
            {
                message_body += title + "\n";
            }
            if(failed_links.Count > 0)
            {
                message_body += "\nFailed links:\n";
                foreach(var link in failed_links)
                {
                    message_body += link + "\n";
                }
            }
            message_body += "\nRegards,\n";
            message_body += "Download Assistant\n";
            client.Send("send2sharer@gmail.com", to, "On sharer", message_body);
            Console.WriteLine("Email sent");
        }


    }
}
