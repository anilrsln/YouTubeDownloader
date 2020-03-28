using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using VideoLibrary;
using System.Data.OleDb;
using System.Data;
using System.Runtime.InteropServices;
using IronXL;
using System.Linq;
using MediaToolkit.Model;
using MediaToolkit;

namespace YouTubeDownloader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(textBox1.Text) || !textBox1.Text.Contains("www.youtube.com") )
            {
                MessageBox.Show("Hatalı link girildi, tekrar deneyiniz.", "HATA");
                return;
            }

            MessageBox.Show("İndirme işleminiz başladı.", "LÜTFEN BEKLEYİNİZ");

            var youTube = YouTube.Default; // starting point for YouTube actions
            var video = youTube.GetVideo(textBox1.Text); // gets a Video object with info about the video


            if (!Directory.Exists(@"C:\YouTube\")) Directory.CreateDirectory(@"C:\YouTube\"); 
            File.WriteAllBytes(@"C:\YouTube\" + video.FullName, video.GetBytes());

            var inputFile = new MediaFile { Filename = @"C:\YouTube\" + video.FullName };
            var outputFile = new MediaFile { Filename = $"{@"C:\YouTube\" + video.FullName}.mp3" };

            using (var engine = new Engine())
            {
                engine.GetMetadata(inputFile);

                engine.Convert(inputFile, outputFile);
            }

            File.Delete(@"C:\YouTube\" + video.FullName);
            MessageBox.Show("Ses dosyasını başarıyla indirdiniz.", "TEBRİKLER");
        }
        
        private void button2_Click(object sender, EventArgs e)
        {

        }


        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
