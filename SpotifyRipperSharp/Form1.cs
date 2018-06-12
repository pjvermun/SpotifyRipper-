using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using CSCore;
using CSCore;
using CSCore.Codecs.WAV;
using CSCore.MediaFoundation;
using CSCore.SoundIn;
using CSCore.Streams;
using CSCore.CoreAudioAPI;
using System.Text.RegularExpressions;
using System.IO;

namespace SpotifyRipperSharp
{
    public partial class Form1 : Form
    {
        private WasapiCapture capture;
        private WaveWriter writer;
        private int songs;

        List<String> _items = new List<string>();
        public Form1()
        {
            InitializeComponent();
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            listBox1.DataSource = _items;
            button5.Enabled = false;
            songs = 0;
        }

        private void startRecording(string name)
        {
            capture = new WasapiLoopbackCapture();
            capture.Initialize();
            //var writer = MediaFoundationEncoder.CreateMP3Encoder(capture.WaveFormat, name + ".mp3", 320000);
            writer = new WaveWriter("./recorded/"+ name + ".wav", capture.WaveFormat);
            capture.DataAvailable += (s, capData) =>
            {
                writer.Write(capData.Data, capData.Offset, capData.ByteCount);
            };
            capture.Start();
        }

        private void stopRecording()
        {
            if (writer != null && capture != null)
            {
                capture.Stop();
                writer.Dispose();
                capture.Dispose();
            }
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            PowerShell test = PowerShell.Create().AddScript("Get-Process |where {$_.mainWindowTItle -and $_.name -eq 'Spotify'} | format-table mainwindowtitle -HideTableHeaders");
            test.AddCommand("Out-String");
            String previous = null;
            String str = null;
            Regex illegalInFileName = new Regex(@"[\\/:*?""<>|]");
            while (true)
            {
                
                if (backgroundWorker1.CancellationPending)
                {
                    stopRecording();
                    return;
                }
                foreach (string stripped in test.Invoke<string>())
                {
                    str = stripped.Replace("\r\n", "");
                    str = illegalInFileName.Replace(str, "");
                    if (previous != str && str != "Spotify        ")
                    {
                        _items.Add(str);
                        stopRecording();
                        startRecording(str);
                        backgroundWorker1.ReportProgress(12);
                        songs++;
                    }
                    
                    previous = str;
                }
                Thread.Sleep(100);
                
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            textBox3.Text = songs.ToString();
            listBox1.DataSource = null;
            listBox1.DataSource = _items;
            textBox1.Text = _items.Last();
            string[] a = Directory.GetFiles("./recorded/", "*.*");
            long b = 0;
            foreach (string name in a)
            {
                FileInfo info = new FileInfo(name);
                b += info.Length;
            }
            b /= 1000000;
            textBox2.Text = b.ToString() + " MB";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
            button3.Enabled = false;
            button5.Enabled = true;
            songs = 0;
            textBox3.Text = "0";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
            button3.Enabled = true;
            button5.Enabled = false;
        }
    }
}
