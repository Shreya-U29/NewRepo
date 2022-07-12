using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.IO;
using System.Management;

namespace Speech_recognition
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer Sara = new SpeechSynthesizer();   
        SpeechRecognitionEngine startlistening = new SpeechRecognitionEngine();  
        Random random = new Random();
        int recTimeOut = 0;    //counter between 1 and 2 
        DateTime timenow = DateTime.Now;



        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            recognizer.SetInputToDefaultAudioDevice();        // default audio to device such as microphone or speaker
            recognizer.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"DefaultCommands.txt")))));
            recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Default_SpeechRecognized);
            recognizer.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(recognizer_SpeechRecognized);
            recognizer.RecognizeAsync(RecognizeMode.Multiple);

            startlistening.SetInputToDefaultAudioDevice();
            startlistening.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"DefaultCommands.txt")))));
            startlistening.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Startlistening_SpeechRecognized);

        }

        private void recognizer_SpeechRecognized(object sender, SpeechDetectedEventArgs e)
        {
            recTimeOut = 0;
        }

        private void Startlistening_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string speech = e.Result.Text;

            if(speech == "Wake up")
            {
                startlistening.RecognizeAsyncCancel();
                Sara.SpeakAsync("Yes, I am here");
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
            }
        }

        private void Default_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            int ranNum;
            string speech = e.Result.Text;

            if(speech == "Hello")
            {
                Sara.SpeakAsync("Hello Shreya");
            }
            if (speech == "How are you")
            {
                Sara.SpeakAsync("I am working good");
            }
            if (speech == "What time is it")
            {
                Sara.SpeakAsync(DateTime.Now.ToString("h mm tt"));   // for example 1.45 PM
            }
            if(speech == "Stop listening")
            {
                Sara.SpeakAsyncCancelAll();
                ranNum = random.Next(1,2);
                if(ranNum == 1)
                {
                    Sara.SpeakAsync("Yes sir");
                }
                          
                if(ranNum == 2)
                {
                    Sara.SpeakAsync("I am sorry I will be quiet");
                }                           
            }
            if(speech == "Stop talking")
            {
                Sara.SpeakAsync("If you need me just ask");
                recognizer.RecognizeAsyncCancel();
                startlistening.RecognizeAsync(RecognizeMode.Multiple);
            }
            if(speech == "Show commands")
            {
                string[] commands = (File.ReadAllLines(@"DefaultCommands.txt"));
                listBox1.Items.Clear();
                listBox1.SelectionMode = SelectionMode.None;
                listBox1.Visible = true;
                foreach(string command in commands)
                {
                    listBox1.Items.Add(command);
                }
            }
            if(speech == "Hide Commands")
            {
                listBox1.Visible = false;
            }
            if(speech == "Close it")
            {
                Application.Exit();

            }
            if(speech == "What is the battery status")
            {
                BatteryStatus();
            }
            if(speech == "Open google")
            {
                System.Diagnostics.Process.Start("https://www.google.com/");
            }
            if(speech == "Play DNA")
            {
                System.Diagnostics.Process.Start("https://www.youtube.com/watch?v=MBdVXkSdhwU");
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(recTimeOut == 10)
            {
                recognizer.RecognizeAsyncCancel();
            }
            else if(recTimeOut == 11)
            {
                timer1.Stop();
                startlistening.RecognizeAsync(RecognizeMode.Multiple);
                recTimeOut = 0;
            }
        }
        private void BatteryStatus()
        {
            System.Management.ManagementClass wmi = new System.Management.ManagementClass("Win32_Battery");
            var allBatteries = wmi.GetInstances();

            foreach(var battery in allBatteries)
            {
                int estimatedChargeRemaining = Convert.ToInt32(battery["EstimatedChargeremaining"]);
                Sara.SpeakAsync("The estimated charge is" + estimatedChargeRemaining + "%");
                
            }
        }
        
    }
}
