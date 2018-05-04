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
using System.Speech;
using System.Speech.Recognition;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

namespace voicerecognizer
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        SpeechSynthesizer synt = new SpeechSynthesizer();

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsync(RecognizeMode.Multiple);
            button1.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Choices command = new Choices();
            command.Add(new string[] { "say hello", "print my name", "open calender", "speak selector text", "open notepad", "close notepad", 
                "open google chrome","open camera","what date is today","navigate to google","facebook","youtube",
           "on which os you are running"," file explorer","close camera","close file explorer","close youtube","close facebook" });
            GrammarBuilder gbuilder = new GrammarBuilder();
            gbuilder.Append(command);
            Grammar grammer = new Grammar(gbuilder);
            recEngine.LoadGrammarAsync(grammer);
            recEngine.SetInputToDefaultAudioDevice();
            recEngine.SpeechRecognized += recEngine_SpeechRecognized;
        }

      

       void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text) {
                case "say hello":
                    synt.SpeakAsync("hello how are you ");
                    break;
                case "print my name":
                    richTextBox1.Text += "\nshafay";
                    break;

                case "Open calender":
                    synt.SpeakAsync("opening calender");
                    Process.Start("calender.exe");
                    break;

                case "open notepad":
                   
                    Process.Start("notepad.exe");
                    break;

                case "speak selector text":
                    synt.SpeakAsync(richTextBox1.Text);
                    break;

                case "open google chrome":
                    
                    Process.Start("chrome.exe");
                    break;

                case "open camera":
                    Process process = new Process();
                    process.StartInfo.FileName = "microsoft.windows.camera:";
                    process.Start();
                    break;

           
                case "what date is today":
                    synt.SpeakAsync(DateTime.Now.ToString("MM-dd-yyyy"));
                    break;

                case "facebook":
                    Process.Start("chrome.exe","https://www.facebook.com");
              
                    break;

                case "navigate to google":
                    Process.Start("chrome.exe","https://www.google.co.in");
                    break;
                
                case "youtube":
                    Process.Start("chrome.exe","https://www.youtube.com");
                    break;

                case "on which os you are running":
                    synt.SpeakAsync("microsoft's windows operating system");
                    break;

                case "file explorer":
                    Process.Start("explorer.exe");
                    break;
            }

            Process[] pname2 = Process.GetProcessesByName("notepad");
            if (pname2.Length != 0)
                switch (e.Result.Text)
                {
                    case "close notepad":
                        pname2[0].Kill();
                        break;
                    
                }

            Process[] pname1 = Process.GetProcessesByName("windowscamera");
            if (pname1.Length != 0)
                switch (e.Result.Text)
                {
                    case "close camera":
                        pname1[0].Kill();
                        break;
                   
                }

                Process[] pname3 = Process.GetProcessesByName("explorer");
                if (pname3.Length != 0)
                    switch (e.Result.Text)
                    {
                        case "close file explorer":
                            SendKeys.Send("%{f4}");
                            break;
                      
                    }

                Process[] pname4 = Process.GetProcessesByName("tube");
                if (pname3.Length != 0)
                    switch (e.Result.Text)
                    {
                        case "close youtube":
                            SendKeys.Send("%{f4}");
                            break;

                    }

                Process[] pname4 = Process.GetProcessesByName("book");
                if (pname3.Length != 0)
                    switch (e.Result.Text)
                    {
                        case "close facebook":
                            SendKeys.Send("%{f4}");
                            break;

                    }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsyncStop();
            button1.Enabled = false;
        }

       
    }
}
