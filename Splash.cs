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

namespace CommentToAttendence
    {
    public partial class Splash : Form
        {
        //string port = "COM4";
        //SerialCommunication serial = new SerialCommunication();
        //LoadingForm loading = new LoadingForm();
        //Login login = new Login();
        public Splash()
            {
            InitializeComponent();
            //EnterLoginForm();
            }
        private void Splash_Load(object sender, EventArgs e)
            {
            ProgressTimer.Start();
            }


        int startpoint = 0;

        private void ProgressTimer_Tick(object sender, EventArgs e)
            {
            startpoint += 1;
            progressBar1.Value = startpoint;
            if (progressBar1.Value == 100)
                {
                progressBar1.Value = 0;
                ProgressTimer.Stop();
                Form1 login = new Form1();
                this.Hide();
                login.Show();
                }
            }


        /*
        int i;
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
            {
            //for (i = 0; i <= 30; i++)
              //  {
                serial.actionMatch();
             //   backgroundWorker1.ReportProgress(i);
              //  }
            
            }

        /*
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
            {
            progressBar1.Value = e.ProgressPercentage;
            }
        */
        /*
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            //MessageBox.Show("U");
            login.Show();
            this.Dispose();
            
            }

        private void Splash_FormClosed(object sender, FormClosedEventArgs e)
            {
            login.Show();
            }

        private void Splash_FormClosing(object sender, FormClosingEventArgs e)
            {
            login.Show();
            }
        */


        /*
        private void button1_Click(object sender, EventArgs e)
            {
            backgroundWorker1.RunWorkerAsync();
            }

        private void button2_Click(object sender, EventArgs e)
            {
            backgroundWorker1.CancelAsync();
            }
        */

        /*
        int i;
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
            {
            // for(int i =0;i <=100;i++)
            //  {
            //  if (backgroundWorker1.CancellationPending)
            //   {
            //   e.Cancel = true;
            //     }
            //     else
            //    {
            //HeavyTask();
            serial.actionMatch();
            backgroundWorker1.ReportProgress(i);
            //   }
            //   }
            }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
            {
            progressBar1.Value = e.ProgressPercentage;
            }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            if (e.Cancelled)
                {
                //display("Work Cancelled");
                progressBar1.Value = 0;
                }
            else
                {
                //display("Work done");
                }
            }
        */

        /*
        void HeavyTask()
            {
            Thread.Sleep(100);
            }
        */




        /*

        void EnterLoginForm()
            {
            //serial.DetectArduino();
            //serial.actionMatch();
            //backgroundwork(serial.actionMatch);
            //this.Close();
            //Login login = new Login();
            //this.Close();
            login.Show();
            }
        void backgroundwork(Action action)
            {
            Thread backgroundThread = new Thread(
                   new ThreadStart(() =>
                   {
                       action.Invoke();
                       //this.Close();
                       //login.Show();
                       // EnterLoginForm();
                       if (this.InvokeRequired)
                          {
                         this.BeginInvoke(new Action(() =>  this.Close()));
                          }
                        else
                           {
                          this.Close();
                          }
                   }
                   ));
            backgroundThread.Start();
            }

        private void Splash_FormClosing(object sender, FormClosingEventArgs e)
            {
            Login login = new Login();
            EnterLoginForm();
            }

        private void Splash_FormClosed(object sender, FormClosedEventArgs e)
            {
            Login login = new Login();
            EnterLoginForm();
            }

      */

        }
    }
