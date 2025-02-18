using System;
using System.Windows.Forms;

namespace LogApp_v1
{
    public partial class Form2 : Form
    {

        int speed = 500;

        public Form2()
        {
            InitializeComponent();
            label1.Parent = this;
            progressBar1.Parent = this;
            label2.Parent = this;


            fadeForm(0.7);


            timer1.Interval = speed;
            timer1.Enabled = true;

        }
        
        public void Form2_Load(object sender, EventArgs e)
        {
            //label1.Parent = this;
            //progressBar1.Parent = this;
            //label2.Parent = this;


           //fadeForm(0.7);


            //timer1.Interval = speed;
            //timer1.Enabled = true;


        }
        public void timer1_Tick(object sender, EventArgs e)
        {
            if (label1.Text ==   "PROCESSING")
            {
                label1.Text = "PROCESSING .";
            }
            else if (label1.Text == "PROCESSING .")
            {
                label1.Text = "PROCESSING . .";
            }
            else if (label1.Text == "PROCESSING . .")
            {
                label1.Text = "PROCESSING . . .";
            }
            else if (label1.Text == "PROCESSING . . .")
            {
                label1.Text = "PROCESSING";
            }
        }
        public void fadeForm(double totalSec)
        {
            if(totalSec == 0)
            {
                Opacity = 1;
                Refresh();
            }
            double then = DateTime.Now.TimeOfDay.TotalSeconds;
            double difference = 0;
            //difference is the percentage of the total seconds elapsed
            while (difference < 1)
            {
                Opacity = difference;

                difference = (DateTime.Now.TimeOfDay.TotalSeconds - then) / totalSec;
                System.Threading.Thread.Sleep(10);
                //Refresh();
            }
            Opacity = 1;
            Refresh();
        }

        public void Form2_Shown(object sender, EventArgs e)
        {
            timer1.Interval = speed;
            timer1.Enabled = true;
        }
    }
}
