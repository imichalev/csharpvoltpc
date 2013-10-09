using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
//using System.Threading;
using Microsoft.Win32;

namespace Volt2Serial
{    
     //public void ReadData();
    public partial class Form1 : Form
    {     
        byte[] SpecialSymbol = new byte[1] {0x1b} ;
        
        /*
         [System.Runtime.InteropServices.DllImport("winmm.dll")] 
        public  timeGetTime_
            Lib "winmm.dll" () as long 
        */
        string ReadValue = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop", null );
        string DesktopPatch;
        string PortName;
        public float volt = 0;

        public Form1()
        {
            InitializeComponent();

            //chart1.Titles.Add("Volt");
           // this.chart1.ChartAreas.Add("Time");
           // this.chart1.ChartAreas.Add("Volt");
           // chart1.ChartAreas["Time"].AxisX.Minimum = 0;
           // chart1.ChartAreas["Time"].AxisX.Maximum = 100;
           // chart1.ChartAreas["Time"].AxisX.Interval = 1;
           // chart1.ChartAreas["Time"].AxisX.MajorGrid.LineColor = Color.White;
            //chart1.ChartAreas["Time"].AxisX.MajorGrid.LineDashStyle=Sys

            chart1.Series.Add("Volt");
          
            chart1.ChartAreas["Volt"].AxisY.Minimum = 0;
            chart1.ChartAreas["Volt"].AxisY.Maximum = 10;
            chart1.ChartAreas["Volt"].AxisY.Interval = 1;

            //Get Desktop Patch
            if (ReadValue == null)
            {
              DesktopPatch="c:\\Voltage.txt";
            }
            
            else 
            {
                DesktopPatch=ReadValue+"\\Votage.txt";
            }
            //Get Serial Port
            foreach ( string  Port   in  System.IO.Ports.SerialPort.GetPortNames())
            {
              comboBox1.Items.Add(Port);

            }
              //Set Comm Port
             comboBox1.SelectedIndex = 0;
             PortName = comboBox1.Text;
             serialPort1.ReadTimeout = 500;
             serialPort1.BaudRate = 9600;
             //Set Timer1 
             timer1.Interval = 100;
             
            
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          

        }

    private bool  ReadData ()
       {
                              
            if(PortName == "" || serialPort1.ReadTimeout==-1 ) 
            {
                MessageBox.Show ("Comm Port is not select");
                 return false;
            }

            if (serialPort1.IsOpen != true)
            {
                try
                {
                    serialPort1.Open();
                }
                catch (Exception )
                {
                   MessageBox.Show("Can't open port");
                   return false ;
                }

             }
                                          
          //Send SpecialSymbol to Serial and Wait for Answert
             serialPort1.Write(SpecialSymbol,0,1);
         // Try to Read Answert as Voltage 
             int ReadSerialChar;
            StringBuilder Voltage = new StringBuilder ();
                do
                  {
                    try
                      {
                      ReadSerialChar = serialPort1.ReadByte();
                       }
                 catch (TimeoutException )
                     {                              
                       MessageBox.Show (string.Format("The Device on port {0} not responding", PortName));
                        serialPort1.Close();
                         return false ;
                      }
                                               
                         Voltage.Append(Convert.ToChar (ReadSerialChar));
                         
                      } while (ReadSerialChar != 0x0d);
                      if (Voltage.Length != 0)
                      {
                          //float voltage = (Convert.ToInt32(Voltage.ToString())) * 5.0f / 1024.0f;
                          Voltage.Remove(Voltage.Length - 1, 1);
                          textBox1.Text = Voltage.ToString ();
                           volt=Convert.ToSingle (Voltage.ToString ());
                      }


                 serialPort1.Close();
                 return  true;
           }


        private void button3_Click(object sender, EventArgs e)
        {
            ReadData();
        } 

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();

            if (textBox1.BackColor  == Color.Green)
            {
                textBox1.BackColor = Color.Red;
            }
            else
            {
                textBox1.BackColor = Color.Green;
            }
           
            //ReadCommPort = true;
           
           
            if (ReadData()!=false )
            {
                timer1.Start(); 
            }
            

            chart1.Series["Volt"].Points.AddXY(1, volt);


        }

        private void button4_Click(object sender, EventArgs e)
        {
             timer1.Start();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            PortName = comboBox1.Text;
            textBox2.Text = PortName;
            if (PortName != "") serialPort1.PortName = PortName;           
        }

        

           
    }
}
