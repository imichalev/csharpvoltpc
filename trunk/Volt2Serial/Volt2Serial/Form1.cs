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
        public double volt = 0;
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

            //chart1.Series.Add("Volt");
            
            //ChartSeries Volt = new ChartSeries("volt", ChartSeriesType.Line);
             //chart1.ChartAreas.Add("Volt1");
             //chart1.ChartAreas["Volt1"].AxisY.Minimum = 50;
             //chart1.ChartAreas["Volt1"].AxisY.Maximum = 100;
             //chart1.ChartAreas["Volt1"].AxisY.Interval = 1;
              //if (MessageBox.Show(" I am here!", "Eto me", MessageBoxButtons.YesNo) == DialogResult.Ignore )
                 
           
            
            //var message = new MessageDialog("dfsd");
            //var dialog = (M3Form)Activator.CreateInstance(dialogType, new PropertyGrid());
            chart1.Series.Add("Volt");
            chart1.Series["Volt"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            //chart1.Series["Volt"].XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.DateTime;
            chart1.Series["Volt"].YValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;

            chart1.ChartAreas.Add("Volt1");
            chart1.ChartAreas["Volt1"].AxisX.MajorGrid.LineColor = Color.LightGray;
            chart1.ChartAreas["Volt1"].AxisY.MajorGrid.LineColor = Color.LightGray;
            chart1.ChartAreas["Volt1"].AxisX.LabelStyle.Font = new System.Drawing.Font("Calibri", 8);
            chart1.ChartAreas["Volt1"].AxisY.LabelStyle.Font = new System.Drawing.Font("Calibri", 8);
           
           // chart1.ChartAreas["Volt1"].AxisX.LabelStyle.Format = "dd MMM\nHH:mm";
            chart1.ChartAreas["Volt1"].AxisX.Title = "Time";
            chart1.ChartAreas["Volt1"].AxisX.Minimum = 0;
            chart1.ChartAreas["Volt1"].AxisX.Maximum = 60;

            chart1.ChartAreas["Volt1"].AxisY.Minimum = 0;
            chart1.ChartAreas["Volt1"].AxisY.Maximum = 0.1;
           // chart1.ChartAreas["Volt1"].AxisY2.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.True;
            //chart1.ChartAreas["Volt1"].AxisY2.MajorGrid.Enabled = false;
            //chart1.ChartAreas["Volt1"].AxisY.MajorGrid.Enabled = false;

            chart1.ChartAreas["Volt1"].AxisY.Title = "Volt";
            //chart1.ChartAreas["Volt1"].AxisY.LabelStyle.Format = Int32;
            //chart1.ChartAreas["Volt1"].AlignmentStyle = System.Windows.Forms.DataVisualization.Charting.AreaAlignmentStyles.PlotPosition;
            


            

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
            if (comboBox1.SelectedIndex > 0)
            {
                comboBox1.SelectedIndex = 0;
                PortName = comboBox1.Text;
            }
             serialPort1.ReadTimeout = 500;
             serialPort1.BaudRate = 9600;
             //Set Timer1 
             timer1.Interval = 1000;
             
            
           
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

            // Form2 Dialog = new Form2();
           byte counter_trying_read = 0;
           byte ReadSerialChar;
           StringBuilder Voltage = new StringBuilder();
           do
           {
               if (!Send_Read_Command()) return false;
               // Try to Read Answert as Voltage 
               do
               {
                   ReadSerialChar = Read_Serial_Char();
                if ((ReadSerialChar >= 0x30 &&  ReadSerialChar <= 0x39) || ReadSerialChar == '.' ) Voltage.Append(Convert.ToChar(ReadSerialChar)); //Only Didital Char

               }  while (ReadSerialChar != 0x0d && ReadSerialChar != 0x00);

               if (ReadSerialChar == 0)
               {
                   serialPort1.Close();
                   System.Threading.Thread.Sleep(100);
                   //this.ShowDialog(this.Text);
                  counter_trying_read++;
               }

               if (counter_trying_read == 5)
               {
                  serialPort1.Close();
                    MessageBox.Show("Can't communicate with device!");
                  return false;
               }


           }  while (ReadSerialChar == 0x00);


                      if (Voltage.ToString () !="")
                      {
                          //float voltage = (Convert.ToInt32(Voltage.ToString())) * 5.0f / 1024.0f;
                         // Voltage.Remove(Voltage.Length - 1, 1);
                           textBox1.Text = Voltage.ToString ();

                           if (double.TryParse(Voltage.ToString(), out volt))
                           {
                               chart1.Series["Volt"].Points.AddY(volt);
                           }

                          // if (chart1.Series["Volt"].Points.Count > 10)
                         //  {
                          //     chart1.Series["Volt"].Points.RemoveAt(10);
                          // }

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
            

            //chart1.Series["Volt"].Points.AddXY(1, volt);


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

        private void button1_Click_1(object sender, EventArgs e)
        {

            for (double i = 1; i <= 10; i+=0.1)
            {
                chart1.Series["Volt"].Points.AddY(i);


            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            timer1.Stop();

        }

        private byte Read_Serial_Char ()
        {
            //ReadSerialChar = 0;
            byte ReadSerialChar;
            try
            {
                ReadSerialChar = (byte)serialPort1.ReadByte();             
            }
            catch (TimeoutException)
            {        
              return 0;
            }        
            return ReadSerialChar;
          // return false;
        }


        private bool Send_Read_Command()
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
             return true;
             }
    }
}
