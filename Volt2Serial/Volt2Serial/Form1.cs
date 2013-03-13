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
        bool ReadCommPort ;
        string message;
        /*
         [System.Runtime.InteropServices.DllImport("winmm.dll")] 
        public  timeGetTime_
            Lib "winmm.dll" () as long 
        */
        string ReadValue = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop", null );
        string DesktopPatch;
        string PortName;

        public Form1()
        {
            InitializeComponent();

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

    private void   ReadData ()
       {
           message = "";
           ReadCommPort = false;
           // MessageBox.Show("Comm Port is not select");
          
            if(PortName == "" || serialPort1.ReadTimeout==-1 ) 
            {
               // MessageBox.Show ("Comm Port is not select");
                message = "Comm Port is not select";
                return ;
            }

            if (serialPort1.IsOpen != true)
            {
                try
                {
                    serialPort1.Open();
                }
                catch (Exception )
                {
                    //MessageBox.Show("Port is busy or missing");
                    message = "Can't open port";
                    return ;
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
                          catch (TimeoutException ex )
                          {
                            serialPort1.Close();
                          // DialogResult result= MessageBox.Show (string.Format("The Device on port {0} not responding",PortName ),"Ehoo",MessageBoxButtons.OK);
                            message = string.Format("The Device on port {0} not responding", PortName);
                            return ;
                           }

                         
                          

                         Voltage.Append(Convert.ToChar (ReadSerialChar));
                         
                      } while (ReadSerialChar != 0x0d);
                      if (Voltage.Length != 0)
                      {
                          float voltage = (Convert.ToInt32(Voltage.ToString())) * 5.0f / 1024.0f;
                          textBox1.Text = Convert.ToString(voltage);
                      }
                 serialPort1.Close();
                 ReadCommPort = true;
           }


        private void button3_Click(object sender, EventArgs e)
        {
            ReadData();
        } 

        private void timer1_Tick(object sender, EventArgs e)
        {    
            if (button4.ForeColor == Color.Green)
            {
                button4.ForeColor = Color.Red;
            }
            else
            {
                button4.ForeColor = Color.Green;
            }
            //ReadCommPort = true;
            ReadData();
            if (ReadCommPort == false)
            {
                textBox3.Text = message;
                timer1.Stop();
            }
            else
            {
                textBox3.Text = "";
            }
            
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

      public class Mycode 
      {






      }

       
    }
}
