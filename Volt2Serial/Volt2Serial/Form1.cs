using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using Microsoft.Win32;

namespace Volt2Serial
{    
    public partial class Form1 : Form
    {

        string RxBuff1;
        string RXBuff2;
        System.IO.StreamReader oRead1;
        System.IO.StreamWriter oWrite;
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
            
            foreach ( string  Port   in  System.IO.Ports.SerialPort.GetPortNames())
            {
              comboBox1.Items.Add(Port);

            }
             comboBox1.SelectedIndex = 0;
             PortName = comboBox1.Text;
          
            
           
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
            button1.Text = PortName;
             if(PortName !="")
             {
                 serialPort1.ReadTimeout = 500;
                 serialPort1.BaudRate = 9600;
                 //serialPort1.Parity = 'None';
                 serialPort1.PortName = PortName;
             }

        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (serialPort1.IsOpen == true)
            {
              MessageBox.Show (String.Format("Port {0} already is open !", PortName)  );

            //  Console.WriteLine("Port %0 already is open!", PortName);
            }

            else 
            {

             if (PortName != "" )
            {
                try
                {
                    serialPort1.Open(); 
                }
                 catch (TimeoutException )
                     {
                       MessageBox .Show ("Port is busy or missing",PortName );
                     }
                

            }

            }
            
        }
    }
}
