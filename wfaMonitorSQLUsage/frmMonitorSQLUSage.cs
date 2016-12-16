using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Timers;
using System.IO;

namespace wfaMonitorSQLUsage
{
    public partial class frmMonitorSQLUSage : Form
    {
        int minutes=15;
        SqlConnection conn;
        System.Timers.Timer aTimer;

        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            checkSQLUsage();
        }

        private void checkSQLUsage() {
            StringBuilder sb = new StringBuilder();
            this.conn = new SqlConnection(this.txtStringConn.Text);
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;
            string path = Application.StartupPath.ToString();
            string filename = DateTime.Now.ToString()+".csv";
            filename=filename.Replace("/","_");
            filename = filename.Replace(":", "_");
            filename = filename.Replace(" ", "_");
            //MessageBox.Show(path);
            StreamWriter sw = new StreamWriter(path+ "\\" + filename);
            this.log("File to be generated under " + path + "\\" + filename);
            string [] columnNames;
            cmd.CommandText = "sp_who";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conn;

            conn.Open();

            reader = cmd.ExecuteReader();
            
            columnNames=new string[(int)reader.FieldCount];

            for(int i=0; i<reader.FieldCount;i++){
                columnNames[i] = (string)reader.GetName(i); 
            }

            //Create headers
            sb.Append(string.Join(",", columnNames));

            //Append Line
            sb.AppendLine();
            while (reader.Read()) {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    string value = reader[i].ToString();
                    if (value.Contains(","))
                        value = "\"" + value + "\"";

                    
                    sb.Append(value.Replace(Environment.NewLine, " ").Trim() + ",");
                }
                sb.Length--; // Remove the last comma
                sb.AppendLine();
            }
            conn.Close();
            sw.Write(sb.ToString());
            sw.Close();
            this.log("File generated");
        }
        public frmMonitorSQLUSage()
        {
            InitializeComponent();
        }
        
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            stopToolStripMenuItem.Enabled = true;
            runToolStripMenuItem.Enabled = false;
            exitToolStripMenuItem.Enabled = false;
            
            this.minutes = (int)this.nUD.Value;
            this.tsStatus.Text = "Running...";
            this.log("Running... every "+this.minutes.ToString()+ " minute(s)");
            try
            {
                this.run();
            }
            catch (Exception ex) {
                aTimer.Stop();
                MessageBox.Show(ex.Message);
            }
            
        }

        private void run()
        {
            aTimer = new System.Timers.Timer();
            this.minutes = (int)this.nUD.Value;

            // Hook up the Elapsed event for the timer.
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);

            // Set the Interval to 2 seconds (2000 milliseconds).
            aTimer.Interval = (1000 * 60 * this.minutes);
            aTimer.Enabled = true;

        }

        private void frmMonitorSQLUSage_Load(object sender, EventArgs e)
        {
			txtStringConn.Text = "Server=SQLServer;Database=master;Trusted_Connection=True;";
            tsUser.Text = Environment.UserDomainName.ToString() + "\\" + Environment.UserName ;
        }

        private void saveConfigurationToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            aTimer.Stop();
            tsStatus.Text = "Stopped";
            stopToolStripMenuItem.Enabled = false;
            runToolStripMenuItem.Enabled = true;
            exitToolStripMenuItem.Enabled = true;
        }

        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void log(string value) {
            
            if (this.txtLog.InvokeRequired)
            {
                // It's on a different thread, so use Invoke.
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke
                    (d, new object[] { value + " (Invoke)" });
            }
            else
            {
                // It's on the same thread, no need for Invoke 
                //this.textBox1.Text = text + " (No Invoke)";
                txtLog.Text += DateTime.Now.ToString() + " - " + value + Environment.NewLine;
            }
        }

        // This delegate enables asynchronous calls for setting
        // the text property on a TextBox control.
        delegate void SetTextCallback(string value);

        // This method is passed in to the SetTextCallBack delegate 
        // to set the Text property of textBox1. 
        private void SetText(string value)
        {
            txtLog.Text += DateTime.Now.ToString() + " - " + value + Environment.NewLine;
        }
    }
}
