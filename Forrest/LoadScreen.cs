using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace Forrest {
    public partial class LoadScreen : Form {

        public LoadScreen() {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e) {

            //update progress bar and status
            progressBar1.Value = Globals.LoadProgress;
            label1.Text = Globals.Status;

            if(Globals.LoadProgress > 99) {
                //done loading
                //make sure timer doesn't fire again
                timer1.Enabled = false;

                Thread.Sleep(200);

                ///close load screen
                this.Close();

                //abort thread
                Thread.CurrentThread.Abort();
            }
        }


    }

}
