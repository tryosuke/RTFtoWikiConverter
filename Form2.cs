using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public void SetProgressMax(int value)
        {
            progressBar1.Maximum = value;
            this.Refresh();
        }
        public void SetProgressBar(int value)
        {
            
            progressBar1.Value += value;
        }
    }
}
