using AForge;
using AForge.Imaging;
using AForge.Imaging.Filters;
using AForge.Math.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sweeper_Bot
{
    public partial class SweeperBot : Form
    {

        public SweeperBot()
        {
            InitializeComponent();
        }

        private async void Button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            await Start();
        }

        private async Task Start()
        {
            MineSweeper mineSweeper = new MineSweeper();
            //await Task.Run(() => mineSweeper.Begin());
            await Task.Run(() =>
            {
                mineSweeper.Begin(toolStripStatusLabel1);
            });
            button1.Enabled = true;
        }
    }
}
