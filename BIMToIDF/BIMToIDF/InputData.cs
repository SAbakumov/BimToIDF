using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
namespace BIMToIDF
{
    public partial class InputData : Form
    {
        public bool differentFloors;
        public int numFloors;
        public int numSamples = 1000;
        public Dictionary<string, double[]> buildingConstruction = new Dictionary<string, double[]>();
        public Dictionary<String, double[]> windowConstruction = new Dictionary<string, double[]>();

        //oPH, iHG, venti
        public double[] operatingHours;
        public double[] iHG;
        public double[] infiltration;

        public InputData()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            numFloors = (int)number.Value;
            differentFloors = diffNFloors.Checked;
            numSamples = (int)sampleCount.Value;
            windowConstruction.Add("wWR1", new double[] { (double)wWR1_mean.Value, .01 * (double)wWR1_var.Value });
            windowConstruction.Add("wWR2", new double[] { (double)wWR1_mean.Value, .01 * (double)wWR1_var.Value });
            windowConstruction.Add("wWR3", new double[] { (double)wWR1_mean.Value, .01 * (double)wWR1_var.Value });
            windowConstruction.Add("wWR4", new double[] { (double)wWR1_mean.Value, .01 * (double)wWR1_var.Value });

            buildingConstruction.Add("uWall", new double[] { (double) uWall_mean.Value, .01 * (double) uWall_var.Value });
            buildingConstruction.Add("uGFloor", new double[] { (double)uGFloor_mean.Value, .01 * (double)uGFloor_var.Value });
            buildingConstruction.Add("uRoof", new double[] { (double)uRoof_mean.Value, .01 * (double)uRoof_var.Value });
            buildingConstruction.Add("uIFloor", new double[] { (double)uIFloor_Mean.Value, .01 * (double)uIFloor_var.Value });
            buildingConstruction.Add("uWindow", new double[] { (double)uWindow_mean.Value, .01 * (double)uWindow_var.Value });
            buildingConstruction.Add("gWindow", new double[] { (double)gWindow_mean.Value, .01 * (double)gWindow_var.Value });
            buildingConstruction.Add("HCFloor", new double[] { (double)HCFloor_mean.Value, .01 * (double)HCFloor_var.Value });

            buildingConstruction.Add("BEff", new double[] { (double)bEffMean.Value, .01 * (double)bEffVar.Value });
            buildingConstruction.Add("CCOP", new double[] { (double)CCOPMean.Value, .01 * (double)CCOPVar.Value });

            operatingHours = new double[] { (double)oPH_mean.Value, .01 * (double) oPH_var.Value };
            iHG = new double[] { (double)iHG_mean.Value, .01 * (double)iHG_var.Value };
            infiltration = new double[] { (double)vent_mean.Value , .01 * (double)vent_var.Value };

            this.Close();
        }

        private void diffNFloors_CheckedChanged(object sender, EventArgs e)
        {
            if (diffNFloors.Checked)
            {
                this.label18.Enabled = false;
                this.number.Enabled = false;
            }
            else
            {
                this.label18.Enabled = true;
                this.number.Enabled = true;
            }
        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void CCOPMean_ValueChanged(object sender, EventArgs e)
        {

        }

        private void SampleCount_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
