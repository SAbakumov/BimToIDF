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
        public bool cancel = false;

        public int numFloors;
        public int numSamples = 1000;
        public IDFFile.ProbabilisticBuildingConstruction pBuildingConstruction;
        public IDFFile.ProbabilisticWWR pWindowConstruction;
        public IDFFile.ProbabilisticBuildingOperation pBuildingOperation;

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
            numSamples = (int)sampleCount.Value;
            pWindowConstruction = new IDFFile.ProbabilisticWWR()
            {
                north = GetMinMax(wWR1_mean, wWR1_var),
                east = GetMinMax(wWR2_mean, wWR2_var),
                south = GetMinMax(wWR3_mean, wWR3_var),
                west = GetMinMax(wWR4_mean, wWR4_var)
            };

            pBuildingConstruction = new IDFFile.ProbabilisticBuildingConstruction
            {
                uWall = GetMinMax(uWall_mean, uWall_var),
                uGFloor = GetMinMax(uGFloor_mean,  uGFloor_var),
                uRoof = GetMinMax(uRoof_mean,  uRoof_var),
                uIFloor = GetMinMax(uIFloor_Mean,  uIFloor_var),
                uIWall = GetMinMax(uWall_mean, uWall_var),
                uWindow = GetMinMax(uWindow_mean,  uWindow_var),
                gWindow = GetMinMax(gWindow_mean,  gWindow_var),
                hcSlab = GetMinMax(HCFloor_mean,  HCFloor_var),
                infiltration = GetMinMax(vent_mean,  vent_var)
            };

            pBuildingOperation = new IDFFile.ProbabilisticBuildingOperation()
            {
                boilerEfficiency = GetMinMax(bEffMean,  bEffVar),
                chillerCOP = GetMinMax(CCOPMean,  CCOPVar),

                operatingHours = GetMinMax(oPH_mean,  oPH_var),
                internalHeatGain = GetMinMax(iHG_mean,  iHG_var)
            };
            this.Close();
        }

        private double[] GetMinMax(NumericUpDown mean, NumericUpDown var)
        {
            return new double[] { (double)mean.Value * (1 - .01 * (double)var.Value), (double)mean.Value * (1+ .01 * (double)var.Value) };
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

        private void InputData_Load(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            cancel = true;
            this.Close();
        }
    }
}
