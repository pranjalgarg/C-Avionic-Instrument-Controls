/*****************************************************************************/
/* Project  : AvionicsInstrumentControlDemo                                  */
/* File     : DemoWondow.cs                                                  */
/* Version  : 1                                                              */
/* Language : C#                                                             */
/* Summary  : Start window of the project, use to test the instruments       */
/* Creation : 30/06/2008                                                     */
/* Autor    : Guillaume CHOUTEAU                                             */
/* History  :                                                                */
/*****************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using IronXL;

namespace AvionicsInstrumentControlDemo
{
    public partial class DemoWinow : Form
    {
        public DemoWinow()
        {
            InitializeComponent();
        }

        private void trackBarAirSpeed_Scroll(object sender, EventArgs e)
        {
            airSpeedInstrumentControl1.SetAirSpeedIndicatorParameters(trackBarAirSpeed.Value);
            
        }

        private void trackBarVerticalSpeed_Scroll(object sender, EventArgs e)
        {
            verticalSpeedInstrumentControl1.SetVerticalSpeedIndicatorParameters(trackBarVerticalSpeed.Value);
            if (!backgroundWorker1.IsBusy)
                backgroundWorker1.RunWorkerAsync();
            if (!backgroundWorker2.IsBusy)
                backgroundWorker2.RunWorkerAsync();


        }

        private void trackPitchAngle_Scroll(object sender, EventArgs e)
        {
            horizonInstrumentControl1.SetAttitudeIndicatorParameters(trackPitchAngle.Value, trackBarRollAngle.Value);
        }

        private void trackBarRollAngle_Scroll(object sender, EventArgs e)
        {
            horizonInstrumentControl1.SetAttitudeIndicatorParameters(trackPitchAngle.Value, trackBarRollAngle.Value);
        }

        private void trackBarAltitude_Scroll(object sender, EventArgs e)
        {
            altimeterInstrumentControl1.SetAlimeterParameters(trackBarAltitude.Value);
        }

        private void trackBarHeading_Scroll(object sender, EventArgs e)
        {
            headingIndicatorInstrumentControl1.SetHeadingIndicatorParameters(trackBarHeading.Value);
        }

        private void trackBarTurnRate_Scroll(object sender, EventArgs e)
        {
            turnCoordinatorInstrumentControl1.SetTurnCoordinatorParameters((trackBarTurnRate.Value / 10), trackBarTurnQuality.Value);
        }

        private void trackBarTurnQuality_Scroll(object sender, EventArgs e)
        {
            turnCoordinatorInstrumentControl1.SetTurnCoordinatorParameters((trackBarTurnRate.Value / 10), trackBarTurnQuality.Value);
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            double[] airspeed = ReadAirSpeedData();
            foreach(var speed in airspeed)
            {
                Action action = () => airSpeedInstrumentControl1.SetAirSpeedIndicatorParameters(Convert.ToInt32(speed));
                verticalSpeedInstrumentControl1.Invoke(action);
            }
           
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            double[] pitchAngle = ReadPitchAngleData();
            foreach (var angle in pitchAngle)
            {
                Action action = () => horizonInstrumentControl1.SetAttitudeIndicatorParameters(angle, trackBarRollAngle.Value);
                horizonInstrumentControl1.Invoke(action);
            }
            double[] RollAngle = ReadRollAngleData();
            foreach (var angle in RollAngle)
            {
                Action action = () => horizonInstrumentControl1.SetAttitudeIndicatorParameters(trackPitchAngle.Value, angle);
                horizonInstrumentControl1.Invoke(action);
            }

        }
        private double[] ReadAirSpeedData()
        {
            WorkBook workbook = WorkBook.Load("Actual_Flight_Data_Dump.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            double[] airspeed = new double[3347];
            int i = 0;
            foreach (var cell in sheet["K5:K3347"])
            {
                if (i < 3347)
                   airspeed[i] = Convert.ToDouble(cell.Text);
                   i++;
            }
            return airspeed;
        }
        private double[] ReadPitchAngleData()
        {
            WorkBook workbook = WorkBook.Load("Actual_Flight_Data_Dump.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            double[] pitchAngle = new double[3347];
            int j = 0;
            foreach (var cell in sheet["N5:N3347"])
            {
                if (j < 3347)
                    pitchAngle[j] = Convert.ToDouble(cell.Text);
                j++;
            }
            return pitchAngle;
        }
        private double[] ReadRollAngleData()
        {
            WorkBook workbook = WorkBook.Load("Actual_Flight_Data_Dump.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            double[] RollAngle = new double[3347];
            int k = 0;
            foreach (var cell in sheet["O5:O3347"])
            {
                if (k < 3347)
                    RollAngle[k] = Convert.ToDouble(cell.Text);
                k++;
            }
            return RollAngle;
        }
    }
}