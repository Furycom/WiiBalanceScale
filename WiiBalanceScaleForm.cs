/*********************************************************************************
WiiBalanceScale

MIT License

Copyright (c) 2017-2023 Bernhard Schelling

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
**********************************************************************************/

using System;
using System.Drawing;
using System.Windows.Forms;

namespace WiiBalanceScale
{
    class WiiBalanceScaleForm : Form
    {
        public WiiBalanceScaleForm()
        {
            InitializeComponent();
            try { this.Icon = Icon.ExtractAssociatedIcon(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName); } catch (Exception) { }
        }

        internal Label lblWeight;
        internal Label lblQuality;
        internal Label lblUnit;
        internal Label lblGuidance;
        internal Label lblAdvice;
        internal Label lblWeightVsHeight;
        internal Button btnReset;
        internal GroupBox unitSelector;
        internal RadioButton unitSelectorKg;
        internal RadioButton unitSelectorLb;
        internal RadioButton unitSelectorStone;

        internal Label lblTopLeft;
        internal Label lblTopRight;
        internal Label lblBottomLeft;
        internal Label lblBottomRight;
        internal Label lblLeftRight;
        internal Label lblFrontBack;
        internal Panel pnlCenterOfPressure;
        internal Panel pnlHistory;
        internal Label lblHistorySummary;

        internal ComboBox cmbProfiles;
        internal TextBox txtProfileName;
        internal TextBox txtProfileHeightCm;
        internal Button btnAddProfile;
        internal Button btnClearSession;
        internal Button btnExportCsv;
        internal Button btnExportJson;

        internal CheckBox chkHardwareTest;
        internal Label lblDiagnosticsStatus;
        internal Label lblDiagnosticsDetails;
        internal Label lblCalibrationStatus;
        internal Button btnCaptureCalibration;
        internal Label lblHardwareHint;

        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lblWeight = new Label();
            this.lblQuality = new Label();
            this.lblUnit = new Label();
            this.lblGuidance = new Label();
            this.lblAdvice = new Label();
            this.lblWeightVsHeight = new Label();
            this.btnReset = new Button();
            this.unitSelector = new GroupBox();
            this.unitSelectorKg = new RadioButton();
            this.unitSelectorLb = new RadioButton();
            this.unitSelectorStone = new RadioButton();

            this.cmbProfiles = new ComboBox();
            this.txtProfileName = new TextBox();
            this.txtProfileHeightCm = new TextBox();
            this.btnAddProfile = new Button();
            this.btnClearSession = new Button();
            this.btnExportCsv = new Button();
            this.btnExportJson = new Button();

            this.lblTopLeft = new Label();
            this.lblTopRight = new Label();
            this.lblBottomLeft = new Label();
            this.lblBottomRight = new Label();
            this.lblLeftRight = new Label();
            this.lblFrontBack = new Label();
            this.pnlCenterOfPressure = new Panel();
            this.pnlHistory = new Panel();
            this.lblHistorySummary = new Label();

            this.chkHardwareTest = new CheckBox();
            this.lblDiagnosticsStatus = new Label();
            this.lblDiagnosticsDetails = new Label();
            this.lblCalibrationStatus = new Label();
            this.btnCaptureCalibration = new Button();
            this.lblHardwareHint = new Label();

            GroupBox grpWeight = new GroupBox();
            GroupBox grpProfiles = new GroupBox();
            GroupBox grpSensors = new GroupBox();
            GroupBox grpBalance = new GroupBox();
            GroupBox grpCop = new GroupBox();
            GroupBox grpHistory = new GroupBox();
            GroupBox grpDiagnostics = new GroupBox();
            GroupBox grpCalibration = new GroupBox();

            this.SuspendLayout();

            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1100, 760);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "WiiBalanceScaleForm";
            this.Text = "Wii Balance Scale";

            grpWeight.Location = new Point(10, 8);
            grpWeight.Size = new Size(760, 230);
            grpWeight.Text = "Live weight";

            this.lblWeight.Font = new Font("Lucida Console", 72F);
            this.lblWeight.Location = new Point(8, 20);
            this.lblWeight.Size = new Size(620, 102);
            this.lblWeight.Text = "088.710";
            this.lblWeight.TextAlign = ContentAlignment.MiddleCenter;

            this.lblUnit.Font = new Font("Microsoft Sans Serif", 30F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(128)));
            this.lblUnit.Location = new Point(630, 46);
            this.lblUnit.Size = new Size(120, 60);
            this.lblUnit.Text = "kg";

            Label lblStabilityCaption = new Label();
            lblStabilityCaption.Location = new Point(12, 122);
            lblStabilityCaption.Size = new Size(100, 18);
            lblStabilityCaption.Text = "Stability:";

            this.lblQuality.Font = new Font("Wingdings", 42F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.lblQuality.Location = new Point(108, 116);
            this.lblQuality.Size = new Size(300, 32);
            this.lblQuality.Text = "®®®¡¡";

            this.lblGuidance.Location = new Point(12, 146);
            this.lblGuidance.Size = new Size(740, 20);
            this.lblGuidance.Text = "Stand still for best accuracy. Center your weight on the board.";

            this.lblAdvice.Location = new Point(12, 166);
            this.lblAdvice.Size = new Size(740, 26);
            this.lblAdvice.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold);
            this.lblAdvice.Text = "Advice: Waiting for measurement...";

            this.lblWeightVsHeight.Location = new Point(12, 194);
            this.lblWeightVsHeight.Size = new Size(740, 28);
            this.lblWeightVsHeight.Text = "Weight vs height: add profile height to view this indicator.";

            grpWeight.Controls.Add(this.lblWeight);
            grpWeight.Controls.Add(this.lblUnit);
            grpWeight.Controls.Add(lblStabilityCaption);
            grpWeight.Controls.Add(this.lblQuality);
            grpWeight.Controls.Add(this.lblGuidance);
            grpWeight.Controls.Add(this.lblAdvice);
            grpWeight.Controls.Add(this.lblWeightVsHeight);

            this.unitSelector.Location = new Point(10, 244);
            this.unitSelector.Size = new Size(760, 45);
            this.unitSelector.Text = "Units";
            this.unitSelector.Visible = false;

            this.unitSelectorKg.AutoSize = true;
            this.unitSelectorKg.Location = new Point(14, 18);
            this.unitSelectorKg.Text = "Kilograms (kg)";

            this.unitSelectorLb.AutoSize = true;
            this.unitSelectorLb.Location = new Point(175, 18);
            this.unitSelectorLb.Text = "Pounds (lbs)";

            this.unitSelectorStone.AutoSize = true;
            this.unitSelectorStone.Location = new Point(324, 18);
            this.unitSelectorStone.Text = "Stone/Pounds (st/lbs)";

            this.unitSelector.Controls.Add(this.unitSelectorKg);
            this.unitSelector.Controls.Add(this.unitSelectorLb);
            this.unitSelector.Controls.Add(this.unitSelectorStone);

            grpProfiles.Location = new Point(780, 8);
            grpProfiles.Size = new Size(310, 170);
            grpProfiles.Text = "Profile";

            Label lblProfile = new Label();
            lblProfile.Location = new Point(10, 24);
            lblProfile.Size = new Size(80, 20);
            lblProfile.Text = "Current:";

            this.cmbProfiles.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbProfiles.Location = new Point(92, 20);
            this.cmbProfiles.Size = new Size(208, 22);

            this.txtProfileName.Location = new Point(12, 48);
            this.txtProfileName.Size = new Size(186, 20);

            this.txtProfileHeightCm.Location = new Point(204, 48);
            this.txtProfileHeightCm.Size = new Size(96, 20);

            this.btnAddProfile.Location = new Point(12, 74);
            this.btnAddProfile.Size = new Size(288, 24);
            this.btnAddProfile.Text = "Save profile name + height (cm)";

            this.btnClearSession.Location = new Point(12, 106);
            this.btnClearSession.Size = new Size(92, 24);
            this.btnClearSession.Text = "Clear Session";

            this.btnExportCsv.Location = new Point(112, 106);
            this.btnExportCsv.Size = new Size(92, 24);
            this.btnExportCsv.Text = "Export CSV";

            this.btnExportJson.Location = new Point(210, 106);
            this.btnExportJson.Size = new Size(90, 24);
            this.btnExportJson.Text = "Export JSON";

            Label lblProfileHelp = new Label();
            lblProfileHelp.Location = new Point(12, 136);
            lblProfileHelp.Size = new Size(288, 28);
            lblProfileHelp.Text = "Use name and height in cm (example: 175).";

            grpProfiles.Controls.Add(lblProfile);
            grpProfiles.Controls.Add(this.cmbProfiles);
            grpProfiles.Controls.Add(this.txtProfileName);
            grpProfiles.Controls.Add(this.txtProfileHeightCm);
            grpProfiles.Controls.Add(this.btnAddProfile);
            grpProfiles.Controls.Add(this.btnClearSession);
            grpProfiles.Controls.Add(this.btnExportCsv);
            grpProfiles.Controls.Add(this.btnExportJson);
            grpProfiles.Controls.Add(lblProfileHelp);

            this.btnReset.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(128)));
            this.btnReset.Location = new Point(780, 186);
            this.btnReset.Size = new Size(310, 46);
            this.btnReset.Text = "Zero";
            this.btnReset.UseVisualStyleBackColor = true;

            grpSensors.Location = new Point(10, 298);
            grpSensors.Size = new Size(430, 170);
            grpSensors.Text = "Corner load (kg)";

            this.lblTopLeft.Location = new Point(18, 30);
            this.lblTopLeft.Size = new Size(390, 24);
            this.lblTopLeft.Text = "Top Left: 0.00";
            this.lblTopRight.Location = new Point(18, 60);
            this.lblTopRight.Size = new Size(390, 24);
            this.lblTopRight.Text = "Top Right: 0.00";
            this.lblBottomLeft.Location = new Point(18, 90);
            this.lblBottomLeft.Size = new Size(390, 24);
            this.lblBottomLeft.Text = "Bottom Left: 0.00";
            this.lblBottomRight.Location = new Point(18, 120);
            this.lblBottomRight.Size = new Size(390, 24);
            this.lblBottomRight.Text = "Bottom Right: 0.00";

            grpSensors.Controls.Add(this.lblTopLeft);
            grpSensors.Controls.Add(this.lblTopRight);
            grpSensors.Controls.Add(this.lblBottomLeft);
            grpSensors.Controls.Add(this.lblBottomRight);

            grpBalance.Location = new Point(450, 298);
            grpBalance.Size = new Size(320, 170);
            grpBalance.Text = "Balance";

            this.lblLeftRight.Location = new Point(18, 38);
            this.lblLeftRight.Size = new Size(280, 42);
            this.lblLeftRight.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold);
            this.lblLeftRight.Text = "Left / Right balance: 50.0% / 50.0%";

            this.lblFrontBack.Location = new Point(18, 84);
            this.lblFrontBack.Size = new Size(280, 42);
            this.lblFrontBack.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold);
            this.lblFrontBack.Text = "Front / Back balance: 50.0% / 50.0%";

            grpBalance.Controls.Add(this.lblLeftRight);
            grpBalance.Controls.Add(this.lblFrontBack);

            grpCop.Location = new Point(780, 238);
            grpCop.Size = new Size(310, 230);
            grpCop.Text = "Pressure point";

            this.pnlCenterOfPressure.BorderStyle = BorderStyle.FixedSingle;
            this.pnlCenterOfPressure.Location = new Point(14, 24);
            this.pnlCenterOfPressure.Size = new Size(280, 170);
            this.pnlCenterOfPressure.BackColor = Color.White;

            Label lblCopHelp = new Label();
            lblCopHelp.Location = new Point(14, 198);
            lblCopHelp.Size = new Size(280, 24);
            lblCopHelp.Text = "Crosshair is center. Dot shows pressure point.";

            grpCop.Controls.Add(this.pnlCenterOfPressure);
            grpCop.Controls.Add(lblCopHelp);

            grpHistory.Location = new Point(10, 476);
            grpHistory.Size = new Size(760, 272);
            grpHistory.Text = "Session trend";

            this.pnlHistory.BorderStyle = BorderStyle.FixedSingle;
            this.pnlHistory.Location = new Point(12, 24);
            this.pnlHistory.Size = new Size(736, 180);
            this.pnlHistory.BackColor = Color.White;

            this.lblHistorySummary.Location = new Point(12, 210);
            this.lblHistorySummary.Size = new Size(736, 52);
            this.lblHistorySummary.Text = "No measurements yet.";

            grpHistory.Controls.Add(this.pnlHistory);
            grpHistory.Controls.Add(this.lblHistorySummary);

            grpDiagnostics.Location = new Point(780, 476);
            grpDiagnostics.Size = new Size(310, 170);
            grpDiagnostics.Text = "Hardware check";

            this.chkHardwareTest.Location = new Point(12, 22);
            this.chkHardwareTest.Size = new Size(286, 22);
            this.chkHardwareTest.Text = "Enable hardware check mode";

            this.lblHardwareHint.Location = new Point(12, 44);
            this.lblHardwareHint.Size = new Size(286, 30);
            this.lblHardwareHint.Text = "Helps detect weak corners and idle drift.";

            this.lblDiagnosticsStatus.Location = new Point(12, 78);
            this.lblDiagnosticsStatus.Size = new Size(286, 22);
            this.lblDiagnosticsStatus.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold);
            this.lblDiagnosticsStatus.Text = "Hardware status: Waiting for data";

            this.lblDiagnosticsDetails.Location = new Point(12, 102);
            this.lblDiagnosticsDetails.Size = new Size(286, 60);

            grpDiagnostics.Controls.Add(this.chkHardwareTest);
            grpDiagnostics.Controls.Add(this.lblHardwareHint);
            grpDiagnostics.Controls.Add(this.lblDiagnosticsStatus);
            grpDiagnostics.Controls.Add(this.lblDiagnosticsDetails);

            grpCalibration.Location = new Point(780, 650);
            grpCalibration.Size = new Size(310, 98);
            grpCalibration.Text = "Zero calibration";

            this.btnCaptureCalibration.Location = new Point(12, 24);
            this.btnCaptureCalibration.Size = new Size(286, 24);
            this.btnCaptureCalibration.Text = "Capture zero offsets (board unloaded)";

            this.lblCalibrationStatus.Location = new Point(12, 54);
            this.lblCalibrationStatus.Size = new Size(286, 34);
            this.lblCalibrationStatus.Text = "No zero offsets captured this session.";

            grpCalibration.Controls.Add(this.btnCaptureCalibration);
            grpCalibration.Controls.Add(this.lblCalibrationStatus);

            this.Controls.Add(grpWeight);
            this.Controls.Add(this.unitSelector);
            this.Controls.Add(grpProfiles);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(grpSensors);
            this.Controls.Add(grpBalance);
            this.Controls.Add(grpCop);
            this.Controls.Add(grpHistory);
            this.Controls.Add(grpDiagnostics);
            this.Controls.Add(grpCalibration);

            this.ResumeLayout(false);
        }
    }
}
