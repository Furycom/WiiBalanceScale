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
        internal Label lblSessionInfo;
        internal Label lblSessionSummary;
        internal Label lblSessionComparison;
        internal Label lblTrendHighlights;
        internal Label lblPatternSummary;
        internal Label lblSessionReview;
        internal Panel pnlWeightTrend;
        internal Panel pnlLeftRightTrend;
        internal Panel pnlFrontBackTrend;
        internal Panel pnlStabilityTrend;
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

        internal ComboBox cmbProfiles;
        internal TextBox txtProfileName;
        internal TextBox txtProfileHeightCm;
        internal Button btnAddProfile;
        internal Button btnClearSession;
        internal Button btnExportCsv;
        internal Button btnExportJson;

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
            this.lblSessionInfo = new Label();
            this.lblSessionSummary = new Label();
            this.lblSessionComparison = new Label();
            this.lblTrendHighlights = new Label();
            this.lblPatternSummary = new Label();
            this.lblSessionReview = new Label();
            this.pnlWeightTrend = new Panel();
            this.pnlLeftRightTrend = new Panel();
            this.pnlFrontBackTrend = new Panel();
            this.pnlStabilityTrend = new Panel();
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

            GroupBox grpWeight = new GroupBox();
            GroupBox grpProfiles = new GroupBox();
            GroupBox grpSensors = new GroupBox();
            GroupBox grpBalance = new GroupBox();
            GroupBox grpPressurePoint = new GroupBox();
            GroupBox grpTrends = new GroupBox();

            this.SuspendLayout();

            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1040, 760);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "WiiBalanceScaleForm";
            this.Text = "Wii Balance Scale";

            grpWeight.Location = new Point(10, 8);
            grpWeight.Size = new Size(690, 360);
            grpWeight.Text = "Live weight";

            this.lblWeight.Font = new Font("Lucida Console", 68F);
            this.lblWeight.Location = new Point(8, 20);
            this.lblWeight.Size = new Size(550, 98);
            this.lblWeight.Text = "088.710";
            this.lblWeight.TextAlign = ContentAlignment.MiddleCenter;

            this.lblUnit.Font = new Font("Microsoft Sans Serif", 28F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(128)));
            this.lblUnit.Location = new Point(560, 42);
            this.lblUnit.Size = new Size(110, 54);
            this.lblUnit.Text = "kg";

            Label lblStabilityCaption = new Label();
            lblStabilityCaption.Location = new Point(12, 118);
            lblStabilityCaption.Size = new Size(100, 18);
            lblStabilityCaption.Text = "Stability:";

            this.lblQuality.Font = new Font("Wingdings", 40F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.lblQuality.Location = new Point(108, 112);
            this.lblQuality.Size = new Size(290, 32);
            this.lblQuality.Text = "®®®¡¡";

            this.lblGuidance.Location = new Point(12, 145);
            this.lblGuidance.Size = new Size(670, 20);
            this.lblGuidance.Text = "Stand still for best accuracy. Center your weight on the board.";

            this.lblAdvice.Location = new Point(12, 166);
            this.lblAdvice.Size = new Size(670, 24);
            this.lblAdvice.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold);
            this.lblAdvice.Text = "Advice: Waiting for measurement...";

            this.lblWeightVsHeight.Location = new Point(12, 192);
            this.lblWeightVsHeight.Size = new Size(670, 24);
            this.lblWeightVsHeight.Text = "Weight vs height: add profile height to view this indicator.";

            this.lblSessionInfo.Location = new Point(12, 218);
            this.lblSessionInfo.Size = new Size(670, 30);
            this.lblSessionInfo.Text = "Session samples: 0";

            this.lblSessionSummary.Location = new Point(12, 246);
            this.lblSessionSummary.Size = new Size(670, 18);
            this.lblSessionSummary.Text = "Session summary: waiting for enough samples...";

            this.lblSessionComparison.Location = new Point(12, 266);
            this.lblSessionComparison.Size = new Size(670, 18);
            this.lblSessionComparison.Text = "Compared with your previous session: not available yet.";

            this.lblTrendHighlights.Location = new Point(12, 286);
            this.lblTrendHighlights.Size = new Size(670, 18);
            this.lblTrendHighlights.Text = "Recent trend: not enough history yet.";

            this.lblPatternSummary.Location = new Point(12, 306);
            this.lblPatternSummary.Size = new Size(670, 18);
            this.lblPatternSummary.Text = "Repeated pattern: not enough strong history yet.";

            this.lblSessionReview.Location = new Point(12, 326);
            this.lblSessionReview.Size = new Size(670, 26);
            this.lblSessionReview.Text = "Session review: waiting for enough samples...";

            grpWeight.Controls.Add(this.lblWeight);
            grpWeight.Controls.Add(this.lblUnit);
            grpWeight.Controls.Add(lblStabilityCaption);
            grpWeight.Controls.Add(this.lblQuality);
            grpWeight.Controls.Add(this.lblGuidance);
            grpWeight.Controls.Add(this.lblAdvice);
            grpWeight.Controls.Add(this.lblWeightVsHeight);
            grpWeight.Controls.Add(this.lblSessionInfo);
            grpWeight.Controls.Add(this.lblSessionSummary);
            grpWeight.Controls.Add(this.lblSessionComparison);
            grpWeight.Controls.Add(this.lblTrendHighlights);
            grpWeight.Controls.Add(this.lblPatternSummary);
            grpWeight.Controls.Add(this.lblSessionReview);

            this.unitSelector.Location = new Point(10, 372);
            this.unitSelector.Size = new Size(690, 45);
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

            grpProfiles.Location = new Point(710, 8);
            grpProfiles.Size = new Size(320, 178);
            grpProfiles.Text = "Profile";

            Label lblProfile = new Label();
            lblProfile.Location = new Point(10, 24);
            lblProfile.Size = new Size(80, 20);
            lblProfile.Text = "Current:";

            this.cmbProfiles.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbProfiles.Location = new Point(92, 20);
            this.cmbProfiles.Size = new Size(218, 22);

            this.txtProfileName.Location = new Point(12, 48);
            this.txtProfileName.Size = new Size(196, 20);

            this.txtProfileHeightCm.Location = new Point(214, 48);
            this.txtProfileHeightCm.Size = new Size(96, 20);

            this.btnAddProfile.Location = new Point(12, 74);
            this.btnAddProfile.Size = new Size(298, 24);
            this.btnAddProfile.Text = "Save profile name + height (cm)";

            this.btnClearSession.Location = new Point(12, 106);
            this.btnClearSession.Size = new Size(94, 24);
            this.btnClearSession.Text = "Clear Session";

            this.btnExportCsv.Location = new Point(114, 106);
            this.btnExportCsv.Size = new Size(94, 24);
            this.btnExportCsv.Text = "Export CSV";

            this.btnExportJson.Location = new Point(216, 106);
            this.btnExportJson.Size = new Size(94, 24);
            this.btnExportJson.Text = "Export JSON";

            Label lblProfileHelp = new Label();
            lblProfileHelp.Location = new Point(12, 136);
            lblProfileHelp.Size = new Size(298, 32);
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
            this.btnReset.Location = new Point(710, 194);
            this.btnReset.Size = new Size(320, 48);
            this.btnReset.Text = "Zero";
            this.btnReset.UseVisualStyleBackColor = true;

            grpSensors.Location = new Point(10, 422);
            grpSensors.Size = new Size(350, 220);
            grpSensors.Text = "Corner load (kg)";

            this.lblTopLeft.Location = new Point(18, 34);
            this.lblTopLeft.Size = new Size(320, 26);
            this.lblTopLeft.Text = "Top left: 0.00";

            this.lblTopRight.Location = new Point(18, 72);
            this.lblTopRight.Size = new Size(320, 26);
            this.lblTopRight.Text = "Top right: 0.00";

            this.lblBottomLeft.Location = new Point(18, 110);
            this.lblBottomLeft.Size = new Size(320, 26);
            this.lblBottomLeft.Text = "Bottom left: 0.00";

            this.lblBottomRight.Location = new Point(18, 148);
            this.lblBottomRight.Size = new Size(320, 26);
            this.lblBottomRight.Text = "Bottom right: 0.00";

            grpSensors.Controls.Add(this.lblTopLeft);
            grpSensors.Controls.Add(this.lblTopRight);
            grpSensors.Controls.Add(this.lblBottomLeft);
            grpSensors.Controls.Add(this.lblBottomRight);

            grpBalance.Location = new Point(370, 422);
            grpBalance.Size = new Size(330, 220);
            grpBalance.Text = "Balance";

            this.lblLeftRight.Location = new Point(18, 42);
            this.lblLeftRight.Size = new Size(300, 56);
            this.lblLeftRight.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold);
            this.lblLeftRight.Text = "Left / Right balance: 50.0% / 50.0%";

            this.lblFrontBack.Location = new Point(18, 112);
            this.lblFrontBack.Size = new Size(300, 56);
            this.lblFrontBack.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold);
            this.lblFrontBack.Text = "Front / Back balance: 50.0% / 50.0%";

            grpBalance.Controls.Add(this.lblLeftRight);
            grpBalance.Controls.Add(this.lblFrontBack);

            grpPressurePoint.Location = new Point(710, 250);
            grpPressurePoint.Size = new Size(320, 296);
            grpPressurePoint.Text = "Pressure point";

            this.pnlCenterOfPressure.BorderStyle = BorderStyle.FixedSingle;
            this.pnlCenterOfPressure.Location = new Point(14, 24);
            this.pnlCenterOfPressure.Size = new Size(292, 234);
            this.pnlCenterOfPressure.BackColor = Color.White;

            Label lblPressureHelp = new Label();
            lblPressureHelp.Location = new Point(14, 262);
            lblPressureHelp.Size = new Size(292, 24);
            lblPressureHelp.Text = "Crosshair is center. Dot shows pressure point.";

            grpPressurePoint.Controls.Add(this.pnlCenterOfPressure);
            grpPressurePoint.Controls.Add(lblPressureHelp);

            grpTrends.Location = new Point(10, 646);
            grpTrends.Size = new Size(1020, 106);
            grpTrends.Text = "Session trends (selected profile)";

            Label lblWeightTrend = new Label();
            lblWeightTrend.Location = new Point(10, 20);
            lblWeightTrend.Size = new Size(240, 16);
            lblWeightTrend.Text = "Weight trend";

            this.pnlWeightTrend.Location = new Point(10, 38);
            this.pnlWeightTrend.Size = new Size(240, 58);
            this.pnlWeightTrend.BorderStyle = BorderStyle.FixedSingle;
            this.pnlWeightTrend.BackColor = Color.White;

            Label lblLeftRightTrend = new Label();
            lblLeftRightTrend.Location = new Point(264, 20);
            lblLeftRightTrend.Size = new Size(240, 16);
            lblLeftRightTrend.Text = "Left/Right balance trend";

            this.pnlLeftRightTrend.Location = new Point(264, 38);
            this.pnlLeftRightTrend.Size = new Size(240, 58);
            this.pnlLeftRightTrend.BorderStyle = BorderStyle.FixedSingle;
            this.pnlLeftRightTrend.BackColor = Color.White;

            Label lblFrontBackTrend = new Label();
            lblFrontBackTrend.Location = new Point(518, 20);
            lblFrontBackTrend.Size = new Size(240, 16);
            lblFrontBackTrend.Text = "Front/Back balance trend";

            this.pnlFrontBackTrend.Location = new Point(518, 38);
            this.pnlFrontBackTrend.Size = new Size(240, 58);
            this.pnlFrontBackTrend.BorderStyle = BorderStyle.FixedSingle;
            this.pnlFrontBackTrend.BackColor = Color.White;

            Label lblStabilityTrend = new Label();
            lblStabilityTrend.Location = new Point(772, 20);
            lblStabilityTrend.Size = new Size(240, 16);
            lblStabilityTrend.Text = "Stability trend";

            this.pnlStabilityTrend.Location = new Point(772, 38);
            this.pnlStabilityTrend.Size = new Size(240, 58);
            this.pnlStabilityTrend.BorderStyle = BorderStyle.FixedSingle;
            this.pnlStabilityTrend.BackColor = Color.White;

            grpTrends.Controls.Add(lblWeightTrend);
            grpTrends.Controls.Add(this.pnlWeightTrend);
            grpTrends.Controls.Add(lblLeftRightTrend);
            grpTrends.Controls.Add(this.pnlLeftRightTrend);
            grpTrends.Controls.Add(lblFrontBackTrend);
            grpTrends.Controls.Add(this.pnlFrontBackTrend);
            grpTrends.Controls.Add(lblStabilityTrend);
            grpTrends.Controls.Add(this.pnlStabilityTrend);

            this.Controls.Add(grpWeight);
            this.Controls.Add(this.unitSelector);
            this.Controls.Add(grpProfiles);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(grpSensors);
            this.Controls.Add(grpBalance);
            this.Controls.Add(grpPressurePoint);
            this.Controls.Add(grpTrends);

            this.ResumeLayout(false);
        }
    }
}
