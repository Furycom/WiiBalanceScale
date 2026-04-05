/*********************************************************************************
WiiBalanceScale

MIT License

Copyright (c) 2017-2023 Bernhard Schelling
Copyright (c) 2023 Carl Ansell

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
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using WiimoteLib;
using RadioButton = System.Windows.Forms.RadioButton;

[assembly: System.Reflection.AssemblyTitle("WiiBalanceScale")]
[assembly: System.Reflection.AssemblyProduct("WiiBalanceScale")]
[assembly: System.Reflection.AssemblyVersion("1.3.1.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.3.1.0")]
[assembly: System.Runtime.InteropServices.ComVisible(false)]

namespace WiiBalanceScale
{
    internal class WiiBalanceScale
    {
        enum EConnectionError { None, NoBluetoothAdapter, PermissionDenied, NoDeviceFound, WrongDeviceType, ConnectionFailed }
        enum EUnit { Kg, Lb, Stone }

        struct SensorSample
        {
            public float TopLeft;
            public float TopRight;
            public float BottomLeft;
            public float BottomRight;
            public float Total;
        }

        class MeasurementSample
        {
            public DateTime TimestampUtc;
            public string ProfileName;
            public float ProfileHeightCm;
            public float TotalWeightKg;
            public float SensorTopLeftKg;
            public float SensorTopRightKg;
            public float SensorBottomLeftKg;
            public float SensorBottomRightKg;
            public float LeftPercent;
            public float RightPercent;
            public float FrontPercent;
            public float BackPercent;
            public float PressurePointX;
            public float PressurePointY;
            public int StabilityLevel;
        }

        class ProfileInfo
        {
            public string Name;
            public float HeightCm;
        }

        static bool CanShowUnicode = GetCanShowUnicode();
        static char CharFilledStar = (CanShowUnicode ? '\u2739' : '\u00AE');
        static char CharHollowCircle = (CanShowUnicode ? '\u3007' : '\u00A1');
        static char CharHourglass = (CanShowUnicode ? '\u23F3' : '\u0036');
        static WiiBalanceScaleForm f = null;
        static Wiimote bb = null;
        static ConnectionManager cm = null;
        static Timer BoardTimer = null;
        static float ZeroedWeight = 0;
        static float[] History = new float[100];
        static int HistoryBest = 1, HistoryCursor = -1;
        static EUnit SelectedUnit = EUnit.Kg;
        static EConnectionError LastConnectionError = EConnectionError.None;
        static string LastConnectionErrorDetail = "";
        static bool DidTryElevation = false;

        static readonly List<MeasurementSample> SessionMeasurements = new List<MeasurementSample>();
        static readonly List<SensorSample> DiagnosticsWindow = new List<SensorSample>();
        static readonly List<ProfileInfo> Profiles = new List<ProfileInfo>();

        static DateTime LastSessionRecordUtc = DateTime.MinValue;
        static int LastStabilityLevel = 0;
        static PointF LastPressurePoint = new PointF(0.0f, 0.0f);
        static float LastDisplayedWeightKg = 0.0f;
        static float LastLeftPercent = 50.0f;
        static float LastFrontPercent = 50.0f;
        static string LastHardwareSignal = "";

        static bool HasCalibration = false;
        static float CalibrationTopLeftOffset = 0.0f;
        static float CalibrationTopRightOffset = 0.0f;
        static float CalibrationBottomLeftOffset = 0.0f;
        static float CalibrationBottomRightOffset = 0.0f;

        static readonly string ProfilesPath = Path.Combine(Application.StartupPath, "profiles.csv");
        static readonly string LegacyProfilesPath = Path.Combine(Application.StartupPath, "profiles.txt");
        static readonly string CsvHeader = "timestamp_utc,profile,profile_height_cm,total_weight_kg,sensor_top_left_kg,sensor_top_right_kg,sensor_bottom_left_kg,sensor_bottom_right_kg,left_percent,right_percent,front_percent,back_percent,pressure_point_x,pressure_point_y,stability_level";

        static bool GetCanShowUnicode()
        {
            try { return int.Parse(Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuildNumber", "").ToString()) >= 19000; } catch { return false; }
        }

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            f = new WiiBalanceScaleForm();
            f.lblWeight.Text = "";
            f.lblUnit.Text = "";
            f.lblQuality.Text = "";
            if (CanShowUnicode) f.lblQuality.Font = new Font("Segoe UI Symbol", 42F, FontStyle.Regular, GraphicsUnit.Pixel);

            LoadProfiles();
            BindProfilesToUi();

            f.btnReset.Click += btnReset_Click;
            f.btnAddProfile.Click += btnAddProfile_Click;
            f.cmbProfiles.SelectedIndexChanged += cmbProfiles_SelectedIndexChanged;
            f.btnClearSession.Click += btnClearSession_Click;
            f.btnExportCsv.Click += btnExportCsv_Click;
            f.btnExportJson.Click += btnExportJson_Click;
            f.btnCaptureCalibration.Click += btnCaptureCalibration_Click;
            f.chkHardwareTest.CheckedChanged += chkHardwareTest_CheckedChanged;
            f.pnlCenterOfPressure.Paint += pnlCenterOfPressure_Paint;
            f.pnlHistory.Paint += pnlHistory_Paint;

            EventHandler unitRadioButton_Change = unitRadioButton_ChangeHandler;
            f.unitSelectorKg.CheckedChanged += unitRadioButton_Change;
            f.unitSelectorLb.CheckedChanged += unitRadioButton_Change;
            f.unitSelectorStone.CheckedChanged += unitRadioButton_Change;
            f.unitSelectorKg.Checked = true;

            ConnectBalanceBoard(false);
            if (f == null) return;

            BoardTimer = new System.Windows.Forms.Timer();
            BoardTimer.Interval = 50;
            BoardTimer.Tick += new System.EventHandler(BoardTimer_Tick);
            BoardTimer.Start();

            Application.Run(f);
            Shutdown();
        }

        static void Shutdown()
        {
            if (BoardTimer != null) { BoardTimer.Stop(); BoardTimer = null; }
            if (cm != null) { cm.Cancel(); cm = null; }
            CleanupBalanceBoard();
            if (f != null) { if (f.Visible) f.Close(); f = null; }
        }

        static void CleanupBalanceBoard()
        {
            if (bb == null) return;
            try { bb.Disconnect(); } catch { }
            try { bb.Dispose(); } catch { }
            bb = null;
        }

        static void LoadProfiles()
        {
            Profiles.Clear();
            if (File.Exists(ProfilesPath))
            {
                string[] lines = File.ReadAllLines(ProfilesPath);
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i] == null) continue;
                    string line = lines[i].Trim();
                    if (line.Length == 0) continue;
                    string[] parts = line.Split(',');
                    string name = parts[0].Trim();
                    if (name.Length == 0 || FindProfileByName(name) != null) continue;
                    float height = 0.0f;
                    if (parts.Length > 1)
                        float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out height);
                    Profiles.Add(new ProfileInfo() { Name = name, HeightCm = height });
                }
            }
            else if (File.Exists(LegacyProfilesPath))
            {
                string[] lines = File.ReadAllLines(LegacyProfilesPath);
                for (int i = 0; i < lines.Length; i++)
                {
                    string name = (lines[i] == null ? "" : lines[i].Trim());
                    if (name.Length == 0 || FindProfileByName(name) != null) continue;
                    Profiles.Add(new ProfileInfo() { Name = name, HeightCm = 0.0f });
                }
            }

            if (FindProfileByName("Default") == null)
                Profiles.Insert(0, new ProfileInfo() { Name = "Default", HeightCm = 0.0f });
            if (Profiles.Count == 0)
                Profiles.Add(new ProfileInfo() { Name = "Default", HeightCm = 0.0f });
        }

        static void SaveProfiles()
        {
            try
            {
                string[] lines = new string[Profiles.Count];
                for (int i = 0; i < Profiles.Count; i++)
                {
                    ProfileInfo p = Profiles[i];
                    lines[i] = p.Name + "," + p.HeightCm.ToString("0.0", CultureInfo.InvariantCulture);
                }
                File.WriteAllLines(ProfilesPath, lines);
            }
            catch (Exception ex)
            {
                MessageBox.Show(f, "Could not save profiles.\n\n" + ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        static ProfileInfo FindProfileByName(string name)
        {
            for (int i = 0; i < Profiles.Count; i++)
            {
                if (string.Equals(Profiles[i].Name, name, StringComparison.OrdinalIgnoreCase))
                    return Profiles[i];
            }
            return null;
        }

        static void BindProfilesToUi()
        {
            f.cmbProfiles.Items.Clear();
            for (int i = 0; i < Profiles.Count; i++)
                f.cmbProfiles.Items.Add(Profiles[i].Name);
            if (f.cmbProfiles.Items.Count > 0)
                f.cmbProfiles.SelectedIndex = 0;
            ApplyProfileToInputs();
        }

        static ProfileInfo GetSelectedProfile()
        {
            if (f.cmbProfiles.SelectedItem == null) return FindProfileByName("Default");
            ProfileInfo selected = FindProfileByName(f.cmbProfiles.SelectedItem.ToString());
            if (selected != null) return selected;
            return FindProfileByName("Default");
        }

        static void ApplyProfileToInputs()
        {
            ProfileInfo p = GetSelectedProfile();
            if (p == null) return;
            f.txtProfileName.Text = p.Name;
            f.txtProfileHeightCm.Text = (p.HeightCm > 0.0f ? p.HeightCm.ToString("0.0", CultureInfo.InvariantCulture) : "");
        }

        static void SetConnectionError(EConnectionError error, string detail)
        {
            LastConnectionError = error;
            LastConnectionErrorDetail = detail;
        }

        static string GetConnectionErrorMessage()
        {
            string text;
            switch (LastConnectionError)
            {
                case EConnectionError.PermissionDenied:
                    text = "Bluetooth adapter detected but access denied. Try running as administrator.";
                    break;
                case EConnectionError.NoDeviceFound:
                    text = "No Wii Balance Board found. Press SYNC on the device.";
                    break;
                case EConnectionError.WrongDeviceType:
                    text = "A Nintendo device was detected, but it is not a Wii Balance Board.";
                    break;
                case EConnectionError.NoBluetoothAdapter:
                    text = "No compatible bluetooth adapter found.";
                    break;
                default:
                    text = "Connection to the Wii Balance Board failed.";
                    break;
            }
            if (!string.IsNullOrEmpty(LastConnectionErrorDetail))
                text += "\n\nDetails: " + LastConnectionErrorDetail;
            return text;
        }

        static void ApplyScannerStatusToConnectionError()
        {
            ConnectionManager.ScanStatus scanStatus = (cm == null ? null : cm.GetLastScanStatus());
            if (scanStatus == null || scanStatus.Result == ConnectionManager.EScanResult.None)
                return;

            string detail = scanStatus.Detail;
            if (scanStatus.ErrorCode != 0)
                detail += (string.IsNullOrEmpty(detail) ? "" : " ") + "(Error " + scanStatus.ErrorCode.ToString() + ")";

            switch (scanStatus.Result)
            {
                case ConnectionManager.EScanResult.PermissionDenied:
                    SetConnectionError(EConnectionError.PermissionDenied, detail);
                    return;
                case ConnectionManager.EScanResult.BluetoothApiFailure:
                    SetConnectionError(EConnectionError.NoBluetoothAdapter, detail);
                    return;
                case ConnectionManager.EScanResult.NoDeviceFound:
                    SetConnectionError(EConnectionError.NoDeviceFound, detail);
                    return;
            }
        }

        static void ConnectBalanceBoard(bool WasJustConnected)
        {
            bool Connected = true;
            CleanupBalanceBoard();
            try
            {
                bb = new Wiimote();
                bb.Connect();
                bb.SetLEDs(1);
                bb.GetStatus();
            }
            catch (Exception ex)
            {
                Connected = false;
                SetConnectionError((ex is UnauthorizedAccessException ? EConnectionError.PermissionDenied : EConnectionError.ConnectionFailed), ex.Message);
            }

            if (!Connected)
            {
                if (cm == null) cm = new ConnectionManager();
                cm.ConnectNextWiiMote();
                return;
            }

            if (bb.WiimoteState.ExtensionType != ExtensionType.BalanceBoard)
            {
                SetConnectionError(EConnectionError.WrongDeviceType, "Detected extension type: " + bb.WiimoteState.ExtensionType.ToString());
                if (cm == null) cm = new ConnectionManager();
                cm.ConnectNextWiiMote();
                return;
            }

            if (cm != null) { cm.Cancel(); cm = null; }
            LastConnectionError = EConnectionError.None;
            LastConnectionErrorDetail = "";
            DidTryElevation = false;

            f.unitSelector.Visible = true;
            f.lblWeight.Text = "...";
            f.lblQuality.Text = "";
            f.lblUnit.Text = "";
            f.Refresh();

            ZeroedWeight = 0.0f;
            int InitWeightCount = 0;
            for (int CountMax = (WasJustConnected ? 100 : 50); InitWeightCount < CountMax || bb.WiimoteState.BalanceBoardState.WeightKg == 0.0f; InitWeightCount++)
            {
                if (bb.WiimoteState.BalanceBoardState.WeightKg < -200) break;
                ZeroedWeight += bb.WiimoteState.BalanceBoardState.WeightKg;
                bb.GetStatus();
            }
            ZeroedWeight = (InitWeightCount > 0 ? (ZeroedWeight / (float)InitWeightCount) : 0.0f);

            HistoryCursor = HistoryBest = History.Length / 2;
            for (int i = 0; i < History.Length; i++)
                History[i] = (i > HistoryCursor ? float.MinValue : ZeroedWeight);

            DiagnosticsWindow.Clear();
            LastSessionRecordUtc = DateTime.MinValue;
        }

        static void BoardTimer_Tick(object sender, System.EventArgs e)
        {
            if (cm != null)
            {
                if (cm.IsRunning())
                {
                    f.lblWeight.Text = "WAIT...";
                    f.lblQuality.Text = (f.lblQuality.Text.Length >= 5 ? "" : f.lblQuality.Text) + CharHourglass;
                    f.lblAdvice.Text = "Advice: Waiting for board connection...";
                    return;
                }

                if (cm.HadError())
                {
                    ApplyScannerStatusToConnectionError();

                    if (!DidTryElevation)
                    {
                        DidTryElevation = true;
                        try
                        {
                            if (ConnectionManager.ElevateProcessNeedRestart()) { Shutdown(); return; }
                        }
                        catch (Exception ex)
                        {
                            SetConnectionError(EConnectionError.PermissionDenied, ex.Message);
                        }
                    }

                    if (LastConnectionError == EConnectionError.None)
                        SetConnectionError(EConnectionError.NoBluetoothAdapter, "Connection manager reported a bluetooth error without a specific cause.");

                    BoardTimer.Stop();
                    MessageBox.Show(f, GetConnectionErrorMessage(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Shutdown();
                    return;
                }

                if (!cm.DidConnect())
                {
                    ApplyScannerStatusToConnectionError();
                    ConnectionManager.ScanStatus scanStatus = cm.GetLastScanStatus();
                    if (scanStatus != null && scanStatus.Result == ConnectionManager.EScanResult.Cancelled)
                        return;

                    if (LastConnectionError == EConnectionError.None)
                        SetConnectionError(EConnectionError.NoDeviceFound, "Bluetooth scan timed out without finding a Wii Balance Board.");

                    BoardTimer.Stop();
                    MessageBox.Show(f, GetConnectionErrorMessage(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Shutdown();
                    return;
                }

                ConnectBalanceBoard(true);
                return;
            }

            float kg = bb.WiimoteState.BalanceBoardState.WeightKg, HistorySum = 0.0f, MaxHist = kg, MinHist = kg, MaxDiff = 0.0f;
            if (kg < -200)
            {
                ConnectBalanceBoard(false);
                return;
            }

            HistoryCursor++;
            History[HistoryCursor % History.Length] = kg;
            for (HistoryBest = 0; HistoryBest < History.Length; HistoryBest++)
            {
                float HistoryEntry = History[(HistoryCursor + History.Length - HistoryBest) % History.Length];
                if (Math.Abs(MaxHist - HistoryEntry) > 1.0f) break;
                if (Math.Abs(MinHist - HistoryEntry) > 1.0f) break;
                if (HistoryEntry > MaxHist) MaxHist = HistoryEntry;
                if (HistoryEntry < MinHist) MinHist = HistoryEntry;
                float Diff = Math.Max(Math.Abs(HistoryEntry - kg), Math.Abs((HistorySum + HistoryEntry) / (HistoryBest + 1) - kg));
                if (Diff > MaxDiff) MaxDiff = Diff;
                if (Diff > 1.0f) break;
                HistorySum += HistoryEntry;
            }

            if (HistoryBest <= 0) return;
            kg = HistorySum / HistoryBest - ZeroedWeight;
            LastDisplayedWeightKg = kg;

            float accuracy = 1.0f / HistoryBest;
            float weight = (float)Math.Floor(kg / accuracy + 0.5f) * accuracy;

            if (SelectedUnit != EUnit.Kg) weight *= 2.20462262f;
            if (SelectedUnit == EUnit.Stone)
            {
                string sign = weight < 0.0f ? "-" : "";
                weight = Math.Abs(weight);
                f.lblWeight.Text = sign + Math.Floor(weight / 14.0f).ToString("00") + ":" + (weight % 14.0f).ToString("00.0");
                f.lblUnit.Text = "st:lbs";
            }
            else
            {
                f.lblWeight.Text = weight <= -100.0f ? weight.ToString("00.00") : weight.ToString("00.000");
                f.lblUnit.Text = (SelectedUnit != EUnit.Kg ? "lbs" : "kg");
            }

            LastStabilityLevel = ((HistoryBest + 5) / (History.Length / 5));
            f.lblQuality.Text = "";
            for (int i = 0; i < 5; i++)
                f.lblQuality.Text += (i < LastStabilityLevel ? CharFilledStar : CharHollowCircle);

            UpdateLiveSensorUi();
            UpdateWeightVsHeightIndicator();
            UpdateAdviceMessage();
            MaybeRecordSession();
            UpdateHistorySummary();
            f.pnlCenterOfPressure.Invalidate();
            f.pnlHistory.Invalidate();
        }

        static void UpdateLiveSensorUi()
        {
            BalanceBoardSensorsF rawSensors = bb.WiimoteState.BalanceBoardState.SensorValuesKg;
            float tl = rawSensors.TopLeft - (HasCalibration ? CalibrationTopLeftOffset : 0.0f);
            float tr = rawSensors.TopRight - (HasCalibration ? CalibrationTopRightOffset : 0.0f);
            float bl = rawSensors.BottomLeft - (HasCalibration ? CalibrationBottomLeftOffset : 0.0f);
            float br = rawSensors.BottomRight - (HasCalibration ? CalibrationBottomRightOffset : 0.0f);

            if (tl < 0.0f) tl = 0.0f;
            if (tr < 0.0f) tr = 0.0f;
            if (bl < 0.0f) bl = 0.0f;
            if (br < 0.0f) br = 0.0f;

            float total = tl + tr + bl + br;
            float left = tl + bl;
            float right = tr + br;
            float front = tl + tr;
            float back = bl + br;

            float leftPct = total <= 0.0001f ? 50.0f : (left / total * 100.0f);
            float rightPct = 100.0f - leftPct;
            float frontPct = total <= 0.0001f ? 50.0f : (front / total * 100.0f);
            float backPct = 100.0f - frontPct;
            LastLeftPercent = leftPct;
            LastFrontPercent = frontPct;

            float pressureX = total <= 0.0001f ? 0.0f : ((right - left) / total);
            float pressureY = total <= 0.0001f ? 0.0f : ((front - back) / total);
            LastPressurePoint = new PointF(pressureX, pressureY);

            f.lblTopLeft.Text = "Top left: " + tl.ToString("0.00", CultureInfo.InvariantCulture) + " kg";
            f.lblTopRight.Text = "Top right: " + tr.ToString("0.00", CultureInfo.InvariantCulture) + " kg";
            f.lblBottomLeft.Text = "Bottom left: " + bl.ToString("0.00", CultureInfo.InvariantCulture) + " kg";
            f.lblBottomRight.Text = "Bottom right: " + br.ToString("0.00", CultureInfo.InvariantCulture) + " kg";
            f.lblLeftRight.Text = "Left / Right balance: " + leftPct.ToString("0.0", CultureInfo.InvariantCulture) + "% / " + rightPct.ToString("0.0", CultureInfo.InvariantCulture) + "%";
            f.lblFrontBack.Text = "Front / Back balance: " + frontPct.ToString("0.0", CultureInfo.InvariantCulture) + "% / " + backPct.ToString("0.0", CultureInfo.InvariantCulture) + "%";

            if (DiagnosticsWindow.Count >= 80)
                DiagnosticsWindow.RemoveAt(0);
            DiagnosticsWindow.Add(new SensorSample() { TopLeft = tl, TopRight = tr, BottomLeft = bl, BottomRight = br, Total = total });

            UpdateHardwareCheckPanel();
        }

        static void UpdateWeightVsHeightIndicator()
        {
            ProfileInfo profile = GetSelectedProfile();
            if (profile == null || profile.HeightCm <= 0.0f)
            {
                f.lblWeightVsHeight.Text = "Weight vs height: add profile height to view this indicator.";
                return;
            }

            float heightM = profile.HeightCm / 100.0f;
            if (heightM <= 0.1f)
            {
                f.lblWeightVsHeight.Text = "Weight vs height: height is too low to calculate.";
                return;
            }

            float indicator = LastDisplayedWeightKg / (heightM * heightM);
            string band;
            if (indicator < 18.5f) band = "below the typical range";
            else if (indicator < 25.0f) band = "in the typical range";
            else if (indicator < 30.0f) band = "above the typical range";
            else band = "well above the typical range";

            f.lblWeightVsHeight.Text = "Weight vs height: " + band + " (simple indicator, not a diagnosis).";
        }

        static void UpdateAdviceMessage()
        {
            string advice;
            if (!string.IsNullOrEmpty(LastHardwareSignal))
            {
                advice = "Advice: A sensor corner may need inspection.";
            }
            else if (LastStabilityLevel <= 2)
            {
                advice = "Advice: Stand still a little longer.";
            }
            else if (LastLeftPercent < 45.0f)
            {
                advice = "Advice: Your weight is slightly shifted to the right.";
            }
            else if (LastLeftPercent > 55.0f)
            {
                advice = "Advice: Your weight is slightly shifted to the left.";
            }
            else if (LastFrontPercent < 45.0f)
            {
                advice = "Advice: Your weight is slightly backward.";
            }
            else if (LastFrontPercent > 55.0f)
            {
                advice = "Advice: Your weight is slightly forward.";
            }
            else
            {
                advice = "Advice: Measurement is stable.";
            }
            f.lblAdvice.Text = advice;
        }

        static void MaybeRecordSession()
        {
            DateTime now = DateTime.UtcNow;
            if (LastSessionRecordUtc != DateTime.MinValue && (now - LastSessionRecordUtc).TotalMilliseconds < 250)
                return;

            BalanceBoardSensorsF rawSensors = bb.WiimoteState.BalanceBoardState.SensorValuesKg;
            float tl = rawSensors.TopLeft - (HasCalibration ? CalibrationTopLeftOffset : 0.0f);
            float tr = rawSensors.TopRight - (HasCalibration ? CalibrationTopRightOffset : 0.0f);
            float bl = rawSensors.BottomLeft - (HasCalibration ? CalibrationBottomLeftOffset : 0.0f);
            float br = rawSensors.BottomRight - (HasCalibration ? CalibrationBottomRightOffset : 0.0f);
            if (tl < 0.0f) tl = 0.0f;
            if (tr < 0.0f) tr = 0.0f;
            if (bl < 0.0f) bl = 0.0f;
            if (br < 0.0f) br = 0.0f;

            float total = tl + tr + bl + br;
            float left = tl + bl;
            float front = tl + tr;
            float leftPct = total <= 0.0001f ? 50.0f : (left / total * 100.0f);
            float rightPct = 100.0f - leftPct;
            float frontPct = total <= 0.0001f ? 50.0f : (front / total * 100.0f);
            float backPct = 100.0f - frontPct;
            float pressureX = total <= 0.0001f ? 0.0f : ((tr + br) - (tl + bl)) / total;
            float pressureY = total <= 0.0001f ? 0.0f : ((tl + tr) - (bl + br)) / total;

            ProfileInfo profile = GetSelectedProfile();
            SessionMeasurements.Add(new MeasurementSample()
            {
                TimestampUtc = now,
                ProfileName = (profile == null ? "Default" : profile.Name),
                ProfileHeightCm = (profile == null ? 0.0f : profile.HeightCm),
                TotalWeightKg = LastDisplayedWeightKg,
                SensorTopLeftKg = tl,
                SensorTopRightKg = tr,
                SensorBottomLeftKg = bl,
                SensorBottomRightKg = br,
                LeftPercent = leftPct,
                RightPercent = rightPct,
                FrontPercent = frontPct,
                BackPercent = backPct,
                PressurePointX = pressureX,
                PressurePointY = pressureY,
                StabilityLevel = LastStabilityLevel
            });

            if (SessionMeasurements.Count > 2000)
                SessionMeasurements.RemoveAt(0);

            LastSessionRecordUtc = now;
        }

        static void UpdateHistorySummary()
        {
            int count = SessionMeasurements.Count;
            if (count <= 0)
            {
                f.lblHistorySummary.Text = "No measurements yet.";
                return;
            }

            int begin = Math.Max(0, count - 80);
            float min = SessionMeasurements[begin].TotalWeightKg;
            float max = min;
            float sum = 0.0f;
            for (int i = begin; i < count; i++)
            {
                float w = SessionMeasurements[i].TotalWeightKg;
                if (w < min) min = w;
                if (w > max) max = w;
                sum += w;
            }
            float avg = sum / (count - begin);

            ProfileInfo profile = GetSelectedProfile();
            string profileDesc = (profile == null ? "Default" : profile.Name);
            if (profile != null && profile.HeightCm > 0.0f)
                profileDesc += " (" + profile.HeightCm.ToString("0.0", CultureInfo.InvariantCulture) + " cm)";

            f.lblHistorySummary.Text = "Samples: " + count.ToString() + "   Profile: " + profileDesc +
                "   Recent avg: " + avg.ToString("0.000", CultureInfo.InvariantCulture) + " kg" +
                "   Range: " + min.ToString("0.000", CultureInfo.InvariantCulture) + " .. " + max.ToString("0.000", CultureInfo.InvariantCulture) + " kg";
        }

        static void UpdateHardwareCheckPanel()
        {
            LastHardwareSignal = "";
            if (!f.chkHardwareTest.Checked)
            {
                f.lblDiagnosticsStatus.Text = "Hardware status: check mode is off";
                f.lblDiagnosticsDetails.Text = "Enable check mode to inspect corners and idle drift.";
                return;
            }

            if (DiagnosticsWindow.Count < 10)
            {
                f.lblDiagnosticsStatus.Text = "Hardware status: collecting baseline";
                f.lblDiagnosticsDetails.Text = "Need a little more sensor data.";
                return;
            }

            float totalSum = 0.0f;
            float tlSum = 0.0f, trSum = 0.0f, blSum = 0.0f, brSum = 0.0f;
            float tlMin = float.MaxValue, trMin = float.MaxValue, blMin = float.MaxValue, brMin = float.MaxValue;
            float tlMax = float.MinValue, trMax = float.MinValue, blMax = float.MinValue, brMax = float.MinValue;

            for (int i = 0; i < DiagnosticsWindow.Count; i++)
            {
                SensorSample s = DiagnosticsWindow[i];
                totalSum += s.Total;
                tlSum += s.TopLeft; trSum += s.TopRight; blSum += s.BottomLeft; brSum += s.BottomRight;
                if (s.TopLeft < tlMin) tlMin = s.TopLeft; if (s.TopLeft > tlMax) tlMax = s.TopLeft;
                if (s.TopRight < trMin) trMin = s.TopRight; if (s.TopRight > trMax) trMax = s.TopRight;
                if (s.BottomLeft < blMin) blMin = s.BottomLeft; if (s.BottomLeft > blMax) blMax = s.BottomLeft;
                if (s.BottomRight < brMin) brMin = s.BottomRight; if (s.BottomRight > brMax) brMax = s.BottomRight;
            }

            float sampleCount = DiagnosticsWindow.Count;
            float totalAvg = totalSum / sampleCount;
            float tlAvg = tlSum / sampleCount;
            float trAvg = trSum / sampleCount;
            float blAvg = blSum / sampleCount;
            float brAvg = brSum / sampleCount;
            float allAvg = (tlAvg + trAvg + blAvg + brAvg) / 4.0f;

            List<string> issues = new List<string>();
            if (totalAvg > 10.0f)
            {
                if (tlAvg < allAvg * 0.35f) issues.Add("Top-left may be weak");
                if (trAvg < allAvg * 0.35f) issues.Add("Top-right may be weak");
                if (blAvg < allAvg * 0.35f) issues.Add("Bottom-left may be weak");
                if (brAvg < allAvg * 0.35f) issues.Add("Bottom-right may be weak");
            }
            else
            {
                if ((tlMax - tlMin) > 0.7f) issues.Add("Top-left drift while unloaded");
                if ((trMax - trMin) > 0.7f) issues.Add("Top-right drift while unloaded");
                if ((blMax - blMin) > 0.7f) issues.Add("Bottom-left drift while unloaded");
                if ((brMax - brMin) > 0.7f) issues.Add("Bottom-right drift while unloaded");
            }

            if (issues.Count == 0)
            {
                f.lblDiagnosticsStatus.Text = "Hardware status: looks okay";
                f.lblDiagnosticsDetails.Text = "No obvious corner issue detected by basic checks.";
            }
            else
            {
                LastHardwareSignal = "needs inspection";
                f.lblDiagnosticsStatus.Text = "Hardware status: please inspect";
                string joined = "";
                for (int i = 0; i < issues.Count; i++)
                {
                    if (i > 0) joined += "; ";
                    joined += issues[i];
                }
                f.lblDiagnosticsDetails.Text = joined;
            }
        }

        static void btnReset_Click(object sender, EventArgs e)
        {
            if (HistoryBest <= 0) return;
            float historySum = 0.0f;
            for (int i = 0; i < HistoryBest; i++)
                historySum += History[(HistoryCursor + History.Length - i) % History.Length];
            ZeroedWeight = historySum / HistoryBest;
        }

        static void unitRadioButton_ChangeHandler(object sender, EventArgs e)
        {
            RadioButton radio = sender as RadioButton;
            if (radio == null || !radio.Checked) return;
            if (sender == f.unitSelectorKg) SelectedUnit = EUnit.Kg;
            else if (sender == f.unitSelectorLb) SelectedUnit = EUnit.Lb;
            else if (sender == f.unitSelectorStone) SelectedUnit = EUnit.Stone;
        }

        static void cmbProfiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyProfileToInputs();
            UpdateHistorySummary();
            UpdateWeightVsHeightIndicator();
        }

        static void btnAddProfile_Click(object sender, EventArgs e)
        {
            string name = (f.txtProfileName.Text == null ? "" : f.txtProfileName.Text.Trim());
            if (name.Length == 0)
            {
                MessageBox.Show(f, "Enter a profile name.", "Profile", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (name.IndexOf(',') >= 0)
            {
                MessageBox.Show(f, "Profile name cannot contain commas.", "Profile", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            float heightCm;
            if (!float.TryParse((f.txtProfileHeightCm.Text == null ? "" : f.txtProfileHeightCm.Text.Trim()), NumberStyles.Float, CultureInfo.InvariantCulture, out heightCm))
            {
                MessageBox.Show(f, "Enter a numeric height in centimeters.", "Profile", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (heightCm < 80.0f || heightCm > 250.0f)
            {
                MessageBox.Show(f, "Height should be between 80 and 250 cm.", "Profile", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ProfileInfo existing = FindProfileByName(name);
            if (existing == null)
            {
                Profiles.Add(new ProfileInfo() { Name = name, HeightCm = heightCm });
                f.cmbProfiles.Items.Add(name);
            }
            else
            {
                existing.HeightCm = heightCm;
            }

            SaveProfiles();
            f.cmbProfiles.SelectedItem = name;
            ApplyProfileToInputs();
            UpdateWeightVsHeightIndicator();
        }

        static void btnClearSession_Click(object sender, EventArgs e)
        {
            SessionMeasurements.Clear();
            LastSessionRecordUtc = DateTime.MinValue;
            UpdateHistorySummary();
            f.pnlHistory.Invalidate();
        }

        static void btnExportCsv_Click(object sender, EventArgs e)
        {
            if (SessionMeasurements.Count == 0)
            {
                MessageBox.Show(f, "No session samples to export yet.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "CSV files (*.csv)|*.csv";
            dialog.FileName = BuildExportFileName("csv");
            if (dialog.ShowDialog(f) != DialogResult.OK) return;

            try
            {
                ExportCsv(dialog.FileName);
                MessageBox.Show(f, "CSV export complete.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(f, "CSV export failed.\n\n" + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static void btnExportJson_Click(object sender, EventArgs e)
        {
            if (SessionMeasurements.Count == 0)
            {
                MessageBox.Show(f, "No session samples to export yet.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "JSON files (*.json)|*.json";
            dialog.FileName = BuildExportFileName("json");
            if (dialog.ShowDialog(f) != DialogResult.OK) return;

            try
            {
                ExportJson(dialog.FileName);
                MessageBox.Show(f, "JSON export complete.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(f, "JSON export failed.\n\n" + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static string BuildExportFileName(string extension)
        {
            ProfileInfo p = GetSelectedProfile();
            string safeProfile = (p == null ? "default" : p.Name).Replace(' ', '_');
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            return safeProfile + "_session_" + stamp + "." + extension;
        }

        static string CsvEscape(string value)
        {
            if (value == null) return "";
            if (value.IndexOf('"') >= 0 || value.IndexOf(',') >= 0)
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
        }

        static void ExportCsv(string path)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(CsvHeader);
            for (int i = 0; i < SessionMeasurements.Count; i++)
            {
                MeasurementSample s = SessionMeasurements[i];
                sb.Append(s.TimestampUtc.ToString("o", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(CsvEscape(s.ProfileName)); sb.Append(',');
                sb.Append(s.ProfileHeightCm.ToString("0.0", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.TotalWeightKg.ToString("0.000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.SensorTopLeftKg.ToString("0.000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.SensorTopRightKg.ToString("0.000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.SensorBottomLeftKg.ToString("0.000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.SensorBottomRightKg.ToString("0.000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.LeftPercent.ToString("0.00", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.RightPercent.ToString("0.00", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.FrontPercent.ToString("0.00", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.BackPercent.ToString("0.00", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.PressurePointX.ToString("0.0000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.PressurePointY.ToString("0.0000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(s.StabilityLevel.ToString(CultureInfo.InvariantCulture));
                sb.AppendLine();
            }
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        static string JsonEscape(string text)
        {
            if (text == null) return "";
            return text.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }

        static void ExportJson(string path)
        {
            StringBuilder sb = new StringBuilder();
            ProfileInfo profile = GetSelectedProfile();
            sb.AppendLine("{");
            sb.Append("  \"profile\": \"").Append(JsonEscape(profile == null ? "Default" : profile.Name)).AppendLine("\",");
            sb.Append("  \"profile_height_cm\": ").Append(profile == null ? "0.0" : profile.HeightCm.ToString("0.0", CultureInfo.InvariantCulture)).AppendLine(",");
            sb.Append("  \"exported_at_utc\": \"").Append(DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)).AppendLine("\",");
            sb.AppendLine("  \"samples\": [");
            for (int i = 0; i < SessionMeasurements.Count; i++)
            {
                MeasurementSample s = SessionMeasurements[i];
                sb.AppendLine("    {");
                sb.Append("      \"timestamp_utc\": \"").Append(s.TimestampUtc.ToString("o", CultureInfo.InvariantCulture)).AppendLine("\",");
                sb.Append("      \"profile\": \"").Append(JsonEscape(s.ProfileName)).AppendLine("\",");
                sb.Append("      \"profile_height_cm\": ").Append(s.ProfileHeightCm.ToString("0.0", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"total_weight_kg\": ").Append(s.TotalWeightKg.ToString("0.000", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"sensor_top_left_kg\": ").Append(s.SensorTopLeftKg.ToString("0.000", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"sensor_top_right_kg\": ").Append(s.SensorTopRightKg.ToString("0.000", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"sensor_bottom_left_kg\": ").Append(s.SensorBottomLeftKg.ToString("0.000", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"sensor_bottom_right_kg\": ").Append(s.SensorBottomRightKg.ToString("0.000", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"left_percent\": ").Append(s.LeftPercent.ToString("0.00", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"right_percent\": ").Append(s.RightPercent.ToString("0.00", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"front_percent\": ").Append(s.FrontPercent.ToString("0.00", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"back_percent\": ").Append(s.BackPercent.ToString("0.00", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"pressure_point_x\": ").Append(s.PressurePointX.ToString("0.0000", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"pressure_point_y\": ").Append(s.PressurePointY.ToString("0.0000", CultureInfo.InvariantCulture)).AppendLine(",");
                sb.Append("      \"stability_level\": ").Append(s.StabilityLevel.ToString(CultureInfo.InvariantCulture)).AppendLine();
                sb.Append("    }");
                if (i < SessionMeasurements.Count - 1) sb.Append(',');
                sb.AppendLine();
            }
            sb.AppendLine("  ]");
            sb.AppendLine("}");

            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        static void btnCaptureCalibration_Click(object sender, EventArgs e)
        {
            if (DiagnosticsWindow.Count < 10)
            {
                MessageBox.Show(f, "Need more samples before calibration. Keep board unloaded for a few seconds.", "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            float totalAvg = 0.0f;
            float tl = 0.0f, tr = 0.0f, bl = 0.0f, br = 0.0f;
            for (int i = 0; i < DiagnosticsWindow.Count; i++)
            {
                SensorSample s = DiagnosticsWindow[i];
                totalAvg += s.Total;
                tl += s.TopLeft;
                tr += s.TopRight;
                bl += s.BottomLeft;
                br += s.BottomRight;
            }
            totalAvg /= DiagnosticsWindow.Count;
            if (totalAvg > 5.0f)
            {
                MessageBox.Show(f, "Calibration capture requires an unloaded board. Please step off and try again.", "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            CalibrationTopLeftOffset = tl / DiagnosticsWindow.Count;
            CalibrationTopRightOffset = tr / DiagnosticsWindow.Count;
            CalibrationBottomLeftOffset = bl / DiagnosticsWindow.Count;
            CalibrationBottomRightOffset = br / DiagnosticsWindow.Count;
            HasCalibration = true;
            f.lblCalibrationStatus.Text = "Zero offsets captured and applied to corner load values.";
        }

        static void chkHardwareTest_CheckedChanged(object sender, EventArgs e)
        {
            UpdateHardwareCheckPanel();
            UpdateAdviceMessage();
        }

        static void pnlCenterOfPressure_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rc = f.pnlCenterOfPressure.ClientRectangle;
            e.Graphics.Clear(Color.White);
            using (Pen gridPen = new Pen(Color.LightGray, 1.0f))
            using (Pen axisPen = new Pen(Color.Gray, 1.5f))
            using (Brush dotBrush = new SolidBrush(Color.Red))
            {
                e.Graphics.DrawRectangle(gridPen, 1, 1, rc.Width - 3, rc.Height - 3);
                int cx = rc.Width / 2;
                int cy = rc.Height / 2;
                e.Graphics.DrawLine(axisPen, cx, 0, cx, rc.Height);
                e.Graphics.DrawLine(axisPen, 0, cy, rc.Width, cy);

                float px = cx + LastPressurePoint.X * (rc.Width * 0.45f);
                float py = cy - LastPressurePoint.Y * (rc.Height * 0.45f);
                float r = 6.0f;
                e.Graphics.FillEllipse(dotBrush, px - r, py - r, r * 2.0f, r * 2.0f);
            }
        }

        static void pnlHistory_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rc = f.pnlHistory.ClientRectangle;
            e.Graphics.Clear(Color.White);
            if (SessionMeasurements.Count < 2)
            {
                using (Brush textBrush = new SolidBrush(Color.Gray))
                    e.Graphics.DrawString("Trend appears after more samples.", SystemFonts.DefaultFont, textBrush, 8, 8);
                return;
            }

            int count = SessionMeasurements.Count;
            int start = Math.Max(0, count - rc.Width);
            float min = SessionMeasurements[start].TotalWeightKg;
            float max = min;
            for (int i = start; i < count; i++)
            {
                float w = SessionMeasurements[i].TotalWeightKg;
                if (w < min) min = w;
                if (w > max) max = w;
            }

            float span = (max - min);
            if (span < 0.01f) span = 0.01f;

            using (Pen axisPen = new Pen(Color.LightGray, 1.0f))
            using (Pen linePen = new Pen(Color.DodgerBlue, 1.5f))
            {
                e.Graphics.DrawRectangle(axisPen, 1, 1, rc.Width - 3, rc.Height - 3);
                e.Graphics.DrawLine(axisPen, 0, rc.Height / 2, rc.Width, rc.Height / 2);

                PointF[] points = new PointF[count - start];
                for (int i = start; i < count; i++)
                {
                    float x = (float)(i - start);
                    float n = (SessionMeasurements[i].TotalWeightKg - min) / span;
                    float y = (rc.Height - 4) - (n * (rc.Height - 8)) + 2;
                    points[i - start] = new PointF(x, y);
                }
                if (points.Length >= 2)
                    e.Graphics.DrawLines(linePen, points);
            }
        }
    }
}
