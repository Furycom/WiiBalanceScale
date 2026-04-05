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
[assembly: System.Reflection.AssemblyVersion("1.4.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.4.0.0")]
[assembly: System.Runtime.InteropServices.ComVisible(false)]

namespace WiiBalanceScale
{
    internal class WiiBalanceScale
    {
        enum EConnectionError { None, NoBluetoothAdapter, PermissionDenied, NoDeviceFound, WrongDeviceType, ConnectionFailed }
        enum EUnit { Kg, Lb, Stone }
        enum ESessionQuality { GoodSession, UsableSession, WeakSession, UnstableSession }

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

        class SessionInsight
        {
            public bool HasEnoughSamples;
            public int SampleCount;
            public string ProfileName;
            public float ProfileHeightCm;
            public DateTime StartedUtc;
            public DateTime EndedUtc;
            public float StableWeightKg;
            public float AverageLeftPercent;
            public float AverageFrontPercent;
            public float AverageStabilityLevel;
            public float AveragePressurePointX;
            public float AveragePressurePointY;
            public float TimeToStabilizeSeconds;
            public bool HasStableWindow;
            public float StabilityScore;
            public string StabilityBandText;
            public string StanceTendencyText;
            public string RepeatedPatternText;
            public string StabilityPatternText;
            public bool MatchesUsualPattern;
            public string MainAdvice;
            public string WeightVsHeightText;
            public string LeftRightTendencyText;
            public string FrontBackTendencyText;
            public string StabilityQualityText;
            public ESessionQuality SessionQuality;
            public string SessionQualityText;
            public string SessionQualityReasonText;
            public bool IsStrongEnoughForComparison;
            public string PressurePointTendencyText;
            public string SimilarityText;
            public string SummaryText;
            public string ComparisonText;
            public string TrendText;
            public string ReviewText;
            public List<string> AdviceMessages = new List<string>();
        }

        class SessionHistoryRecord
        {
            public DateTime TimestampUtc;
            public string ProfileName;
            public float ProfileHeightCm;
            public int SampleCount;
            public float StableWeightKg;
            public float AverageLeftPercent;
            public float AverageFrontPercent;
            public float AverageStabilityLevel;
            public float AveragePressurePointX;
            public float AveragePressurePointY;
            public float TimeToStabilizeSeconds;
            public float StabilityScore;
            public string StabilityBandText;
            public string StanceTendencyText;
            public string SessionQualityText;
            public string WeightVsHeightText;
            public string SummaryText;
            public string ReviewText;
            public string ComparisonText;
            public string TrendText;
            public string AdviceText;
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
        static readonly List<ProfileInfo> Profiles = new List<ProfileInfo>();
        static readonly List<SessionHistoryRecord> SessionHistory = new List<SessionHistoryRecord>();
        static readonly Dictionary<string, DateTime> LastSavedSessionEndUtcByProfile = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
        static SessionInsight LastSessionInsight = null;
        static string LastSessionAdviceText = "Advice: Waiting for measurement...";
        static string LastActiveProfileName = "Default";

        static DateTime LastSessionRecordUtc = DateTime.MinValue;
        static int LastStabilityLevel = 0;
        static float LastDisplayedWeightKg = 0.0f;
        static float LastLeftPercent = 50.0f;
        static float LastFrontPercent = 50.0f;
        static System.Drawing.PointF LastPressurePoint = new System.Drawing.PointF(0.0f, 0.0f);

        static readonly string ProfilesPath = Path.Combine(Application.StartupPath, "profiles.csv");
        static readonly string LegacyProfilesPath = Path.Combine(Application.StartupPath, "profiles.txt");
        static readonly string CsvHeader = "timestamp_utc,profile,profile_height_cm,total_weight_kg,sensor_top_left_kg,sensor_top_right_kg,sensor_bottom_left_kg,sensor_bottom_right_kg,left_percent,right_percent,front_percent,back_percent,pressure_point_x,pressure_point_y,stability_level,session_quality,time_to_stabilize_seconds,stability_score,stability_band,stance_tendency,repeated_pattern_if_available,main_advice,session_review_summary,session_summary,session_comparison,recent_trend_interpretation,current_advice,weight_vs_height_text";
        static readonly string SessionHistoryPath = Path.Combine(Application.StartupPath, "session_history.csv");
        static readonly string SessionHistoryHeader = "timestamp_utc,profile,profile_height_cm,sample_count,stable_weight_kg,avg_left_percent,avg_front_percent,avg_stability_level,avg_pressure_x,avg_pressure_y,time_to_stabilize_seconds,stability_score,stability_band,stance_tendency,session_quality,weight_vs_height_text,summary_text,review_text,comparison_text,trend_text,advice_text";

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
            if (CanShowUnicode) f.lblQuality.Font = new Font("Segoe UI Symbol", 40F, FontStyle.Regular, GraphicsUnit.Pixel);

            LoadProfiles();
            LoadSessionHistory();
            BindProfilesToUi();

            f.btnReset.Click += btnReset_Click;
            f.btnAddProfile.Click += btnAddProfile_Click;
            f.cmbProfiles.SelectedIndexChanged += cmbProfiles_SelectedIndexChanged;
            f.btnClearSession.Click += btnClearSession_Click;
            f.btnExportCsv.Click += btnExportCsv_Click;
            f.btnExportJson.Click += btnExportJson_Click;
            f.pnlCenterOfPressure.Paint += pnlCenterOfPressure_Paint;
            f.pnlWeightTrend.Paint += pnlWeightTrend_Paint;
            f.pnlLeftRightTrend.Paint += pnlLeftRightTrend_Paint;
            f.pnlFrontBackTrend.Paint += pnlFrontBackTrend_Paint;
            f.pnlStabilityTrend.Paint += pnlStabilityTrend_Paint;

            EventHandler unitRadioButton_Change = unitRadioButton_ChangeHandler;
            f.unitSelectorKg.CheckedChanged += unitRadioButton_Change;
            f.unitSelectorLb.CheckedChanged += unitRadioButton_Change;
            f.unitSelectorStone.CheckedChanged += unitRadioButton_Change;
            f.unitSelectorKg.Checked = true;
            UpdateWeightVsHeightIndicator();
            UpdateSessionInfo();
            ProfileInfo initialProfile = GetSelectedProfile();
            LastActiveProfileName = (initialProfile == null ? "Default" : initialProfile.Name);

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
            SaveCurrentSessionToHistory();
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

                    string[] parts = SplitCsvLine(line);
                    if (parts.Length == 0) continue;

                    string firstValue = (parts[0] == null ? "" : parts[0].Trim());
                    if (firstValue.Length == 0) continue;
                    if (string.Equals(firstValue, "profile", StringComparison.OrdinalIgnoreCase) || firstValue.StartsWith("#"))
                        continue;
                    if (FindProfileByName(firstValue) != null) continue;

                    float height = 0.0f;
                    if (parts.Length > 1)
                    {
                        string heightRaw = (parts[1] == null ? "" : parts[1].Trim());
                        if (heightRaw.Length > 0)
                            float.TryParse(heightRaw, NumberStyles.Float, CultureInfo.InvariantCulture, out height);
                    }
                    Profiles.Add(new ProfileInfo() { Name = firstValue, HeightCm = Math.Max(0.0f, height) });
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
                    lines[i] = Profiles[i].Name + "," + Profiles[i].HeightCm.ToString("0.0", CultureInfo.InvariantCulture);
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
                if (string.Equals(Profiles[i].Name, name, StringComparison.OrdinalIgnoreCase))
                    return Profiles[i];
            return null;
        }

        static void BindProfilesToUi()
        {
            f.cmbProfiles.Items.Clear();
            for (int i = 0; i < Profiles.Count; i++) f.cmbProfiles.Items.Add(Profiles[i].Name);
            if (f.cmbProfiles.Items.Count > 0) f.cmbProfiles.SelectedIndex = 0;
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

        static string[] SplitCsvLine(string line)
        {
            List<string> values = new List<string>();
            if (line == null) return values.ToArray();

            StringBuilder current = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            current.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
                else
                {
                    if (c == ',')
                    {
                        values.Add(current.ToString());
                        current.Length = 0;
                    }
                    else if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
            }
            values.Add(current.ToString());
            return values.ToArray();
        }

        static void LoadSessionHistory()
        {
            SessionHistory.Clear();
            LastSavedSessionEndUtcByProfile.Clear();
            if (!File.Exists(SessionHistoryPath)) return;

            try
            {
                string[] lines = File.ReadAllLines(SessionHistoryPath);
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = (lines[i] == null ? "" : lines[i].Trim());
                    if (line.Length == 0 || line == SessionHistoryHeader) continue;

                    string[] parts = SplitCsvLine(line);
                    if (parts.Length < 15) continue;

                    DateTime timestampUtc;
                    if (!DateTime.TryParse(parts[0], null, DateTimeStyles.RoundtripKind, out timestampUtc)) continue;

                    SessionHistoryRecord record = new SessionHistoryRecord();
                    record.TimestampUtc = timestampUtc;
                    record.ProfileName = parts[1];
                    float.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out record.ProfileHeightCm);
                    int.TryParse(parts[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out record.SampleCount);
                    float.TryParse(parts[4], NumberStyles.Float, CultureInfo.InvariantCulture, out record.StableWeightKg);
                    float.TryParse(parts[5], NumberStyles.Float, CultureInfo.InvariantCulture, out record.AverageLeftPercent);
                    float.TryParse(parts[6], NumberStyles.Float, CultureInfo.InvariantCulture, out record.AverageFrontPercent);
                    float.TryParse(parts[7], NumberStyles.Float, CultureInfo.InvariantCulture, out record.AverageStabilityLevel);
                    float.TryParse(parts[8], NumberStyles.Float, CultureInfo.InvariantCulture, out record.AveragePressurePointX);
                    float.TryParse(parts[9], NumberStyles.Float, CultureInfo.InvariantCulture, out record.AveragePressurePointY);
                    if (parts.Length >= 21)
                    {
                        float.TryParse(parts[10], NumberStyles.Float, CultureInfo.InvariantCulture, out record.TimeToStabilizeSeconds);
                        float.TryParse(parts[11], NumberStyles.Float, CultureInfo.InvariantCulture, out record.StabilityScore);
                        record.StabilityBandText = parts[12];
                        record.StanceTendencyText = parts[13];
                        record.SessionQualityText = parts[14];
                        record.WeightVsHeightText = parts[15];
                        record.SummaryText = parts[16];
                        record.ReviewText = parts[17];
                        record.ComparisonText = parts[18];
                        record.TrendText = parts[19];
                        record.AdviceText = parts[20];
                    }
                    else if (parts.Length >= 17)
                    {
                        record.TimeToStabilizeSeconds = 0.0f;
                        record.StabilityScore = Math.Max(0.0f, Math.Min(100.0f, ((record.AverageStabilityLevel - 1.0f) / 4.0f) * 100.0f));
                        record.StabilityBandText = GetStabilityBandFromScore(record.StabilityScore);
                        record.StanceTendencyText = ClassifyStanceTendency(record.AverageLeftPercent, record.AverageFrontPercent, 1.5f);
                        record.SessionQualityText = parts[10];
                        record.WeightVsHeightText = parts[11];
                        record.SummaryText = parts[12];
                        record.ReviewText = parts[13];
                        record.ComparisonText = parts[14];
                        record.TrendText = parts[15];
                        record.AdviceText = parts[16];
                    }
                    else
                    {
                        record.TimeToStabilizeSeconds = 0.0f;
                        record.StabilityScore = Math.Max(0.0f, Math.Min(100.0f, ((record.AverageStabilityLevel - 1.0f) / 4.0f) * 100.0f));
                        record.StabilityBandText = GetStabilityBandFromScore(record.StabilityScore);
                        record.StanceTendencyText = ClassifyStanceTendency(record.AverageLeftPercent, record.AverageFrontPercent, 1.5f);
                        record.SessionQualityText = "usable session";
                        record.WeightVsHeightText = parts[10];
                        record.SummaryText = parts[11];
                        record.ReviewText = parts[11];
                        record.ComparisonText = parts[12];
                        record.TrendText = parts[13];
                        record.AdviceText = parts[14];
                    }
                    SessionHistory.Add(record);
                    DateTime previousLastSavedUtc;
                    if (!LastSavedSessionEndUtcByProfile.TryGetValue(record.ProfileName, out previousLastSavedUtc) || record.TimestampUtc > previousLastSavedUtc)
                        LastSavedSessionEndUtcByProfile[record.ProfileName] = record.TimestampUtc;
                }
            }
            catch
            {
                SessionHistory.Clear();
            }
        }

        static void SaveSessionHistory()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(SessionHistoryHeader);
            for (int i = 0; i < SessionHistory.Count; i++)
            {
                SessionHistoryRecord r = SessionHistory[i];
                sb.Append(r.TimestampUtc.ToString("o", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(CsvEscape(r.ProfileName)); sb.Append(',');
                sb.Append(Math.Max(0.0f, r.ProfileHeightCm).ToString("0.0", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(Math.Max(0, r.SampleCount).ToString(CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(r.StableWeightKg.ToString("0.000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(r.AverageLeftPercent.ToString("0.00", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(r.AverageFrontPercent.ToString("0.00", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(r.AverageStabilityLevel.ToString("0.00", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(r.AveragePressurePointX.ToString("0.0000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(r.AveragePressurePointY.ToString("0.0000", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(Math.Max(0.0f, r.TimeToStabilizeSeconds).ToString("0.0", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(Math.Max(0.0f, Math.Min(100.0f, r.StabilityScore)).ToString("0.0", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(CsvEscape(r.StabilityBandText)); sb.Append(',');
                sb.Append(CsvEscape(r.StanceTendencyText)); sb.Append(',');
                sb.Append(CsvEscape(r.SessionQualityText)); sb.Append(',');
                sb.Append(CsvEscape(r.WeightVsHeightText)); sb.Append(',');
                sb.Append(CsvEscape(r.SummaryText)); sb.Append(',');
                sb.Append(CsvEscape(r.ReviewText)); sb.Append(',');
                sb.Append(CsvEscape(r.ComparisonText)); sb.Append(',');
                sb.Append(CsvEscape(r.TrendText)); sb.Append(',');
                sb.Append(CsvEscape(r.AdviceText));
                sb.AppendLine();
            }

            string tempPath = SessionHistoryPath + ".tmp";
            File.WriteAllText(tempPath, sb.ToString(), Encoding.UTF8);
            if (File.Exists(SessionHistoryPath))
                File.Delete(SessionHistoryPath);
            File.Move(tempPath, SessionHistoryPath);
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

            LastSessionRecordUtc = DateTime.MinValue;
        }

        static void BoardTimer_Tick(object sender, EventArgs e)
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

            if (bb == null)
            {
                ConnectBalanceBoard(false);
                return;
            }

            float kg = bb.WiimoteState.BalanceBoardState.WeightKg, historySum = 0.0f, maxHist = kg, minHist = kg, maxDiff = 0.0f;
            if (kg < -200)
            {
                ConnectBalanceBoard(false);
                return;
            }

            HistoryCursor++;
            History[HistoryCursor % History.Length] = kg;
            for (HistoryBest = 0; HistoryBest < History.Length; HistoryBest++)
            {
                float historyEntry = History[(HistoryCursor + History.Length - HistoryBest) % History.Length];
                if (Math.Abs(maxHist - historyEntry) > 1.0f) break;
                if (Math.Abs(minHist - historyEntry) > 1.0f) break;
                if (historyEntry > maxHist) maxHist = historyEntry;
                if (historyEntry < minHist) minHist = historyEntry;
                float diff = Math.Max(Math.Abs(historyEntry - kg), Math.Abs((historySum + historyEntry) / (HistoryBest + 1) - kg));
                if (diff > maxDiff) maxDiff = diff;
                if (diff > 1.0f) break;
                historySum += historyEntry;
            }

            if (HistoryBest <= 0) return;
            kg = historySum / HistoryBest - ZeroedWeight;
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

            UpdateLiveMetrics();
            UpdateWeightVsHeightIndicator();
            UpdateAdviceMessage();
            MaybeRecordSession();
            UpdateSessionInfo();
            f.pnlCenterOfPressure.Invalidate();
        }

        static void UpdateLiveMetrics()
        {
            BalanceBoardSensorsF sensors = bb.WiimoteState.BalanceBoardState.SensorValuesKg;
            float tl = Math.Max(0.0f, sensors.TopLeft);
            float tr = Math.Max(0.0f, sensors.TopRight);
            float bl = Math.Max(0.0f, sensors.BottomLeft);
            float br = Math.Max(0.0f, sensors.BottomRight);

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

            LastPressurePoint = new System.Drawing.PointF(
                total <= 0.0001f ? 0.0f : ((right - left) / total),
                total <= 0.0001f ? 0.0f : ((front - back) / total));

            f.lblTopLeft.Text = "Top left: " + tl.ToString("0.00", CultureInfo.InvariantCulture) + " kg";
            f.lblTopRight.Text = "Top right: " + tr.ToString("0.00", CultureInfo.InvariantCulture) + " kg";
            f.lblBottomLeft.Text = "Bottom left: " + bl.ToString("0.00", CultureInfo.InvariantCulture) + " kg";
            f.lblBottomRight.Text = "Bottom right: " + br.ToString("0.00", CultureInfo.InvariantCulture) + " kg";
            f.lblLeftRight.Text = "Left / Right balance: " + leftPct.ToString("0.0", CultureInfo.InvariantCulture) + "% / " + rightPct.ToString("0.0", CultureInfo.InvariantCulture) + "%";
            f.lblFrontBack.Text = "Front / Back balance: " + frontPct.ToString("0.0", CultureInfo.InvariantCulture) + "% / " + backPct.ToString("0.0", CultureInfo.InvariantCulture) + "%";
        }

        static void UpdateWeightVsHeightIndicator()
        {
            ProfileInfo profile = GetSelectedProfile();
            f.lblWeightVsHeight.Text = BuildWeightVsHeightText(profile, LastDisplayedWeightKg);
        }

        static string BuildWeightVsHeightText(ProfileInfo profile, float weightKg)
        {
            if (profile == null || profile.HeightCm <= 0.0f)
                return "Weight vs height: add profile height to view this indicator.";

            float heightM = profile.HeightCm / 100.0f;
            if (heightM <= 0.1f)
                return "Weight vs height: height is too low to calculate.";

            float indicator = weightKg / (heightM * heightM);
            string band;
            if (indicator < 18.5f) band = "below the typical range";
            else if (indicator < 25.0f) band = "in the typical range";
            else if (indicator < 30.0f) band = "above the typical range";
            else band = "well above the typical range";

            return "Weight vs height: " + band + " (simple indicator, not a diagnosis).";
        }

        static SessionInsight BuildSessionInsight(ProfileInfo profile)
        {
            SessionInsight insight = new SessionInsight();
            insight.ProfileName = (profile == null ? "Default" : profile.Name);
            insight.ProfileHeightCm = (profile == null ? 0.0f : profile.HeightCm);

            List<MeasurementSample> samples = new List<MeasurementSample>();
            for (int i = 0; i < SessionMeasurements.Count; i++)
                if (string.Equals(SessionMeasurements[i].ProfileName, insight.ProfileName, StringComparison.OrdinalIgnoreCase))
                    samples.Add(SessionMeasurements[i]);

            insight.SampleCount = samples.Count;
            insight.HasEnoughSamples = samples.Count >= 8;
            if (samples.Count == 0)
            {
                insight.SummaryText = "Session summary: waiting for enough samples...";
                insight.ComparisonText = "Compared with your previous session: not available yet.";
                insight.TrendText = "Recent trend: not enough history yet.";
                insight.ReviewText = "Session review: waiting for enough samples...";
                insight.RepeatedPatternText = "Repeated pattern: not enough strong history yet.";
                insight.StabilityPatternText = "stability pattern not ready yet";
                insight.MainAdvice = "Stay still for a few seconds so the board can settle.";
                return insight;
            }

            insight.StartedUtc = samples[0].TimestampUtc;
            insight.EndedUtc = samples[samples.Count - 1].TimestampUtc;

            float sumWeight = 0.0f, sumLeft = 0.0f, sumFront = 0.0f, sumStability = 0.0f, sumX = 0.0f, sumY = 0.0f;
            int stableCount = 0;
            float stableWeightSum = 0.0f;
            int leftHeavyCount = 0, rightHeavyCount = 0, frontHeavyCount = 0, backHeavyCount = 0;
            int driftCount = 0;
            float minWeight = float.MaxValue;
            float maxWeight = float.MinValue;
            float previousLeft = samples[0].LeftPercent;
            float previousFront = samples[0].FrontPercent;
            float copPath = 0.0f;
            float sumLeftDeltaAbs = 0.0f;
            float sumFrontDeltaAbs = 0.0f;
            float sumLeftVariance = 0.0f;
            float sumFrontVariance = 0.0f;

            for (int i = 0; i < samples.Count; i++)
            {
                MeasurementSample s = samples[i];
                sumWeight += s.TotalWeightKg;
                sumLeft += s.LeftPercent;
                sumFront += s.FrontPercent;
                sumStability += s.StabilityLevel;
                sumX += s.PressurePointX;
                sumY += s.PressurePointY;
                if (s.TotalWeightKg < minWeight) minWeight = s.TotalWeightKg;
                if (s.TotalWeightKg > maxWeight) maxWeight = s.TotalWeightKg;
                if (s.StabilityLevel >= 3)
                {
                    stableCount++;
                    stableWeightSum += s.TotalWeightKg;
                }
                if (s.LeftPercent > 52.0f) leftHeavyCount++;
                else if (s.LeftPercent < 48.0f) rightHeavyCount++;
                if (s.FrontPercent > 52.0f) frontHeavyCount++;
                else if (s.FrontPercent < 48.0f) backHeavyCount++;
                if (i > 0)
                {
                    if (Math.Abs(s.LeftPercent - previousLeft) >= 4.0f || Math.Abs(s.FrontPercent - previousFront) >= 4.0f)
                        driftCount++;
                    sumLeftDeltaAbs += Math.Abs(s.LeftPercent - previousLeft);
                    sumFrontDeltaAbs += Math.Abs(s.FrontPercent - previousFront);
                    float dx = s.PressurePointX - samples[i - 1].PressurePointX;
                    float dy = s.PressurePointY - samples[i - 1].PressurePointY;
                    copPath += (float)Math.Sqrt(dx * dx + dy * dy);
                    previousLeft = s.LeftPercent;
                    previousFront = s.FrontPercent;
                }
            }

            insight.StableWeightKg = (stableCount >= 4 ? stableWeightSum / stableCount : sumWeight / samples.Count);
            insight.AverageLeftPercent = sumLeft / samples.Count;
            insight.AverageFrontPercent = sumFront / samples.Count;
            insight.AverageStabilityLevel = sumStability / samples.Count;
            insight.AveragePressurePointX = sumX / samples.Count;
            insight.AveragePressurePointY = sumY / samples.Count;
            insight.WeightVsHeightText = BuildWeightVsHeightText(profile, insight.StableWeightKg);
            insight.TimeToStabilizeSeconds = EstimateTimeToStabilizeSeconds(samples, out insight.HasStableWindow);

            for (int i = 0; i < samples.Count; i++)
            {
                float leftDelta = samples[i].LeftPercent - insight.AverageLeftPercent;
                float frontDelta = samples[i].FrontPercent - insight.AverageFrontPercent;
                sumLeftVariance += leftDelta * leftDelta;
                sumFrontVariance += frontDelta * frontDelta;
            }
            float leftStdDev = (float)Math.Sqrt(sumLeftVariance / samples.Count);
            float frontStdDev = (float)Math.Sqrt(sumFrontVariance / samples.Count);
            bool variableStance = leftStdDev >= 2.6f || frontStdDev >= 2.6f;
            insight.StanceTendencyText = (variableStance ? "variable stance" : ClassifyStanceTendency(insight.AverageLeftPercent, insight.AverageFrontPercent, 1.4f));

            if (leftHeavyCount > rightHeavyCount + 3) insight.LeftRightTendencyText = "slightly more weight on the left";
            else if (rightHeavyCount > leftHeavyCount + 3) insight.LeftRightTendencyText = "slightly more weight on the right";
            else insight.LeftRightTendencyText = "left and right mostly even";

            if (frontHeavyCount > backHeavyCount + 3) insight.FrontBackTendencyText = "slightly more weight forward";
            else if (backHeavyCount > frontHeavyCount + 3) insight.FrontBackTendencyText = "slightly more weight backward";
            else insight.FrontBackTendencyText = "front and back mostly even";

            insight.StabilityScore = BuildStabilityScore(insight.AverageStabilityLevel, copPath, driftCount, Math.Max(0.0f, maxWeight - minWeight), samples.Count, sumLeftDeltaAbs, sumFrontDeltaAbs);
            insight.StabilityBandText = GetStabilityBandFromScore(insight.StabilityScore);
            insight.StabilityQualityText = insight.StabilityBandText;

            if (Math.Abs(insight.AveragePressurePointX) < 0.08f && Math.Abs(insight.AveragePressurePointY) < 0.08f) insight.PressurePointTendencyText = "pressure point stayed near center";
            else if (Math.Abs(insight.AveragePressurePointX) > Math.Abs(insight.AveragePressurePointY)) insight.PressurePointTendencyText = (insight.AveragePressurePointX > 0 ? "pressure point tended to the right" : "pressure point tended to the left");
            else insight.PressurePointTendencyText = (insight.AveragePressurePointY > 0 ? "pressure point tended forward" : "pressure point tended backward");

            float weightRange = Math.Max(0.0f, maxWeight - minWeight);
            insight.SessionQuality = ClassifySessionQuality(samples.Count, insight.AverageStabilityLevel, weightRange, driftCount);
            insight.SessionQualityText = GetSessionQualityText(insight.SessionQuality);
            insight.SessionQualityReasonText = BuildSessionQualityReason(samples.Count, insight.AverageStabilityLevel, weightRange, driftCount);
            insight.IsStrongEnoughForComparison = (insight.SessionQuality == ESessionQuality.GoodSession || insight.SessionQuality == ESessionQuality.UsableSession);

            BuildComparisonAndAdvice(insight);
            insight.SummaryText = "Session summary: " + insight.StableWeightKg.ToString("0.0", CultureInfo.InvariantCulture) + " kg, " + insight.SessionQualityText + ", " + insight.StabilityBandText + ", stabilized in " + insight.TimeToStabilizeSeconds.ToString("0.0", CultureInfo.InvariantCulture) + " s.";
            insight.ReviewText = "Session review: final stable weight " + insight.StableWeightKg.ToString("0.0", CultureInfo.InvariantCulture) + " kg, " + insight.SessionQualityText + ", time to stabilize " + insight.TimeToStabilizeSeconds.ToString("0.0", CultureInfo.InvariantCulture) + " s, stability " + insight.StabilityBandText + " (" + insight.StabilityScore.ToString("0", CultureInfo.InvariantCulture) + "/100), stance " + insight.StanceTendencyText + ", pattern check: " + GetPatternCheckText(insight) + ".";
            return insight;
        }

        static float EstimateTimeToStabilizeSeconds(List<MeasurementSample> samples, out bool hasStableWindow)
        {
            hasStableWindow = false;
            if (samples == null || samples.Count < 2) return 0.0f;
            int window = 4;
            if (samples.Count < window)
                return (float)Math.Max(0.0, (samples[samples.Count - 1].TimestampUtc - samples[0].TimestampUtc).TotalSeconds);

            for (int i = 0; i <= samples.Count - window; i++)
            {
                bool stable = true;
                float minWeight = float.MaxValue, maxWeight = float.MinValue;
                float minLeft = float.MaxValue, maxLeft = float.MinValue;
                float minFront = float.MaxValue, maxFront = float.MinValue;
                for (int j = 0; j < window; j++)
                {
                    MeasurementSample s = samples[i + j];
                    if (s.StabilityLevel < 3) { stable = false; break; }
                    if (s.TotalWeightKg < minWeight) minWeight = s.TotalWeightKg;
                    if (s.TotalWeightKg > maxWeight) maxWeight = s.TotalWeightKg;
                    if (s.LeftPercent < minLeft) minLeft = s.LeftPercent;
                    if (s.LeftPercent > maxLeft) maxLeft = s.LeftPercent;
                    if (s.FrontPercent < minFront) minFront = s.FrontPercent;
                    if (s.FrontPercent > maxFront) maxFront = s.FrontPercent;
                }
                if (!stable) continue;
                if ((maxWeight - minWeight) > 0.35f) continue;
                if ((maxLeft - minLeft) > 2.8f || (maxFront - minFront) > 2.8f) continue;
                hasStableWindow = true;
                return (float)Math.Max(0.0, (samples[i + window - 1].TimestampUtc - samples[0].TimestampUtc).TotalSeconds);
            }
            return (float)Math.Max(0.0, (samples[samples.Count - 1].TimestampUtc - samples[0].TimestampUtc).TotalSeconds);
        }

        static float BuildStabilityScore(float avgStabilityLevel, float copPath, int driftCount, float weightRangeKg, int sampleCount, float sumLeftDeltaAbs, float sumFrontDeltaAbs)
        {
            if (sampleCount <= 1) return 0.0f;
            float stabilityNorm = Math.Max(0.0f, Math.Min(1.0f, (avgStabilityLevel - 1.0f) / 4.0f));
            float avgCopStep = copPath / (sampleCount - 1);
            float copMovePenalty = Math.Max(0.0f, Math.Min(1.0f, (avgCopStep - 0.012f) / 0.06f));
            float driftPenalty = Math.Max(0.0f, Math.Min(1.0f, (float)driftCount / Math.Max(1, sampleCount - 1) / 0.6f));
            float weightPenalty = Math.Max(0.0f, Math.Min(1.0f, (weightRangeKg - 0.18f) / 1.6f));
            float avgBalanceShift = (sumLeftDeltaAbs + sumFrontDeltaAbs) / (sampleCount - 1);
            float shiftPenalty = Math.Max(0.0f, Math.Min(1.0f, (avgBalanceShift - 1.8f) / 4.5f));
            float score = 100.0f * (0.34f * stabilityNorm + 0.23f * (1.0f - copMovePenalty) + 0.18f * (1.0f - driftPenalty) + 0.17f * (1.0f - weightPenalty) + 0.08f * (1.0f - shiftPenalty));
            return Math.Max(0.0f, Math.Min(100.0f, score));
        }

        static string GetStabilityBandFromScore(float score)
        {
            if (score >= 82.0f) return "very steady";
            if (score >= 65.0f) return "fairly steady";
            if (score >= 45.0f) return "somewhat unsteady";
            return "very unsteady";
        }

        static string ClassifyStanceTendency(float avgLeftPercent, float avgFrontPercent, float centeredThreshold)
        {
            float leftDelta = avgLeftPercent - 50.0f;
            float frontDelta = avgFrontPercent - 50.0f;
            if (Math.Abs(leftDelta) <= centeredThreshold && Math.Abs(frontDelta) <= centeredThreshold)
                return "centered stance";
            if (Math.Abs(leftDelta) >= Math.Abs(frontDelta))
                return (leftDelta > 0.0f ? "slight left compensation" : "slight right compensation");
            return (frontDelta > 0.0f ? "slightly forward stance" : "slightly backward stance");
        }

        static ESessionQuality ClassifySessionQuality(int sampleCount, float avgStabilityLevel, float weightRangeKg, int driftCount)
        {
            if (sampleCount < 6 || avgStabilityLevel < 2.0f) return ESessionQuality.UnstableSession;
            if (sampleCount < 10 || avgStabilityLevel < 2.8f || weightRangeKg > 1.8f || driftCount >= 8) return ESessionQuality.WeakSession;
            if (sampleCount < 16 || avgStabilityLevel < 3.5f || weightRangeKg > 1.0f || driftCount >= 5) return ESessionQuality.UsableSession;
            return ESessionQuality.GoodSession;
        }

        static string GetSessionQualityText(ESessionQuality quality)
        {
            if (quality == ESessionQuality.GoodSession) return "good session";
            if (quality == ESessionQuality.UsableSession) return "usable session";
            if (quality == ESessionQuality.WeakSession) return "weak session";
            return "unstable session";
        }

        static string BuildSessionQualityReason(int sampleCount, float avgStabilityLevel, float weightRangeKg, int driftCount)
        {
            if (sampleCount < 6) return "very short capture with too few samples";
            if (avgStabilityLevel < 2.0f) return "high movement made the reading unstable";
            if (weightRangeKg > 1.8f) return "weight changed too much during this reading";
            if (driftCount >= 8) return "balance drift repeated often";
            if (sampleCount < 16 || avgStabilityLevel < 3.5f) return "usable but not as steady as your stronger sessions";
            return "steady sample count, stable stance, and consistent readings";
        }

        static void BuildComparisonAndAdvice(SessionInsight insight)
        {
            List<SessionHistoryRecord> profileHistory = new List<SessionHistoryRecord>();
            for (int i = SessionHistory.Count - 1; i >= 0; i--)
                if (string.Equals(SessionHistory[i].ProfileName, insight.ProfileName, StringComparison.OrdinalIgnoreCase))
                    profileHistory.Add(SessionHistory[i]);

            SessionHistoryRecord previous = null;
            for (int i = 0; i < profileHistory.Count; i++)
            {
                if (IsRecordComparable(profileHistory[i]))
                {
                    previous = profileHistory[i];
                    break;
                }
            }

            List<SessionHistoryRecord> recentComparable = new List<SessionHistoryRecord>();
            for (int i = 0; i < profileHistory.Count && recentComparable.Count < 5; i++)
                if (IsRecordComparable(profileHistory[i]))
                    recentComparable.Add(profileHistory[i]);
            int recentCount = recentComparable.Count;

            float avgWeight = 0.0f, avgLeft = 50.0f, avgFront = 50.0f, avgStability = 3.0f;
            if (recentCount > 0)
            {
                avgLeft = 0.0f; avgFront = 0.0f; avgStability = 0.0f;
                for (int i = 0; i < recentCount; i++)
                {
                    avgWeight += recentComparable[i].StableWeightKg;
                    avgLeft += recentComparable[i].AverageLeftPercent;
                    avgFront += recentComparable[i].AverageFrontPercent;
                    avgStability += recentComparable[i].AverageStabilityLevel;
                }
                avgWeight /= recentCount;
                avgLeft /= recentCount;
                avgFront /= recentCount;
                avgStability /= recentCount;
            }

            insight.RepeatedPatternText = BuildRepeatedPatternText(profileHistory);
            insight.StabilityPatternText = BuildStabilityPatternText(profileHistory);
            insight.MatchesUsualPattern = MatchesRepeatedPattern(insight, insight.RepeatedPatternText);

            if (!insight.IsStrongEnoughForComparison)
            {
                insight.SimilarityText = "session quality is too weak for reliable comparison";
                insight.ComparisonText = "Compared with previous: this session is marked " + insight.SessionQualityText + " and is not used for detailed comparison.";
                insight.TrendText = BuildTrendInterpretation(insight, recentComparable, avgWeight, avgLeft, avgFront, avgStability);
            }
            else if (previous == null)
            {
                insight.SimilarityText = "first recorded session for this profile";
                insight.ComparisonText = "Compared with your previous session: not available yet.";
                insight.TrendText = "Recent trend: not enough history yet.";
            }
            else
            {
                float weightDeltaPrev = insight.StableWeightKg - previous.StableWeightKg;
                float leftDeltaPrev = insight.AverageLeftPercent - previous.AverageLeftPercent;
                float frontDeltaPrev = insight.AverageFrontPercent - previous.AverageFrontPercent;
                float stabilityDeltaPrev = insight.AverageStabilityLevel - previous.AverageStabilityLevel;
                bool similar = Math.Abs(weightDeltaPrev) < 1.0f && Math.Abs(leftDeltaPrev) < 2.5f && Math.Abs(frontDeltaPrev) < 2.5f && Math.Abs(stabilityDeltaPrev) < 0.5f;
                insight.SimilarityText = (similar ? "looks similar to your previous session" : "looks different from your previous session");
                insight.ComparisonText = "Compared with previous: weight " + FormatDelta(weightDeltaPrev, "kg") + ", left/right " + FormatDelta(leftDeltaPrev, "pts") + ", front/back " + FormatDelta(frontDeltaPrev, "pts") + ", stability " + FormatDelta(stabilityDeltaPrev, "lvl") + ".";
                insight.TrendText = BuildTrendInterpretation(insight, recentComparable, avgWeight, avgLeft, avgFront, avgStability);
            }

            insight.AdviceMessages.Clear();
            if (insight.TimeToStabilizeSeconds > 4.0f) insight.AdviceMessages.Add("Settle your feet first, then stay still for a cleaner reading.");
            if (insight.StabilityScore < 45.0f) insight.AdviceMessages.Add("Try to reduce sway for a few seconds before finishing the reading.");
            else if (insight.StabilityScore < 65.0f) insight.AdviceMessages.Add("A little less movement would make this reading stronger.");
            if (!string.Equals(insight.StanceTendencyText, "centered stance", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(insight.StanceTendencyText, "variable stance", StringComparison.OrdinalIgnoreCase))
                insight.AdviceMessages.Add("Your stance showed " + insight.StanceTendencyText.Replace("slight", "a slight") + " today.");
            if (string.Equals(insight.StanceTendencyText, "variable stance", StringComparison.OrdinalIgnoreCase))
                insight.AdviceMessages.Add("Your stance moved around a lot during this session.");
            if (previous != null)
            {
                if (Math.Abs(insight.StableWeightKg - previous.StableWeightKg) >= 1.5f)
                    insight.AdviceMessages.Add("Your weight changed compared with your previous session.");
                if (Math.Abs(insight.AverageLeftPercent - previous.AverageLeftPercent) >= 3.0f || Math.Abs(insight.AverageFrontPercent - previous.AverageFrontPercent) >= 3.0f)
                    insight.AdviceMessages.Add("Your balance pattern changed compared with your previous session.");
            }
            if (recentCount >= 3 && !insight.MatchesUsualPattern && IsRepeatedPatternStrong(insight.RepeatedPatternText))
                insight.AdviceMessages.Add("This session is different from your usual pattern.");
            if (IsRepeatedPatternStrong(insight.RepeatedPatternText))
                insight.AdviceMessages.Add("Your repeated pattern is " + insight.RepeatedPatternText.Replace("Repeated pattern: ", "").Replace(".", "") + ".");
            if (insight.AdviceMessages.Count == 0)
                insight.AdviceMessages.Add("Measurement looked steady and balanced overall.");
            if (!insight.IsStrongEnoughForComparison)
                insight.AdviceMessages.Add("Try one longer and steadier reading so your next comparison is stronger.");
            insight.MainAdvice = insight.AdviceMessages[0];
        }

        static string BuildRepeatedPatternText(List<SessionHistoryRecord> profileHistoryNewestFirst)
        {
            List<SessionHistoryRecord> strong = new List<SessionHistoryRecord>();
            for (int i = 0; i < profileHistoryNewestFirst.Count && strong.Count < 10; i++)
                if (IsRecordComparable(profileHistoryNewestFirst[i]))
                    strong.Add(profileHistoryNewestFirst[i]);
            if (strong.Count < 4)
                return "Repeated pattern: not enough strong history yet.";

            int leftCount = 0, rightCount = 0, forwardCount = 0, backwardCount = 0, centeredCount = 0;
            for (int i = 0; i < strong.Count; i++)
            {
                string stance = ClassifyStanceTendency(strong[i].AverageLeftPercent, strong[i].AverageFrontPercent, 1.4f);
                if (stance == "slight left compensation") leftCount++;
                else if (stance == "slight right compensation") rightCount++;
                else if (stance == "slightly forward stance") forwardCount++;
                else if (stance == "slightly backward stance") backwardCount++;
                else centeredCount++;
            }
            int threshold = (int)Math.Ceiling(strong.Count * 0.7f);
            if (leftCount >= threshold) return "Repeated pattern: often left-leaning.";
            if (rightCount >= threshold) return "Repeated pattern: often right-leaning.";
            if (forwardCount >= threshold) return "Repeated pattern: often forward.";
            if (backwardCount >= threshold) return "Repeated pattern: often backward.";
            if (centeredCount >= threshold) return "Repeated pattern: usually centered.";
            return "Repeated pattern: mixed, no strong direction yet.";
        }

        static string BuildStabilityPatternText(List<SessionHistoryRecord> profileHistoryNewestFirst)
        {
            List<SessionHistoryRecord> recent = new List<SessionHistoryRecord>();
            for (int i = 0; i < profileHistoryNewestFirst.Count && recent.Count < 8; i++)
                recent.Add(profileHistoryNewestFirst[i]);
            if (recent.Count < 4) return "stability pattern not ready yet";

            int good = 0, mixed = 0, weak = 0;
            for (int i = 0; i < recent.Count; i++)
            {
                string quality = (recent[i].SessionQualityText == null ? "" : recent[i].SessionQualityText);
                if (string.Equals(quality, "good session", StringComparison.OrdinalIgnoreCase)) good++;
                else if (string.Equals(quality, "usable session", StringComparison.OrdinalIgnoreCase)) mixed++;
                else weak++;
            }
            if (good >= (int)Math.Ceiling(recent.Count * 0.65f)) return "stability usually good";
            if (weak >= (int)Math.Ceiling(recent.Count * 0.45f)) return "stability usually weak";
            return "stability usually mixed";
        }

        static bool MatchesRepeatedPattern(SessionInsight insight, string repeatedPatternText)
        {
            if (insight == null || repeatedPatternText == null) return false;
            if (repeatedPatternText.Contains("not enough") || repeatedPatternText.Contains("mixed"))
                return false;
            if (repeatedPatternText.Contains("left") && insight.StanceTendencyText.Contains("left")) return true;
            if (repeatedPatternText.Contains("right") && insight.StanceTendencyText.Contains("right")) return true;
            if (repeatedPatternText.Contains("forward") && insight.StanceTendencyText.Contains("forward")) return true;
            if (repeatedPatternText.Contains("backward") && insight.StanceTendencyText.Contains("backward")) return true;
            if (repeatedPatternText.Contains("centered") && insight.StanceTendencyText.Contains("centered")) return true;
            return false;
        }

        static bool IsRepeatedPatternStrong(string repeatedPatternText)
        {
            if (string.IsNullOrEmpty(repeatedPatternText)) return false;
            return !repeatedPatternText.Contains("not enough") && !repeatedPatternText.Contains("mixed");
        }

        static string GetPatternCheckText(SessionInsight insight)
        {
            if (insight == null || !IsRepeatedPatternStrong(insight.RepeatedPatternText))
                return "usual pattern not clear yet";
            return (insight.MatchesUsualPattern ? "matches your usual pattern" : "different from your usual pattern");
        }

        static bool IsRecordComparable(SessionHistoryRecord record)
        {
            if (record == null) return false;
            return string.Equals(record.SessionQualityText, "good session", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(record.SessionQualityText, "usable session", StringComparison.OrdinalIgnoreCase);
        }

        static string BuildTrendInterpretation(SessionInsight insight, List<SessionHistoryRecord> recentComparable, float avgWeight, float avgLeft, float avgFront, float avgStability)
        {
            if (recentComparable.Count == 0)
                return "Recent trend: not enough history yet.";

            List<string> highlights = new List<string>();
            float weightDelta = insight.StableWeightKg - avgWeight;
            if (Math.Abs(weightDelta) < 0.6f) highlights.Add("weight has been stable recently");
            else if (weightDelta > 0.0f) highlights.Add("weight has slowly increased");
            else highlights.Add("weight has slowly decreased");

            float balanceSpread = 0.0f;
            for (int i = 0; i < recentComparable.Count; i++)
            {
                float dist = Math.Abs(recentComparable[i].AverageLeftPercent - 50.0f);
                if (dist > balanceSpread) balanceSpread = dist;
            }
            if (Math.Abs(insight.AverageLeftPercent - 50.0f) < 1.7f && Math.Abs(insight.AverageFrontPercent - 50.0f) < 1.7f)
                highlights.Add("balance is usually centered");
            else if (balanceSpread >= 3.5f)
                highlights.Add("left/right balance varies from session to session");

            if (insight.AverageStabilityLevel < avgStability - 0.5f)
                highlights.Add("this session was less stable than usual");
            else if (insight.AverageStabilityLevel > avgStability + 0.5f)
                highlights.Add("this session was steadier than usual");

            if (highlights.Count == 0)
                highlights.Add("no clear direction yet");
            return "Recent trend: " + string.Join("; ", highlights.ToArray()) + ".";
        }

        static string FormatDelta(float value, string unit)
        {
            string direction = (Math.Abs(value) < 0.0001f ? "no change" : (value > 0.0f ? "up " : "down "));
            if (direction == "no change") return "no change";
            return direction + Math.Abs(value).ToString("0.0", CultureInfo.InvariantCulture) + " " + unit;
        }

        static string TruncateForLabel(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength) return text;
            if (maxLength <= 3) return text.Substring(0, Math.Max(0, maxLength));
            return text.Substring(0, maxLength - 3) + "...";
        }

        static int CountSamplesForProfile(string profileName)
        {
            if (string.IsNullOrEmpty(profileName)) return 0;
            int count = 0;
            for (int i = 0; i < SessionMeasurements.Count; i++)
                if (string.Equals(SessionMeasurements[i].ProfileName, profileName, StringComparison.OrdinalIgnoreCase))
                    count++;
            return count;
        }

        static void UpdateAdviceMessage()
        {
            ProfileInfo profile = GetSelectedProfile();
            LastSessionInsight = BuildSessionInsight(profile);
            if (!string.IsNullOrEmpty(LastSessionInsight.MainAdvice))
                LastSessionAdviceText = "Advice: " + LastSessionInsight.MainAdvice;
            else if (LastSessionInsight.AdviceMessages.Count > 0)
                LastSessionAdviceText = "Advice: " + LastSessionInsight.AdviceMessages[0];
            else
                LastSessionAdviceText = "Advice: Measurement looked steady and balanced overall.";
            f.lblAdvice.Text = LastSessionAdviceText;
        }

        static void MaybeRecordSession()
        {
            DateTime now = DateTime.UtcNow;
            if (LastSessionRecordUtc != DateTime.MinValue && (now - LastSessionRecordUtc).TotalMilliseconds < 250)
                return;

            BalanceBoardSensorsF sensors = bb.WiimoteState.BalanceBoardState.SensorValuesKg;
            float tl = Math.Max(0.0f, sensors.TopLeft);
            float tr = Math.Max(0.0f, sensors.TopRight);
            float bl = Math.Max(0.0f, sensors.BottomLeft);
            float br = Math.Max(0.0f, sensors.BottomRight);

            float total = tl + tr + bl + br;
            float left = tl + bl;
            float right = tr + br;
            float front = tl + tr;
            float back = bl + br;
            float leftPct = total <= 0.0001f ? 50.0f : (left / total * 100.0f);
            float rightPct = 100.0f - leftPct;
            float frontPct = total <= 0.0001f ? 50.0f : (front / total * 100.0f);
            float backPct = 100.0f - frontPct;

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
                PressurePointX = (total <= 0.0001f ? 0.0f : ((right - left) / total)),
                PressurePointY = (total <= 0.0001f ? 0.0f : ((front - back) / total)),
                StabilityLevel = LastStabilityLevel
            });

            if (SessionMeasurements.Count > 2000)
                SessionMeasurements.RemoveAt(0);

            LastSessionRecordUtc = now;
        }

        static void UpdateSessionInfo()
        {
            ProfileInfo profile = GetSelectedProfile();
            string profileDesc = (profile == null ? "Default" : profile.Name);
            if (profile != null && profile.HeightCm > 0.0f)
                profileDesc += " (" + profile.HeightCm.ToString("0.0", CultureInfo.InvariantCulture) + " cm)";

            f.lblSessionInfo.Text = "Session samples: " + SessionMeasurements.Count.ToString() + "   Profile: " + profileDesc;
            if (LastSessionInsight == null || !string.Equals(LastSessionInsight.ProfileName, (profile == null ? "Default" : profile.Name), StringComparison.OrdinalIgnoreCase))
                LastSessionInsight = BuildSessionInsight(profile);

            f.lblSessionSummary.Text = TruncateForLabel(LastSessionInsight.SummaryText, 150);
            f.lblSessionComparison.Text = TruncateForLabel(LastSessionInsight.ComparisonText, 150);
            f.lblTrendHighlights.Text = TruncateForLabel(LastSessionInsight.TrendText, 150);
            f.lblPatternSummary.Text = TruncateForLabel(LastSessionInsight.RepeatedPatternText + " " + LastSessionInsight.StabilityPatternText + ".", 150);
            f.lblSessionReview.Text = TruncateForLabel(LastSessionInsight.ReviewText + " Main advice: " + LastSessionInsight.MainAdvice + ". Quality reason: " + LastSessionInsight.SessionQualityReasonText + ".", 190);
            f.pnlWeightTrend.Invalidate();
            f.pnlLeftRightTrend.Invalidate();
            f.pnlFrontBackTrend.Invalidate();
            f.pnlStabilityTrend.Invalidate();
        }

        static void SaveCurrentSessionToHistory(string profileName = null)
        {
            ProfileInfo profile = (string.IsNullOrEmpty(profileName) ? GetSelectedProfile() : FindProfileByName(profileName));
            if (profile == null) profile = new ProfileInfo() { Name = (string.IsNullOrEmpty(profileName) ? "Default" : profileName), HeightCm = 0.0f };
            SessionInsight insight = BuildSessionInsight(profile);
            if (!insight.HasEnoughSamples) return;
            if (insight.SessionQuality == ESessionQuality.UnstableSession) return;
            if (insight.EndedUtc == DateTime.MinValue) return;

            DateTime lastSavedEndUtc;
            if (LastSavedSessionEndUtcByProfile.TryGetValue(insight.ProfileName, out lastSavedEndUtc) && insight.EndedUtc <= lastSavedEndUtc)
                return;

            bool hasSimilarRecent = false;
            for (int i = SessionHistory.Count - 1; i >= 0; i--)
            {
                SessionHistoryRecord existing = SessionHistory[i];
                if (!string.Equals(existing.ProfileName, insight.ProfileName, StringComparison.OrdinalIgnoreCase))
                    continue;
                if ((DateTime.UtcNow - existing.TimestampUtc).TotalSeconds > 30.0)
                    break;
                if (Math.Abs(existing.StableWeightKg - insight.StableWeightKg) < 0.1f &&
                    Math.Abs(existing.AverageLeftPercent - insight.AverageLeftPercent) < 0.1f &&
                    Math.Abs(existing.AverageFrontPercent - insight.AverageFrontPercent) < 0.1f)
                {
                    hasSimilarRecent = true;
                    break;
                }
            }
            if (hasSimilarRecent) return;

            SessionHistoryRecord record = new SessionHistoryRecord();
            record.TimestampUtc = DateTime.UtcNow;
            record.ProfileName = insight.ProfileName;
            record.ProfileHeightCm = insight.ProfileHeightCm;
            record.SampleCount = insight.SampleCount;
            record.StableWeightKg = insight.StableWeightKg;
            record.AverageLeftPercent = insight.AverageLeftPercent;
            record.AverageFrontPercent = insight.AverageFrontPercent;
            record.AverageStabilityLevel = insight.AverageStabilityLevel;
            record.AveragePressurePointX = insight.AveragePressurePointX;
            record.AveragePressurePointY = insight.AveragePressurePointY;
            record.TimeToStabilizeSeconds = insight.TimeToStabilizeSeconds;
            record.StabilityScore = insight.StabilityScore;
            record.StabilityBandText = insight.StabilityBandText;
            record.StanceTendencyText = insight.StanceTendencyText;
            record.SessionQualityText = insight.SessionQualityText;
            record.WeightVsHeightText = insight.WeightVsHeightText;
            record.SummaryText = insight.SummaryText;
            record.ReviewText = insight.ReviewText;
            record.ComparisonText = insight.ComparisonText;
            record.TrendText = insight.TrendText;
            record.AdviceText = string.Join(" | ", insight.AdviceMessages.ToArray());
            SessionHistory.Add(record);
            LastSavedSessionEndUtcByProfile[insight.ProfileName] = insight.EndedUtc;

            if (SessionHistory.Count > 500)
                SessionHistory.RemoveRange(0, SessionHistory.Count - 500);

            try { SaveSessionHistory(); } catch { }
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
            string previousProfileName = LastActiveProfileName;
            SaveCurrentSessionToHistory(previousProfileName);
            ApplyProfileToInputs();
            UpdateWeightVsHeightIndicator();
            UpdateSessionInfo();
            ProfileInfo selected = GetSelectedProfile();
            LastActiveProfileName = (selected == null ? "Default" : selected.Name);
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

            string heightText = (f.txtProfileHeightCm.Text == null ? "" : f.txtProfileHeightCm.Text.Trim());
            float heightCm = 0.0f;
            if (heightText.Length > 0)
            {
                if (!float.TryParse(heightText, NumberStyles.Float, CultureInfo.InvariantCulture, out heightCm))
                {
                    MessageBox.Show(f, "Enter a numeric height in centimeters (or leave it blank).", "Profile", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (heightCm < 80.0f || heightCm > 250.0f)
                {
                    MessageBox.Show(f, "Height should be between 80 and 250 cm, or left blank.", "Profile", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
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
            UpdateSessionInfo();
        }

        static void btnClearSession_Click(object sender, EventArgs e)
        {
            SaveCurrentSessionToHistory();
            SessionMeasurements.Clear();
            LastSessionRecordUtc = DateTime.MinValue;
            LastSessionInsight = null;
            LastSessionAdviceText = "Advice: Waiting for measurement...";
            f.lblAdvice.Text = LastSessionAdviceText;
            UpdateSessionInfo();
        }

        static void btnExportCsv_Click(object sender, EventArgs e)
        {
            ProfileInfo selectedProfile = GetSelectedProfile();
            string selectedName = (selectedProfile == null ? "Default" : selectedProfile.Name);
            if (CountSamplesForProfile(selectedName) == 0)
            {
                MessageBox.Show(f, "No samples for the selected profile to export yet.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            ProfileInfo selectedProfile = GetSelectedProfile();
            string selectedName = (selectedProfile == null ? "Default" : selectedProfile.Name);
            if (CountSamplesForProfile(selectedName) == 0)
            {
                MessageBox.Show(f, "No samples for the selected profile to export yet.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            SessionInsight insight = BuildSessionInsight(GetSelectedProfile());
            string adviceText = string.Join(" | ", insight.AdviceMessages.ToArray());
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(CsvHeader);
            string selectedProfileName = insight.ProfileName;
            List<MeasurementSample> exportSamples = new List<MeasurementSample>();
            for (int i = 0; i < SessionMeasurements.Count; i++)
            {
                MeasurementSample current = SessionMeasurements[i];
                if (string.Equals(current.ProfileName, selectedProfileName, StringComparison.OrdinalIgnoreCase))
                    exportSamples.Add(current);
            }

            for (int i = 0; i < exportSamples.Count; i++)
            {
                MeasurementSample s = exportSamples[i];
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
                sb.Append(s.StabilityLevel.ToString(CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(CsvEscape(insight.SessionQualityText)); sb.Append(',');
                sb.Append(insight.TimeToStabilizeSeconds.ToString("0.0", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(insight.StabilityScore.ToString("0.0", CultureInfo.InvariantCulture)); sb.Append(',');
                sb.Append(CsvEscape(insight.StabilityBandText)); sb.Append(',');
                sb.Append(CsvEscape(insight.StanceTendencyText)); sb.Append(',');
                sb.Append(CsvEscape(insight.RepeatedPatternText)); sb.Append(',');
                sb.Append(CsvEscape(insight.MainAdvice)); sb.Append(',');
                sb.Append(CsvEscape(insight.ReviewText)); sb.Append(',');
                sb.Append(CsvEscape(insight.SummaryText)); sb.Append(',');
                sb.Append(CsvEscape(insight.ComparisonText)); sb.Append(',');
                sb.Append(CsvEscape(insight.TrendText)); sb.Append(',');
                sb.Append(CsvEscape(adviceText)); sb.Append(',');
                sb.Append(CsvEscape(insight.WeightVsHeightText));
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
            SessionInsight insight = BuildSessionInsight(profile);
            sb.AppendLine("{");
            sb.Append("  \"profile\": \"").Append(JsonEscape(profile == null ? "Default" : profile.Name)).AppendLine("\",");
            sb.Append("  \"profile_height_cm\": ").Append(profile == null ? "0.0" : profile.HeightCm.ToString("0.0", CultureInfo.InvariantCulture)).AppendLine(",");
            sb.Append("  \"exported_at_utc\": \"").Append(DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)).AppendLine("\",");
            sb.AppendLine("  \"session_summary\": {");
            sb.Append("    \"session_quality\": \"").Append(JsonEscape(insight.SessionQualityText)).AppendLine("\",");
            sb.Append("    \"session_quality_reason\": \"").Append(JsonEscape(insight.SessionQualityReasonText)).AppendLine("\",");
            sb.Append("    \"final_stable_weight_kg\": ").Append(insight.StableWeightKg.ToString("0.000", CultureInfo.InvariantCulture)).AppendLine(",");
            sb.Append("    \"time_to_stabilize_seconds\": ").Append(insight.TimeToStabilizeSeconds.ToString("0.0", CultureInfo.InvariantCulture)).AppendLine(",");
            sb.Append("    \"stability_score\": ").Append(insight.StabilityScore.ToString("0.0", CultureInfo.InvariantCulture)).AppendLine(",");
            sb.Append("    \"stability_band\": \"").Append(JsonEscape(insight.StabilityBandText)).AppendLine("\",");
            sb.Append("    \"stance_tendency\": \"").Append(JsonEscape(insight.StanceTendencyText)).AppendLine("\",");
            sb.Append("    \"repeated_pattern_if_available\": \"").Append(JsonEscape(insight.RepeatedPatternText)).AppendLine("\",");
            sb.Append("    \"stability_pattern_if_available\": \"").Append(JsonEscape(insight.StabilityPatternText)).AppendLine("\",");
            sb.Append("    \"matches_usual_pattern\": ").Append(insight.MatchesUsualPattern ? "true" : "false").AppendLine(",");
            sb.Append("    \"main_advice\": \"").Append(JsonEscape(insight.MainAdvice)).AppendLine("\",");
            sb.Append("    \"review_text\": \"").Append(JsonEscape(insight.ReviewText)).AppendLine("\",");
            sb.Append("    \"summary_text\": \"").Append(JsonEscape(insight.SummaryText)).AppendLine("\",");
            sb.Append("    \"comparison_text\": \"").Append(JsonEscape(insight.ComparisonText)).AppendLine("\",");
            sb.Append("    \"recent_trend_interpretation\": \"").Append(JsonEscape(insight.TrendText)).AppendLine("\",");
            sb.Append("    \"weight_vs_height_text\": \"").Append(JsonEscape(insight.WeightVsHeightText)).AppendLine("\",");
            sb.Append("    \"advice\": [");
            for (int i = 0; i < insight.AdviceMessages.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append("\"").Append(JsonEscape(insight.AdviceMessages[i])).Append("\"");
            }
            sb.AppendLine("]");
            sb.AppendLine("  },");
            sb.AppendLine("  \"samples\": [");
            string selectedProfileName = insight.ProfileName;
            List<MeasurementSample> exportSamples = new List<MeasurementSample>();
            for (int i = 0; i < SessionMeasurements.Count; i++)
            {
                MeasurementSample current = SessionMeasurements[i];
                if (string.Equals(current.ProfileName, selectedProfileName, StringComparison.OrdinalIgnoreCase))
                    exportSamples.Add(current);
            }

            for (int i = 0; i < exportSamples.Count; i++)
            {
                MeasurementSample s = exportSamples[i];
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
                if (i < exportSamples.Count - 1) sb.Append(',');
                sb.AppendLine();
            }
            sb.AppendLine("  ]");
            sb.AppendLine("}");
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        static List<SessionHistoryRecord> GetRecentProfileHistory(string profileName, int maxCount, bool comparableOnly)
        {
            List<SessionHistoryRecord> list = new List<SessionHistoryRecord>();
            for (int i = SessionHistory.Count - 1; i >= 0 && list.Count < maxCount; i--)
            {
                SessionHistoryRecord record = SessionHistory[i];
                if (!string.Equals(record.ProfileName, profileName, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (comparableOnly && !IsRecordComparable(record))
                    continue;
                list.Add(record);
            }
            list.Reverse();
            return list;
        }

        static void DrawTrendGraph(Graphics g, Rectangle rc, List<float> values, float centerLineValue, float minValue, float maxValue, Color lineColor)
        {
            g.Clear(Color.White);
            if (rc.Width < 4 || rc.Height < 4) return;

            using (Pen borderPen = new Pen(Color.LightGray, 1.0f))
            using (Pen centerPen = new Pen(Color.Gainsboro, 1.0f))
            using (Pen trendPen = new Pen(lineColor, 2.0f))
            using (Brush pointBrush = new SolidBrush(lineColor))
            {
                g.DrawRectangle(borderPen, 0, 0, rc.Width - 1, rc.Height - 1);
                if (maxValue <= minValue) maxValue = minValue + 1.0f;
                float centerY = rc.Height - 1 - ((centerLineValue - minValue) / (maxValue - minValue)) * (rc.Height - 2);
                g.DrawLine(centerPen, 1, centerY, rc.Width - 2, centerY);

                if (values.Count < 2) return;

                System.Drawing.PointF[] pts = new System.Drawing.PointF[values.Count];
                for (int i = 0; i < values.Count; i++)
                {
                    float t = (values.Count == 1 ? 0.0f : (float)i / (values.Count - 1));
                    float x = 1 + t * (rc.Width - 3);
                    float y = rc.Height - 1 - ((values[i] - minValue) / (maxValue - minValue)) * (rc.Height - 2);
                    pts[i] = new System.Drawing.PointF(x, y);
                }
                g.DrawLines(trendPen, pts);
                for (int i = 0; i < pts.Length; i++)
                    g.FillEllipse(pointBrush, pts[i].X - 2.0f, pts[i].Y - 2.0f, 4.0f, 4.0f);
            }
        }

        static void pnlWeightTrend_Paint(object sender, PaintEventArgs e)
        {
            ProfileInfo profile = GetSelectedProfile();
            string name = (profile == null ? "Default" : profile.Name);
            List<SessionHistoryRecord> history = GetRecentProfileHistory(name, 12, true);
            List<float> values = new List<float>();
            float min = float.MaxValue, max = float.MinValue;
            for (int i = 0; i < history.Count; i++)
            {
                float v = history[i].StableWeightKg;
                values.Add(v);
                if (v < min) min = v;
                if (v > max) max = v;
            }
            if (values.Count == 0) { values.Add(0.0f); min = -1.0f; max = 1.0f; }
            DrawTrendGraph(e.Graphics, f.pnlWeightTrend.ClientRectangle, values, values[values.Count - 1], min - 0.4f, max + 0.4f, Color.SteelBlue);
        }

        static void pnlLeftRightTrend_Paint(object sender, PaintEventArgs e)
        {
            ProfileInfo profile = GetSelectedProfile();
            string name = (profile == null ? "Default" : profile.Name);
            List<SessionHistoryRecord> history = GetRecentProfileHistory(name, 12, true);
            List<float> values = new List<float>();
            for (int i = 0; i < history.Count; i++) values.Add(history[i].AverageLeftPercent);
            if (values.Count == 0) values.Add(50.0f);
            DrawTrendGraph(e.Graphics, f.pnlLeftRightTrend.ClientRectangle, values, 50.0f, 44.0f, 56.0f, Color.DarkOrange);
        }

        static void pnlFrontBackTrend_Paint(object sender, PaintEventArgs e)
        {
            ProfileInfo profile = GetSelectedProfile();
            string name = (profile == null ? "Default" : profile.Name);
            List<SessionHistoryRecord> history = GetRecentProfileHistory(name, 12, true);
            List<float> values = new List<float>();
            for (int i = 0; i < history.Count; i++) values.Add(history[i].AverageFrontPercent);
            if (values.Count == 0) values.Add(50.0f);
            DrawTrendGraph(e.Graphics, f.pnlFrontBackTrend.ClientRectangle, values, 50.0f, 44.0f, 56.0f, Color.MediumSeaGreen);
        }

        static void pnlStabilityTrend_Paint(object sender, PaintEventArgs e)
        {
            ProfileInfo profile = GetSelectedProfile();
            string name = (profile == null ? "Default" : profile.Name);
            List<SessionHistoryRecord> history = GetRecentProfileHistory(name, 12, true);
            List<float> values = new List<float>();
            for (int i = 0; i < history.Count; i++) values.Add(history[i].AverageStabilityLevel);
            if (values.Count == 0) values.Add(3.0f);
            DrawTrendGraph(e.Graphics, f.pnlStabilityTrend.ClientRectangle, values, 3.5f, 1.0f, 5.0f, Color.MediumPurple);
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
    }
}
