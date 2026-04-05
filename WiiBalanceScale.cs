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

        static bool CanShowUnicode = GetCanShowUnicode();
        static char CharFilledStar   = (CanShowUnicode ? '\u2739' : '\u00AE');
        static char CharHollowCircle = (CanShowUnicode ? '\u3007' : '\u00A1');
        static char CharHourglass    = (CanShowUnicode ? '\u23F3' : '\u0036');
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
        enum EUnit { Kg, Lb, Stone };

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
            if (CanShowUnicode)
            {
                f.lblQuality.Location = new System.Drawing.Point(f.lblQuality.Location.X, f.lblQuality.Location.Y - 10);
                f.lblQuality.Font = new System.Drawing.Font("Segoe UI Symbol", 60F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            }
            f.btnReset.Click += (object sender, System.EventArgs e) =>
            {
                if (HistoryBest <= 0) return;
                float HistorySum = 0.0f;
                for (int i = 0; i < HistoryBest; i++)
                    HistorySum += History[(HistoryCursor + History.Length - i) % History.Length];
                ZeroedWeight = HistorySum / HistoryBest;
            };
            System.EventHandler unitRadioButton_Change = (object sender, EventArgs e) =>
            {
                if (!(sender as RadioButton).Checked) return;
                if      (sender == f.unitSelectorKg)    SelectedUnit = EUnit.Kg;
                else if (sender == f.unitSelectorLb)    SelectedUnit = EUnit.Lb;
                else if (sender == f.unitSelectorStone) SelectedUnit = EUnit.Stone;
            };
            f.unitSelectorKg.CheckedChanged += unitRadioButton_Change;
            f.unitSelectorLb.CheckedChanged += unitRadioButton_Change;
            f.unitSelectorStone.CheckedChanged += unitRadioButton_Change;
            f.unitSelectorKg.Checked = true;

            ConnectBalanceBoard(false);
            if (f == null) return; //connecting required application restart, end this process here

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

            //start with half full quality bar
            HistoryCursor = HistoryBest = History.Length / 2;
            for (int i = 0; i < History.Length; i++)
                History[i] = (i > HistoryCursor ? float.MinValue : ZeroedWeight);
        }

        static void BoardTimer_Tick(object sender, System.EventArgs e)
        {
            if (cm != null)
            {
                if (cm.IsRunning())
                {
                    f.lblWeight.Text = "WAIT...";
                    f.lblQuality.Text = (f.lblQuality.Text.Length >= 5 ? "" : f.lblQuality.Text) + CharHourglass;
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
                    System.Windows.Forms.MessageBox.Show(f, GetConnectionErrorMessage(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Shutdown();
                    return;
                }

                if (!cm.DidConnect())
                {
                    ApplyScannerStatusToConnectionError();

                    if (LastConnectionError == EConnectionError.None)
                        SetConnectionError(EConnectionError.NoDeviceFound, "Bluetooth scan timed out without finding a Wii Balance Board.");

                    BoardTimer.Stop();
                    System.Windows.Forms.MessageBox.Show(f, GetConnectionErrorMessage(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Shutdown();
                    return;
                }

                if (LastConnectionError == EConnectionError.None)
                    SetConnectionError(EConnectionError.NoDeviceFound, "Connection scan finished without a matching device.");

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
                if (System.Math.Abs(MaxHist - HistoryEntry) > 1.0f) break;
                if (System.Math.Abs(MinHist - HistoryEntry) > 1.0f) break;
                if (HistoryEntry > MaxHist) MaxHist = HistoryEntry;
                if (HistoryEntry < MinHist) MinHist = HistoryEntry;
                float Diff = System.Math.Max(System.Math.Abs(HistoryEntry - kg), System.Math.Abs((HistorySum + HistoryEntry) / (HistoryBest + 1) - kg));
                if (Diff > MaxDiff) MaxDiff = Diff;
                if (Diff > 1.0f) break;
                HistorySum += HistoryEntry;
            }

            if (HistoryBest <= 0) return;
            kg = HistorySum / HistoryBest - ZeroedWeight;

            float accuracy = 1.0f / HistoryBest;
            float weight = (float)System.Math.Floor(kg / accuracy + 0.5f) * accuracy;

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

            f.lblQuality.Text = "";
            for (int i = 0; i < 5; i++)
                f.lblQuality.Text += (i < ((HistoryBest + 5) / (History.Length / 5)) ? CharFilledStar : CharHollowCircle);
        }
    }
}
