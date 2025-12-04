using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Wpf.Ui.Controls;
using ZibllWindows;

namespace ZibllWindows
{
    class ActOffice
    {
        public static MainWindow mainW = (MainWindow)Application.Current.MainWindow;

        public static async Task OfficeActivate()
        {
            // debug info
            string kmsServerDbg, activateDbg;

            // ospp root
            string root = "";

            // procinfo
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cscript.exe",
                WorkingDirectory = @"C:\",

                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            if (IsOfficeActivated(root))
            {
                return;
            }

            Retail2Vol(root);

            Process kmsServer = new Process
            {
                StartInfo = startInfo
            };
            try
            {
                kmsServer.Start();
            }
            catch (Exception err)
            {
                await ShowError("Exception caught", err.ToString());
                return;
            }
            kmsServerDbg = kmsServer.StandardOutput.ReadToEnd();
            kmsServer.WaitForExit();

            // apply
            startInfo.Arguments = "//Nologo ospp.vbs /act";
            Process activate = new Process
            {
                StartInfo = startInfo
            };
            activate.Start();
            activateDbg = activate.StandardOutput.ReadToEnd();
            activate.WaitForExit();
        }

        public static bool IsOfficeActivated(string root)
        {
            // make vol
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cscript.exe",
                WorkingDirectory = root,
                Arguments = @"//Nologo ospp.vbs /dstatus",

                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            // check license status
            Process activate = new Process
            {
                StartInfo = startInfo
            };
            activate.Start();
            string activateDbg = activate.StandardOutput.ReadToEnd();
            activate.WaitForExit();

            return false;
        }

        public static void Retail2Vol(string installRoot)
        /*
         * try to install vol key and licenses
         * product key is of Office Pro Plus
         */
        {
            string dstatus;
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "cscript.exe",
                    WorkingDirectory = installRoot,
                    Arguments = @"//Nologo ospp.vbs /dstatus",

                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                // check license status
                Process checkLicense = new Process
                {
                    StartInfo = startInfo
                };
                checkLicense.Start();
                dstatus = checkLicense.StandardOutput.ReadToEnd();
                checkLicense.WaitForExit();
            }
            catch (Exception err)
            {
                // We are in non-async method — dispatch async dialog to UI thread
                mainW.Dispatcher.Invoke(async () =>
                {
                    await ShowError("Retail2Vol", err.ToString());
                });
                return;
            }

            if (dstatus.ToLower().Contains("volume"))
            {
                return;
            }

            // convert confirmation: Util.YesNo was used previously; you can keep it,
            // or replace with ShowYesNoAsync. For now keep original call to Util.YesNo.
            if (!Util.YesNo("You are not using a VOL version, try converting?", "Retail2Vol"))
            {
                return;
            }
            string licenseDir = installRoot + @"..\root\License";
            string key, visioKey, version;

            // handle different versions
            if (installRoot.EndsWith("Office16"))
            {
                licenseDir += @"16\";
                if (dstatus.Contains("Office 19"))
                {
                    key = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP";
                    visioKey = "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB";
                    version = "2019";
                }
                else
                {
                    key = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99";
                    visioKey = "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK";
                    version = "2016";
                }
            }
            else if (installRoot.EndsWith("Office15"))
            {
                licenseDir += @"15\";
                key = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT";
                visioKey = "C2FG9-N6J68-H8BTJ-BW3QX-RM3B3";
                version = "2013";
            }
            else if (installRoot.EndsWith("Office14"))
            {
                licenseDir += @"14\";
                key = "VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB";
                visioKey = "7MCW8-VRQVK-G677T-PDJCM-Q8TCP";
                version = "2010";
            }
            else
            {
                mainW.Dispatcher.Invoke(async () =>
                {
                    await ShowError("Goodbye", "No compatible Office version found, exit?");
                });
                return;
            }

            // try to convert retail version to VOL
            string log;
            log = Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "ProPlusVL_KMS_Client-ppd.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "ProPlusVL_KMS_Client-ul.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "ProPlusVL_KMS_Client-ul-oob.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "VisioProVL_KMS_Client-ppd.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "VisioProVL_KMS_Client-ul.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "VisioProVL_KMS_Client-ul-oob.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "pkeyconfig-office.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inpkey:" + key, installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inpkey:" + visioKey, installRoot, false);

            // tell user what happened
            mainW.Dispatcher.Invoke(async () =>
            {
                await ShowInfo("Note: You are NOT using volume version",
                    "Converting Office " + version + " retail version to volume version...\n" + log);
            });
        }

        public static void InstallOffice()
        {
            if (!File.Exists("setup.exe") || !File.Exists("office-proplus.xml"))
            {
                // dispatch an async dialog to UI
                mainW.Dispatcher.Invoke(async () =>
                {
                    await ShowError("Files missing", "Make sure setup.exe and office-proplus.xml exist in current directory");
                });
                return;
            }
            Util.RunProcess("setup.exe", "/configure office-proplus.xml", "./", true);
        }

        public static bool OfficeEnv()
        {

            // look for Office's install path, where OSPP.VBS can be found
            try
            {
                RegistryKey localKey;
                if (Environment.Is64BitOperatingSystem)
                {
                    localKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                }
                else
                {
                    localKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
                }

                string officepath = "";
                RegistryKey officeBaseKey = localKey.OpenSubKey(@"SOFTWARE\Microsoft\Office");
                if (officeBaseKey == null)
                {
                    throw new Exception("Office not installed");
                }
                if (officeBaseKey.OpenSubKey(@"16.0", false) != null)
                {
                    var k = officeBaseKey.OpenSubKey(@"16.0\Word\InstallRoot");
                    if (k == null)
                    {
                        throw new Exception("Office not installed or corrupted");
                    }

                    var val = k.GetValue("Path");
                    if (val == null)
                    {
                        throw new Exception("Office installation corrupted");
                    }
                    officeBaseKey.Close();
                    officepath = val.ToString();
                    if (officepath.Contains("root"))
                    // Office 2019 can only be installed via Click-To-Run, therefore we get "C:\Program Files\Microsoft Office\root\Office16\",
                    // otherwise we get "C:\Program Files\Microsoft Office\Office16\"
                    {
                        // OSPP.VBS is still in "C:\Program Files\Microsoft Office\Office16\"
                        officepath = officepath.Replace("root", "");
                        // Office 2019
                    }
                    else
                    {
                        // Office 2016
                    }
                }
                else if (officeBaseKey.OpenSubKey(@"15.0", false) != null)
                {
                    officepath = officeBaseKey.OpenSubKey(@"15.0\Word\InstallRoot").GetValue("Path").ToString();
                    officeBaseKey.Close();
                }
                else if (officeBaseKey.OpenSubKey(@"14.0", false) != null)
                {
                    officepath = officeBaseKey.OpenSubKey(@"14.0\Word\InstallRoot").GetValue("Path").ToString();
                    officeBaseKey.Close();
                }
                else
                {
                    throw new Exception("Only works with Office 2010 and/or above");
                }

                return true;
            }
            catch (Exception err)
            {
                // show error and optionally install
                mainW.Dispatcher.Invoke(async () =>
                {
                    await ShowError("Error detecting Office path", "Office installation not detected:\n" + err.ToString());
                    bool yes = await ShowYesNo("Install Office", "Download and install Office 2021 with Office Deployment Tool?");
                    if (yes)
                    {
                        InstallOffice();
                    }
                });
                return false;
            }
        }

        #region WPF-UI helper wrappers

        // informational (OK only)
        private static async Task ShowInfo(string title, string content)
        {
            var uiMessageBox = new Wpf.Ui.Controls.MessageBox
            {
                Title = title,
                Content = content,
                CloseButtonText = "OK"
            };
            await uiMessageBox.ShowDialogAsync();
        }

        // error (OK only)
        private static async Task ShowError(string title, string content)
        {
            var uiMessageBox = new Wpf.Ui.Controls.MessageBox
            {
                Title = title,
                Content = content,
                CloseButtonText = "OK"
            };
            await uiMessageBox.ShowDialogAsync();
        }

        // yes/no dialog -> returns true when primary/Yes is clicked
        private static async Task<bool> ShowYesNo(string title, string content)
        {
            var uiMessageBox = new Wpf.Ui.Controls.MessageBox
            {
                Title = title,
                Content = content,
                PrimaryButtonText = "Yes",
                CloseButtonText = "No"
            };

            var result = await uiMessageBox.ShowDialogAsync();
            // Wpf.Ui uses MessageBoxResult (Primary == Affirmative)
            return result == Wpf.Ui.Controls.MessageBoxResult.Primary;
        }

        #endregion
    }
}
