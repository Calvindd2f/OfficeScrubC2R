
using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.Security.Principal;
using System.Security.AccessControl;

/// <summary>
/// OffScrubC2R
/// 
/// Author: Microsoft Customer Support Services
/// Copyright (c) 2014 - 2016 Microsoft Corporation
/// 
/// Script to remove Office Click To Run (C2R) products
/// when a standard uninstall is no longer possible.
///
/// Scope: Office 2013, 2016, and O365 C2R products
/// </summary>/// <summary>
/// Declaration of constants
/// </summary>
public static class Constants
{
    public const string SCRIPTVERSION   = "1.0.0";
    public const string SCRIPTFILE      = "OffScrubC2R.vbs";
    public const string SCRIPTNAME      = "OffScrubC2R";
    public const string RETVALFILE      = "ScrubRetValFile.txt";
    public const string ONAME           = "Office C2R / O365";
    public const int HKCR            = 0x80000000;
    public const int HKCU            = 0x80000001;
    public const int HKLM            = 0x80000002;
    public const int HKU             = 0x80000003;
    public const int PRODLEN         = 13;
    public const int SQUISHED        = 20;
    public const int COMPRESSED      = 32;
    public const string REG_ARP      = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\";
    public const int VB_YES          = 6;
    public const int VB_NO           = 7;

    public const int ERROR_SUCCESS                 = 0;
    public const int ERROR_FAIL                    = 1;
    public const int ERROR_REBOOT_REQUIRED         = 2;
    public const int ERROR_USERCANCEL              = 4;
    public const int ERROR_STAGE1                  = 8;
    public const int ERROR_STAGE2                  = 16;
    public const int ERROR_INCOMPLETE              = 32;
    public const int ERROR_DCAF_FAILURE            = 64;
    public const int ERROR_ELEVATION_USERDECLINED  = 128;
    public const int ERROR_ELEVATION               = 256;
    public const int ERROR_SCRIPTINIT              = 512;
    public const int ERROR_RELAUNCH                = 1024;
    public const int ERROR_UNKNOWN                 = 2048;
    public const int ERROR_ALL                     = 4095;
    public const int ERROR_USER_ABORT              = -1073741510;
    public const int ERROR_SUCCESS_CONFIG_COMPLETE = 1728;
    public const int ERROR_SUCCESS_REBOOT_REQUIRED = 3010;
}

/// <summary>
/// Declaration of variables
/// </summary>

/// <objects>Object-type variables</object>
    object oFso, oMsi, oReg, oWShell, oWmiLocal, oShellApp;
    object ComputerItem, Key, Item, LogStream, TmpKey;

/// <Array>Array</Array>
    object[] arrVersion;

/// <Dictionaries>Dictionaries (using Dictionary<string, object> as a generic replacement)</Dictionaries>
    Dictionary<string, object> dicKeepLis, dicApps, dicKeepFolder, dicDelRegKey, dicKeepReg, dicSC;
    Dictionary<string, object> dicInstalledSku, dicRemoveSku, dicKeepSku, dicC2RSuite, dicDelInUse, dicDelFolder;

/// <Strings>Strings</Strings>
    string sAppData, sScrubDir, sProgramFiles, sProgramFilesX86, sCommonProgramFiles;
    string sAllusersProfile, sOSVersion, sWinDir, sWICacheDir, sCommonProgramFilesX86;
    string sProgramData, sPackageFolder, sLocalAppData, sOInstallRoot, sSkuRemoveList;
    string sOSinfo, sDefault, sTemp, sTmp, sCmd, sLogDir, sProfilesDirectory, sArpUninstallCmd;
    string sRetVal, sScriptDir, sPackageGuid, sValue, sActiveConfiguration, sNotepad;

/// <Integers>Integers</Integers>
    int iVersionNT, iError, iProcCloseCnt;

/// <Booleans>Booleans</Booleans>
    bool f64, fLogInitialized, fNoCancel, fRemoveOse, fDetectOnly, fQuiet, fForce;
    bool fC2R, fRemoveAll, fRebootRequired, fRerun, fSetRunOnce, fTestRerun;
    bool fIsElevated, fNoElevate, fUserConsent, fCScript, fReturnErrorOrSuccess;
    bool fClearTaskBand, fSkipSD, fUnpinMode, fKeepLicense, fOffline, fForceArpUninstall;

/// <Named pipe and file system helper variables>Named pipe and file system helper variables
    string pipename;
    object pipeStream; /// <type>type depends on implementation, e.g., NamedPipeServerStream</type>
    object fs;         /// <type>type depends on use, e.g., FileStream, FileSystemInfo, etc.</type>
/// </Named pipe and file system helper variables>

// -------------------------------------------------------------------------------
//                                   main
//
//                           Main section of script
// -------------------------------------------------------------------------------

public void main()
{
    // Initialize required settings and objects
    Initialize();

    // Call the command line parser
    ParseCmdLine();

    // -----------------------------
    // Stage # 0 - Basic detection
    // -----------------------------
    LogH("Stage # 0 \"Basic detection\"");
    LogY("stage0");

    // Ensure integrity of WI metadata which could otherwise break used APIs
    LogH1($"Ensure Windows Installer metadata integrity ({DateTime.Now:T})");
    EnsureValidWIMetadata(HKCU, @"Software\Classes\Installer\Products", COMPRESSED);
    EnsureValidWIMetadata(HKCR, @"Installer\Products", COMPRESSED);
    EnsureValidWIMetadata(HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products", COMPRESSED);
    EnsureValidWIMetadata(HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components", COMPRESSED);
    EnsureValidWIMetadata(HKCR, @"Installer\Components", COMPRESSED);

    // Build a list with installed/registered Office products
    FindInstalledOProducts();
    if (dicC2RSuite.Count > 0)
    {
        Log("Registered ARP product(s) found:");
        foreach (var key in dicC2RSuite.Keys)
        {
            Log($" - {key} - {dicC2RSuite[key]}");
        }
        // Optionally: foreach (var val in dicC2RSuite.Values) Log($" - {val}");
    }
    else
    {
        Log("No registered product(s) found");
    }

    // Locate the C2R %PackageFolder% and the PackageGuid
    sPackageFolder = "";
    if (RegReadValue(HKLM, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun", "PackageFolder", out sValue, "REG_SZ"))
    {
        sPackageFolder = sValue;
    }
    else if (RegReadValue(HKLM, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun", "PackageFolder", out sValue, "REG_SZ"))
    {
        sPackageFolder = sValue;
    }
    else if (RegReadValue(HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun", "PackageFolder", out sValue, "REG_SZ"))
    {
        sPackageFolder = sValue;
    }

    // If sPackageFolder is invalid, set it to the C2R registry reference string
    if (string.IsNullOrEmpty(sPackageFolder))
    {
        var pf = ExpandEnv("%programfiles%");
        var pf86 = ExpandEnv("%programfiles(x86)%");

        if (FolderExists($"{pf}\\Microsoft Office 15"))
        {
            sPackageFolder = $"{pf}\\Microsoft Office 15";
        }
        else if (FolderExists($"{pf}\\Microsoft Office 16"))
        {
            sPackageFolder = $"{pf}\\Microsoft Office 16";
        }
        else if (FolderExists($"{pf}\\Microsoft Office\\PackageManifests"))
        {
            sPackageFolder = $"{pf}\\Microsoft Office";
        }
        else if (FolderExists($"{pf86}\\Microsoft Office\\PackageManifests"))
        {
            sPackageFolder = $"{pf86}\\Microsoft Office";
        }
    }

    sPackageGuid = "";
    if (RegReadValue(HKLM, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun", "PackageGUID", out sValue, "REG_SZ"))
    {
        sPackageGuid = sValue;
    }
    else if (RegReadValue(HKLM, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun", "PackageGUID", out sValue, "REG_SZ"))
    {
        sPackageGuid = sValue;
    }
    else if (RegReadValue(HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun", "PackageGUID", out sValue, "REG_SZ"))
    {
        sPackageGuid = sValue;
    }

    // Init complete. Reset the return value
    ClearError(ERROR_SCRIPTINIT);

    // -----------------------
    // Stage # 1 - Uninstall  |
    // -----------------------
    LogH($"Stage # 1 \"Uninstall\"");
    LogY("stage1");

    // Clean OSPP
    LogH1("Clean OSPP");
    if (!fKeepLicense)
    {
        CleanOSPP();
    }

    // Clean vNext
    LogH1("Clean vNext Licenses");
    if (!fKeepLicense)
    {
        ClearVNextLicCache();
    }

    // End all running Office applications
    LogH1("End running processes");
    // LogY("stage2");
    if (dicKeepSku.Count == 0)
    {
        ClearShellIntegrationReg();
    }
    CloseOfficeApps();

    // Remove scheduled tasks which might interfere with uninstall
    if (!fDetectOnly)
    {
        DelSchtasks();
    }

    // Unpin shortcuts
    // Need to unpin as long as the shortcuts are still valid!
    LogH1("Clean shortcuts");
    // LogY("stage3");
    CleanShortcuts(sAllusersProfile, true, true);
    CleanShortcuts(sProfilesDirectory, true, true);

    // Uninstall
    LogH1($"Remove {ONAME}");
    Uninstall();

    // ---------------------
    // Stage # 2 - CleanUp |
    // ---------------------
    LogH($"Stage # 2 \"CleanUp\"");
    LogY("stage4");

    // Cleanup registry data
    RegWipe();

    // Cleanup files
    FileWipe();

    // For test purposes only!
    if (fTestRerun)
    {
        LogH2("Enforcing 'Rerun' mode for test purposes");
        fRebootRequired = true;
        SetError(ERROR_REBOOT_REQUIRED);
        Rerun();
    }

    // Ensure Explorer runs
    RestoreExplorer();

    // ------------------
    // Stage #3 - Exit
    // ------------------
    ExitScript();

    // Returncode and reboot handler
    void ExitScript()
    {
        // Update cached error and quit
        if ((iError & (ERROR_FAIL | ERROR_INCOMPLETE)) == 0)
        {
            RegDeleteValue(HKCU, @"SOFTWARE\Microsoft\Office\15.0\CleanC2R", "Rerun", false);
        }
        SetRetVal(iError);

        // log result
        if ((iError & ERROR_INCOMPLETE) != 0)
        {
            LogH2($"Removal result: {iError} - INCOMPLETE. Uninstall requires a system reboot to complete.");
        }
        else
        {
            string sTmp = " - SUCCESS";
            if ((iError & ERROR_USERCANCEL) != 0) sTmp = " - USER CANCELED";
            if ((iError & ERROR_FAIL) != 0) sTmp = " - FAIL";
            LogH2($"Removal result: {iError}{sTmp}");
        }

        if ((iError & ERROR_FAIL) != 0)
        {
            if ((iError & ERROR_REBOOT_REQUIRED) != 0) Log(" - Reboot required");
            if ((iError & ERROR_USERCANCEL) != 0) Log(" - User cancel");
            if ((iError & ERROR_STAGE1) != 0) Log(" - Msiexec failed");
            if ((iError & ERROR_STAGE2) != 0) Log(" - Cleanup failed");
            if ((iError & ERROR_INCOMPLETE) != 0) Log(" - Removal incomplete. Rerun after reboot needed");
            if ((iError & ERROR_DCAF_FAILURE) != 0) Log(" - Second attempt cleanup still incomplete");
            if ((iError & ERROR_ELEVATION_USERDECLINED) != 0) Log(" - User declined elevation");
            if ((iError & ERROR_ELEVATION) != 0) Log(" - Elevation failed");
            if ((iError & ERROR_SCRIPTINIT) != 0) Log(" - Initialization error");
            if ((iError & ERROR_RELAUNCH) != 0) Log(" - Unhandled error during relaunch attempt");
            if ((iError & ERROR_UNKNOWN) != 0) Log(" - Unknown error");
            // ERROR_USER_ABORT is only valid for the temporary cached error file
            // if ((iError & ERROR_USER_ABORT) != 0) Log(" - Process terminated by user");
        }

        LogH2("Removal end.");
        LogY("stage5");

        // Check if we need to show a simplified return code
        // 0 = Success, Non Zero = Error
        if (fReturnErrorOrSuccess)
        {
            bool fOverallSuccess = true;
            if ((iError & ERROR_USERCANCEL) != 0) fOverallSuccess = false;
            if ((iError & ERROR_STAGE2) != 0) fOverallSuccess = false;
            if ((iError & ERROR_DCAF_FAILURE) != 0) fOverallSuccess = false;
            if ((iError & ERROR_ELEVATION_USERDECLINED) != 0) fOverallSuccess = false;
            if ((iError & ERROR_ELEVATION) != 0) fOverallSuccess = false;
            if ((iError & ERROR_SCRIPTINIT) != 0) fOverallSuccess = false;
            if ((iError & ERROR_RELAUNCH) != 0) fOverallSuccess = false;
            if ((iError & ERROR_UNKNOWN) != 0) fOverallSuccess = false;

            string sTmp = "ReturnErrorOrSuccess switch has been set. The current value return code translates to: ";
            if (fOverallSuccess)
            {
                iError = ERROR_SUCCESS;
                LogY("result:stage5:true");
                Log($"{sTmp}SUCCESS");
            }
            else
            {
                LogY("result:stage5:false");
                Log($"{sTmp}ERROR");
            }
        }

        // Reboot handling
        if (fRebootRequired)
        {
            LogY("reboot");
            string sPrompt = "In order to complete uninstall, a system reboot is necessary. Would you like to reboot now?";
            if (!fQuiet)
            {
                // Use modern message box and system reboot call
                var dr = System.Windows.Forms.MessageBox.Show(
                    sPrompt,
                    SCRIPTNAME + " - Reboot Required",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question
                );
                if (dr == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        // Reboot the system using WMI
                        var scope = new System.Management.ManagementScope(@"\\.\root\cimv2");
                        var osQuery = new System.Management.ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                        using (var mos = new System.Management.ManagementObjectSearcher(scope, osQuery))
                        {
                            foreach (System.Management.ManagementObject os in mos.Get())
                            {
                                os.InvokeMethod("Reboot", null);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Failed to reboot system: {ex.Message}");
                    }
                }
            }
        }

        LogY("ok");

        // Exit application with return code
        Environment.Exit(iError);
    }
}

//-------------------------------------------------------------------------------
//   Initialize
//
//   Configure defaults and initialize all required objects
//-------------------------------------------------------------------------------
public void Initialize()
{
    // set variable defaults
    //----------------------
    iError = Constants.ERROR_SUCCESS;
    iProcCloseCnt = 0;
    sLogDir = "";
    sPackageFolder = "";
    sArpUninstallCmd = "";
    f64 = false;
    fCScript = false;
    fLogInitialized = false;
    fNoCancel = false;
    fRemoveOse = false;
    fDetectOnly = false;
    fQuiet = true;
    fForce = false;
    fC2R = true;
    fRebootRequired = false;
    fRerun = false;
    fTestRerun = false;
    fIsElevated = false;
    fNoElevate = false;
    fSetRunOnce = false;
    fUserConsent = false;
    fReturnErrorOrSuccess = false;
    fSkipSD = false;
    fClearTaskBand = false;
    fUnpinMode = false;
    fKeepLicense = false;
    fOffline = false;
    fForceArpUninstall = false;

    // create required objects
    //------------------------
    InitObjects();

    // get environment path values
    //----------------------------
    sAppData = ExpandEnv("%appdata%");
    sLocalAppData = ExpandEnv("%localappdata%");
    sTemp = ExpandEnv("%temp%");
    sAllusersProfile = ExpandEnv("%allusersprofile%");
    if (RegReadValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProfilesDirectory", out sValue, "REG_EXPAND_SZ"))
    {
        sProfilesDirectory = sValue;
    }
    if (!FolderExists(sProfilesDirectory))
    {
        sProfilesDirectory = Path.GetDirectoryName(ExpandEnv("%userprofile%"));
    }
    sProgramFiles = ExpandEnv("%programfiles%");
    //sProgramFilesX86   = deferred. Depends on operating system architecture check
    sCommonProgramFiles = ExpandEnv("%commonprogramfiles%");
    //sCommonProgramFilesX86 = deferred. Depends on operating system architecture check
    sProgramData = ExpandEnv("%programdata%");
    sWinDir = ExpandEnv("%windir%");
    //sPackageFolder      = deferred
    sWICacheDir = $"{sWinDir}\\Installer";
    sScrubDir = $"{sTemp}\\{Constants.SCRIPTNAME}";
    sScriptDir = System.Reflection.Assembly.GetExecutingAssembly().Location;
    sScriptDir = sScriptDir.Substring(0, sScriptDir.LastIndexOf("\\") + 1);
    sNotepad = $"{sWinDir}\\notepad.exe";

    // check if called to unpin a shortcut
    var args = Environment.GetCommandLineArgs();
    if (args.Length > 1)
    {
        if (args[1] == "UNPINSC" && args.Length > 2)
        {
            Unpin(args[2]);
            Environment.Exit(0);
        }
    }

    // ensure 64 bit host if needed
    if (System.Reflection.Assembly.GetExecutingAssembly().Location.ToLower().Contains("syswow64"))
    {
        RelaunchAs64Host();
    }

    // create the temp folder
    //-----------------------
    if (!Directory.Exists(sScrubDir))
    {
        Directory.CreateDirectory(sScrubDir);
    }

    // set the default logging directory
    //----------------------------------
    sLogDir = sScrubDir;

    // detect bitness of the operating system
    //----------------------------------------
    var scope = new System.Management.ManagementScope(@"\\.\root\cimv2");
    var query = new System.Management.ObjectQuery("SELECT * FROM Win32_ComputerSystem");
    using (var searcher = new System.Management.ManagementObjectSearcher(scope, query))
    {
        foreach (System.Management.ManagementObject item in searcher.Get())
        {
            var systemType = item["SystemType"]?.ToString() ?? "";
            f64 = systemType.Substring(0, Math.Min(3, systemType.Length)).Contains("64");
        }
    }
    if (f64)
    {
        sProgramFilesX86 = ExpandEnv("%programfiles(x86)%");
    }
    if (f64)
    {
        sCommonProgramFilesX86 = ExpandEnv("%CommonProgramFiles(x86)%");
    }

    // update error flag
    //------------------
    SetError(Constants.ERROR_SCRIPTINIT);

    // get Win32_OperatingSystem details
    //----------------------------------
    query = new System.Management.ObjectQuery("SELECT * FROM Win32_OperatingSystem");
    using (var searcher = new System.Management.ManagementObjectSearcher(scope, query))
    {
        foreach (System.Management.ManagementObject item in searcher.Get())
        {
            sOSinfo = (sOSinfo ?? "") + (item["Caption"]?.ToString() ?? "");
            sOSinfo += item["OtherTypeDescription"]?.ToString() ?? "";
            sOSinfo += $", SP {item["ServicePackMajorVersion"]}";
            sOSinfo += $", Version: {item["Version"]}";
            sOSVersion = item["Version"]?.ToString() ?? "";
            sOSinfo += $", Codepage: {item["CodeSet"]}";
            sOSinfo += $", Country Code: {item["CountryCode"]}";
            sOSinfo += $", Language: {item["OSLanguage"]}";
        }
    }

    // get VersionNT number
    //---------------------
    if (!string.IsNullOrEmpty(sOSVersion))
    {
        var delimiter = sOSVersion.Contains(".") ? "." : (sOSVersion.Contains(",") ? "," : " ");
        var arrVersion = sOSVersion.Split(new[] { delimiter }, StringSplitOptions.None);
        if (arrVersion.Length >= 2 && int.TryParse(arrVersion[0], out int major) && int.TryParse(arrVersion[1], out int minor))
        {
            iVersionNT = major * 100 + minor;
        }
    }

    // ensure sufficient registry permissions
    //--------------------------------------
    fIsElevated = CheckRegPermissions();
    if (!fIsElevated && !fNoElevate)
    {
        // try to relaunch elevated
        RelaunchElevated();

        // can't relaunch. Exit out
        SetError(Constants.ERROR_ELEVATION);
        var exeName = System.Reflection.Assembly.GetExecutingAssembly().Location;
        if (exeName.Length > 0 && char.ToUpper(exeName[exeName.LastIndexOf("\\") + 1]) == 'C')
        {
            if (!fLogInitialized) CreateLog();
            Log("Error: Insufficient registry access permissions - exiting");
        }
        SetRetVal(iError);
        //Environment.Exit(iError);
        ExitScript();
    }

    // clear error flags
    //------------------
    ClearError(Constants.ERROR_ELEVATION);
    ClearError(Constants.ERROR_SCRIPTINIT);

    // ensure CScript as engine
    //------------------------
    var fullName = System.Reflection.Assembly.GetExecutingAssembly().Location;
    fCScript = fullName.Length > 0 && char.ToUpper(fullName[fullName.LastIndexOf("\\") + 1]) == 'C';
    if (!fCScript && !fQuiet)
    {
        RelaunchAsCScript();
    }

    // set retval for file based logic
    //--------------------------------
    // value needs to be kept on 'user abort'
    SetRetVal(Constants.ERROR_USER_ABORT);

    // create dictionary objects
    //--------------------------
    dicInstalledSku = new Dictionary<string, object>();
    dicRemoveSku = new Dictionary<string, object>();
    dicKeepSku = new Dictionary<string, object>();
    dicKeepLis = new Dictionary<string, object>();
    dicKeepFolder = new Dictionary<string, object>();
    dicApps = new Dictionary<string, object>();
    dicDelRegKey = new Dictionary<string, object>();
    dicKeepReg = new Dictionary<string, object>();
    dicC2RSuite = new Dictionary<string, object>();
    dicDelInUse = new Dictionary<string, object>();
    dicDelFolder = new Dictionary<string, object>();
    dicSC = new Dictionary<string, object>();

    // add initial known .exe files that need to be closed
    //----------------------------------------------------
    dicApps.Add("appvshnotify.exe", "appvshnotify.exe");
    dicApps.Add("integratedoffice.exe", "integratedoffice.exe");
    dicApps.Add("integrator.exe", "integrator.exe");
    dicApps.Add("firstrun.exe", "firstrun.exe");
    //Adding setup.exe to the hard list of processes that are shut down will potentially break wrappers that invoke OffScrub
    //dicApps.Add("setup.exe", "setup.exe");
    dicApps.Add("communicator.exe", "communicator.exe");
    dicApps.Add("msosync.exe", "msosync.exe");
    dicApps.Add("OneNoteM.exe", "OneNoteM.exe");
    dicApps.Add("iexplore.exe", "iexplore.exe");
    dicApps.Add("mavinject32.exe", "mavinject32.exe");
    dicApps.Add("werfault.exe", "werfault.exe");
    dicApps.Add("perfboost.exe", "perfboost.exe");
    dicApps.Add("roamingoffice.exe", "roamingoffice.exe");
    // SP1 additions / changes
    dicApps.Add("officeclicktorun.exe", "officeclicktorun.exe");
    dicApps.Add("officeondemand.exe", "officeondemand.exe");
    dicApps.Add("OfficeC2RClient.exe", "OfficeC2RClient.exe");
}

//-------------------------------------------------------------------------------
//   InitObjects
//
//   Initialize global objects
//-------------------------------------------------------------------------------
public void InitObjects()
{
    oWmiLocal = new System.Management.ManagementScope(@"\\.\root\cimv2");
    oWShell = new System.Diagnostics.Process();
    oShellApp = new object(); // Shell.Application - may need COM interop
    oFso = new object(); // FileSystemObject - may need COM interop or use System.IO
    oMsi = new object(); // WindowsInstaller.Installer - may need COM interop
    oReg = new System.Management.ManagementScope(@"\\.\root\default:StdRegProv");
}

//-------------------------------------------------------------------------------
//   FreeObjects
//
//   Free initialized global objects
//-------------------------------------------------------------------------------
public void FreeObjects()
{
    if (oWmiLocal is IDisposable disposable1) disposable1.Dispose();
    if (oWShell is IDisposable disposable2) disposable2.Dispose();
    oShellApp = null;
    oFso = null;
    oMsi = null;
    if (oReg is IDisposable disposable3) disposable3.Dispose();
}

//-------------------------------------------------------------------------------
//   ParseCmdLine
//
//   Command line parser
//-------------------------------------------------------------------------------
public void ParseCmdLine()
{
    string[] arrArguments;
    string sArguments = "";
    string sArg0;

    var args = Environment.GetCommandLineArgs();
    int iArgCnt = args.Length - 1; // Subtract 1 because args[0] is the executable path
    if (iArgCnt > 0)
    {
        if (args[1] == "UAC")
        {
            if (args.Length == 2) iArgCnt = 0;
        }
    }
    if (iArgCnt == 0)
    {
        var scriptName = Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        switch (scriptName.ToUpper())
        {
            default:
                //Create the log
                CreateLog();
                FindInstalledOProducts();
                sDefault = "ALL";
                arrArguments = sDefault.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (arrArguments.Length == 0) arrArguments = new string[0];
                break;
        }
    }
    else
    {
        arrArguments = new string[iArgCnt];
        for (int iCnt = 0; iCnt < iArgCnt; iCnt++)
        {
            arrArguments[iCnt] = args[iCnt + 1].ToUpper();
            sArguments = sArguments + arrArguments[iCnt] + " ";
        }
    }

    // hardcode to full removal
    sArg0 = "ALL";

    switch (sArg0.ToUpper())
    {
        case "?":
            ShowSyntax();
            break;
        case "ALL":
            fRemoveAll = true;
            fRemoveOse = false;
            break;
        case "C2R":
            fC2R = true;
            fRemoveAll = false;
            fRemoveOse = false;
            break;
        default:
            fRemoveAll = false;
            fRemoveOse = false;
            sSkuRemoveList = sArg0;
            break;
    }

    for (int iCnt = 0; iCnt < arrArguments.Length; iCnt++)
    {
        switch (arrArguments[iCnt])
        {
            case "?":
            case "/?":
            case "-?":
                ShowSyntax();
                break;

            case "/DETECTONLY":
            case "/PREVIEW":
                fDetectOnly = true;
                break;

            case "/FORCEARPUNINSTALL":
                fForceArpUninstall = true;
                break;

            case "/KL":
            case "/KEEPLICENSE":
                fKeepLicense = true;
                break;

            case "/L":
            case "/LOG":
                fLogInitialized = false;
                if (arrArguments.Length > iCnt + 1)
                {
                    if (Directory.Exists(arrArguments[iCnt + 1]))
                    {
                        sLogDir = arrArguments[iCnt + 1];
                    }
                    else
                    {
                        try
                        {
                            Directory.CreateDirectory(arrArguments[iCnt + 1]);
                            sLogDir = arrArguments[iCnt + 1];
                        }
                        catch
                        {
                            sLogDir = sScrubDir;
                        }
                    }
                }
                break;

            case "/N":
            case "/NOCANCEL":
                fNoCancel = true;
                break;

            case "/NE":
            case "/NOELEVATE":
                fNoElevate = true;
                break;

            case "/OFFLINE":
            case "/FORCEOFFLINE":
                fOffline = true;
                break;

            case "/O":
            case "/OSE":
                fRemoveOse = true;
                break;

            case "/Q":
            case "/QUIET":
                fQuiet = true;
                break;

            case "/RETERRORSUCCESS":
            case "/RETURNERRORORSUCCESS":
            case "/REOS":
                fReturnErrorOrSuccess = true;
                break;

            case "/S":
            case "/SKIPSD":
            case "/SKIPSHORTCUTDETECTION":
                fSkipSD = true;
                break;

            // for test purposes only!
            case "/TR":
            case "/TESTRERUN":
                fTestRerun = true;
                break;
        }
    }
    if (!fLogInitialized) CreateLog();
    LogH2($"Arguments: {sArguments}\r\n");
}

//-------------------------------------------------------------------------------
//   ShowSyntax
//
//   Show the expected syntax for the script usage
//-------------------------------------------------------------------------------
public void ShowSyntax()
{
    Console.WriteLine($"\r\n{Constants.SCRIPTFILE} V {Constants.SCRIPTVERSION}\r\n" +
                     "Copyright (c) Microsoft Corporation. All Rights Reserved\r\n\r\n" +
                     $"{Constants.SCRIPTFILE} - Remove {Constants.ONAME}\r\n" +
                     "when a regular uninstall is no longer possible\r\n\r\n" +
                     $"Usage:\t{Constants.SCRIPTFILE}\r\n\r\n" +
                     "\t/?                          ' Displays this help\r\n" +
                     "\t/Log [LogfolderPath]        ' Custom folder for log files\r\n" +
                     "\t/SkipSD                     ' Skips the ShortcutDetection in local profiles\r\n" +
                     "\t/NoCancel                   ' Setup.exe and Msiexec.exe have no Cancel button\r\n" +
                     "\t/Quiet                      ' Script, Setup.exe and Msiexec.exe run quiet with no UI\r\n" +
                     "\t/ReturnErrorOrSuccess        ' Returns 0 for a successful removal. Non-Zero if not.\r\n");
    Environment.Exit(0);
}

//-------------------------------------------------------------------------------
//   FindInstalledOProducts
//
//   Office configuration products are listed with their configuration product
//   name in the "Uninstall" key.
//-------------------------------------------------------------------------------
public void FindInstalledOProducts()
{
    string ArpItem, prod, cult;
    string sCurKey, sValue, sConfigName, sCulture, sDisplayVersion, sVersionFallback;
    string sUninstallString, sProd;
    int iLeft, iRight;
    string[] arrKeys, arrProducts, arrCultures;
    bool fSystemComponent0, fDisplayVersion, fUninstallString;

    const string REG_ARP = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\";
    const string REG_O15RPROPERTYBAG = @"SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag\";
    const string REG_O15C2RCONFIGURATION = @"SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration\";
    const string REG_O15C2RPRODUCTIDS = @"SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\";
    const string REG_O16C2RCONFIGURATION = @"SOFTWARE\Microsoft\Office\16.0\ClickToRun\Configuration\";
    const string REG_O16C2RPRODUCTIDS = @"SOFTWARE\Microsoft\Office\16.0\ClickToRun\ProductReleaseIDs\Active\";
    const string REG_C2RCONFIGURATION = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration\";
    const string REG_C2RPRODUCTIDS = @"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\";

    if (dicInstalledSku.Count > 0) return; //Already done from command line parser

    fDisplayVersion = false;

    // identify C2R products
    LogH1("Detect installed products ");

    LogOnly("Check for O15 C2R products");
    // Check O15 Configuration key
    if (RegReadValue(Constants.HKLM, REG_O15C2RCONFIGURATION, "ProductReleaseIds", out sValue, "REG_SZ"))
    {
        arrProducts = sValue.Split(',');
        if (RegReadValue(Constants.HKLM, REG_O15C2RPRODUCTIDS + "culture", "x-none", out sVersionFallback, "REG_SZ"))
        {
            fDisplayVersion = true;
            // get version from active with fallback on configuration
            foreach (string prodItem in arrProducts)
            {
                prod = prodItem;
                LogOnly($"Found O15 C2R product in Configuration: {prod}");
                // update product dictionary
                if (!dicInstalledSku.ContainsKey(prod.ToLower()))
                {
                    LogOnly($"add new product to dictionary: {prod.ToLower()}");
                    dicInstalledSku.Add(prod.ToLower(), sVersionFallback);
                }
            }
        }
    }

    // Check O15 PropertyBag key
    if (RegReadValue(Constants.HKLM, REG_O15RPROPERTYBAG, "productreleaseid", out sValue, "REG_SZ"))
    {
        arrProducts = sValue.Split(',');
        if (RegReadValue(Constants.HKLM, REG_O15C2RPRODUCTIDS + "culture", "x-none", out sVersionFallback, "REG_SZ"))
        {
            fDisplayVersion = true;
            foreach (string prodItem in arrProducts)
            {
                prod = prodItem;
                LogOnly($"Found O15 C2R product in PropertyBag: {prod}");
                // update product dictionary
                if (!dicInstalledSku.ContainsKey(prod.ToLower()))
                {
                    LogOnly($"add new product to dictionary: {prod.ToLower()}");
                    dicInstalledSku.Add(prod.ToLower(), sVersionFallback);
                }
            }
        }
    }

    //O16 section
    LogOnly("Check for Office C2R products (>=QR8)");
    // Check Office Configuration key
    if (RegReadValue(Constants.HKLM, REG_C2RPRODUCTIDS, "ActiveConfiguration", out sActiveConfiguration, "REG_SZ"))
    {
        // Get DisplayVersion
        //Try QR8 logic first
        fDisplayVersion = RegReadValue(Constants.HKLM, REG_C2RPRODUCTIDS + sActiveConfiguration + "\\culture", "x-none", out sVersionFallback, "REG_SZ");
        if (RegEnumKey(Constants.HKLM, REG_C2RPRODUCTIDS + sActiveConfiguration + "\\culture", out arrCultures))
        {
            foreach (string cultItem in arrCultures)
            {
                cult = cultItem;
                if (cult.ToLower().Contains("x-none"))
                {
                    fDisplayVersion = RegReadValue(Constants.HKLM, REG_C2RPRODUCTIDS + sActiveConfiguration + "\\culture\\" + cult, "Version", out sVersionFallback, "REG_SZ");
                }
            }
        }
        // Update product dic
        if (RegEnumKey(Constants.HKLM, REG_C2RPRODUCTIDS + sActiveConfiguration, out arrProducts))
        {
            foreach (string prodItem in arrProducts)
            {
                prod = prodItem;
                sProd = prod.ToLower();
                int dotIndex = sProd.IndexOf(".");
                if (dotIndex > 0) sProd = sProd.Substring(0, dotIndex);
                switch (sProd)
                {
                    case "culture":
                    case "stream":
                        break;
                    default:
                        LogOnly($"Found Office C2R product in Configuration: {prod}");
                        // update product dictionary
                        if (!dicInstalledSku.ContainsKey(sProd))
                        {
                            LogOnly($"add new product to dictionary: {sProd}");
                            if (RegReadValue(Constants.HKLM, REG_C2RPRODUCTIDS + sActiveConfiguration + "\\" + prod + "\\x-none", "Version", out sDisplayVersion, "REG_SZ"))
                            {
                                dicInstalledSku.Add(sProd, sDisplayVersion);
                            }
                            else
                            {
                                dicInstalledSku.Add(sProd, sVersionFallback);
                            }
                        }
                        break;
                }
            }
        }
    }

    LogOnly("Check for Office C2R products (QR7)");
    // Check Office Configuration key
    if (RegReadValue(Constants.HKLM, REG_C2RCONFIGURATION, "ProductReleaseIds", out sValue, "REG_SZ"))
    {
        arrProducts = sValue.Split(',');
        if (!fDisplayVersion) fDisplayVersion = RegReadValue(Constants.HKLM, REG_C2RPRODUCTIDS + "Active\\culture", "x-none", out sVersionFallback, "REG_SZ");
        if (fDisplayVersion)
        {
            foreach (string prodItem in arrProducts)
            {
                prod = prodItem;
                LogOnly($"Found Office C2R product in Configuration: {prod}");
                // update version tracking
                if (!dicInstalledSku.ContainsKey(prod.ToLower()))
                {
                    LogOnly($"add new product to dictionary: {prod.ToLower()}");
                    dicInstalledSku.Add(prod.ToLower(), sVersionFallback);
                }
            }
        }
    }

    LogOnly("Check for O16 C2R products (QR6)");
    // Check O16 Configuration key
    if (RegReadValue(Constants.HKLM, REG_O16C2RCONFIGURATION, "ProductReleaseIds", out sValue, "REG_SZ"))
    {
        arrProducts = sValue.Split(',');
        if (!fDisplayVersion) fDisplayVersion = RegReadValue(Constants.HKLM, REG_O16C2RPRODUCTIDS + "culture", "x-none", out sVersionFallback, "REG_SZ");
        if (fDisplayVersion)
        {
            foreach (string prodItem in arrProducts)
            {
                prod = prodItem;
                LogOnly($"Found O16 (QR6) C2R product in Configuration: {prod}");
                // update product dictionary
                if (!dicInstalledSku.ContainsKey(prod.ToLower()))
                {
                    LogOnly($"add new product to dictionary: {prod}");
                    dicInstalledSku.Add(prod.ToLower(), sVersionFallback);
                }
            }
        }
    }

    LogOnly("Check ARP for Office C2R products");
    // ARP
    if (RegEnumKey(Constants.HKLM, REG_ARP, out arrKeys) && arrKeys != null)
    {
        foreach (string arpItem in arrKeys)
        {
            ArpItem = arpItem;
            // filter on Office C2R products
            sCurKey = REG_ARP + ArpItem + "\\";
            fUninstallString = RegReadValue(Constants.HKLM, sCurKey, "UninstallString", out sValue, "REG_SZ");
            if (fUninstallString && (sValue.ToUpper().Contains("MICROSOFT OFFICE 1") || sValue.ToUpper().Contains("OFFICECLICKTORUN.EXE")))
            {
                //cache UninstallString for uninstall Fallback
                if (sValue.ToUpper().Contains("OFFICECLICKTORUN.EXE")) sArpUninstallCmd = sValue;
                //get Version
                fDisplayVersion = RegReadValue(Constants.HKLM, sCurKey, "DisplayVersion", out sDisplayVersion, "REG_SZ");
                //extract the productreleaseid
                sValue = sValue.Trim();
                int lastSpace = sValue.LastIndexOf(" ");
                prod = lastSpace >= 0 ? sValue.Substring(lastSpace).Trim() : sValue;
                prod = prod.Replace("productstoremove=", "");
                int underscoreIndex = prod.IndexOf("_");
                if (underscoreIndex > 0) prod = prod.Substring(0, underscoreIndex);
                int dot1Index = prod.IndexOf(".1");
                if (dot1Index > 0) prod = prod.Substring(0, dot1Index);
                LogOnly($"Found C2R product in ARP: {prod}");
                if (!dicInstalledSku.ContainsKey(prod.ToLower()))
                {
                    LogOnly($"add new product to dictionary: {prod}");
                    dicInstalledSku.Add(prod.ToLower(), sDisplayVersion);
                }
                // categorize the SKU as C2R
                if (!dicC2RSuite.ContainsKey(ArpItem)) dicC2RSuite.Add(ArpItem, $"{prod} - {sDisplayVersion}");
            }
            else
            {
                //Legacy logic keep for compat reasons
                sValue = "";
                sDisplayVersion = "";
                string sValueTemp;
                bool hasSystemComponent = RegReadValue(Constants.HKLM, sCurKey, "SystemComponent", out sValueTemp, "REG_DWORD");
                fSystemComponent0 = !(hasSystemComponent && sValueTemp == "1");
                fDisplayVersion = RegReadValue(Constants.HKLM, sCurKey, "DisplayVersion", out sValue, "REG_SZ");
                if (fDisplayVersion)
                {
                    sDisplayVersion = sValue;
                    if (sValue.Length > 1)
                    {
                        try
                        {
                            if (int.TryParse(sValue.Substring(0, 2), out int versionNum))
                            {
                                fDisplayVersion = (versionNum > 14);
                            }
                        }
                        catch { }
                    }
                    else
                    {
                        fDisplayVersion = false;
                    }
                }
                fUninstallString = RegReadValue(Constants.HKLM, sCurKey, "UninstallString", out sUninstallString, "REG_SZ");

                // filter on C2R configuration SKU
                if (fUninstallString && (sUninstallString.ToUpper().Contains("MICROSOFT OFFICE 1") || sUninstallString.ToUpper().Contains("OFFICECLICKTORUN.EXE")))
                {
                    // Extract the ProductReleaseID
                    if (sUninstallString.Contains("productstoremove="))
                    {
                        int lastSpaceIdx = sUninstallString.LastIndexOf(" ");
                        sConfigName = lastSpaceIdx >= 0 ? sUninstallString.Substring(lastSpaceIdx).Trim() : sUninstallString;
                        sConfigName = sConfigName.Replace("productstoremove=", "");
                        int underscoreIdx = sConfigName.IndexOf("_");
                        if (underscoreIdx > 0) sConfigName = sConfigName.Substring(0, underscoreIdx);
                    }
                    else
                    {
                        iLeft = ArpItem.IndexOf(" - ") + 2;
                        iRight = ArpItem.IndexOf(" - ", iLeft);
                        if (iRight > 0)
                        {
                            sConfigName = ArpItem.Substring(iLeft, iRight - iLeft).Trim();
                            sCulture = ArpItem.Substring(iRight + 3);
                        }
                        else
                        {
                            sConfigName = ArpItem.Substring(0, iLeft - 3).Trim();
                            sCulture = ArpItem.Substring(iLeft);
                        }
                        sConfigName = sConfigName.Replace("Microsoft", "").Replace("Office", "")
                            .Replace("Professional", "Pro").Replace("Standard", "Std")
                            .Replace("(Technical Preview)", "").Replace("15", "").Replace("16", "")
                            .Replace("2013", "").Replace("2016", "").Replace(" ", "")
                            .Replace("Project", "Prj").Replace("Visio", "Vis");
                    }
                    if (!dicInstalledSku.ContainsKey(sConfigName.ToLower()))
                    {
                        LogOnly($"add new product to dictionary (ARP Legacy): {sConfigName}");
                        dicInstalledSku.Add(sConfigName.ToLower(), sDisplayVersion);
                    }
                    // categorize the SKU as C2R
                    if (!dicC2RSuite.ContainsKey(ArpItem)) dicC2RSuite.Add(ArpItem, $"{sConfigName} - {sDisplayVersion}");
                }
                else if (fDisplayVersion && (ArpItem.ToUpper().Contains("OFFICE15.") || ArpItem.ToUpper().Contains("OFFICE16.")))
                {
                    // classic .msi install SKU
                    iLeft = ArpItem.IndexOf(".") + 1;
                    iRight = ArpItem.IndexOf("-", iLeft);
                    sConfigName = iRight > 0 ? ArpItem.Substring(iLeft, iRight - iLeft) : ArpItem.Substring(iLeft);
                    sCulture = "";
                    if (!dicKeepSku.ContainsKey(ArpItem)) dicKeepSku.Add(ArpItem, $"{sConfigName} - {sDisplayVersion}");
                }

                // Other products
                if (InScope(ArpItem))
                {
                    string midValue = ArpItem.Length > 14 ? ArpItem.Substring(10, 4) : "";
                    switch (midValue)
                    {
                        case "007E":
                        case "008F":
                        case "008C":
                        case "00DD":
                            sConfigName = "Habanero";
                            RegReadValue(Constants.HKLM, sCurKey, "DisplayName", out sConfigName, "REG_SZ");
                            if (!dicInstalledSku.ContainsKey(ArpItem.ToLower()))
                            {
                                LogOnly($"add new product to dictionary (ARP Integraton Components): {ArpItem}");
                                dicInstalledSku.Add(ArpItem.ToLower(), sDisplayVersion);
                            }
                            if (!dicC2RSuite.ContainsKey(ArpItem)) dicC2RSuite.Add(ArpItem, $"{sConfigName} - {sDisplayVersion}");
                            break;
                        case "24E1":
                        case "237A":
                            sConfigName = "MSOIDLOGIN";
                            if (!dicInstalledSku.ContainsKey(ArpItem.ToLower()))
                            {
                                LogOnly($"add new product to dictionary (ARP MSOIDLogin): {ArpItem}");
                                dicInstalledSku.Add(ArpItem.ToLower(), sDisplayVersion);
                            }
                            if (!dicC2RSuite.ContainsKey(ArpItem)) dicC2RSuite.Add(ArpItem, $"{sConfigName} - {sDisplayVersion}");
                            break;
                        default:
                            if (!dicInstalledSku.ContainsKey(ArpItem.ToLower()))
                            {
                                LogOnly($"add new product to dictionary (ARP other): {ArpItem}");
                                dicInstalledSku.Add(ArpItem.ToLower(), sDisplayVersion);
                            }
                            break;
                    }
                }
                // End legacy logic
            }
        }
    }
}

//-------------------------------------------------------------------------------
//   EnsureValidWIMetadata
//
//   Ensures that only valid metadata entries exist to avoid API failures.
//   Invalid entries will be removed
//-------------------------------------------------------------------------------
public void EnsureValidWIMetadata(int hDefKey, string sKey, int iValidLength)
{
    string[] arrKeys;

    if (sKey.Length > 1)
    {
        if (sKey.EndsWith("\\")) sKey = sKey.Substring(0, sKey.Length - 1);
    }

    if (RegEnumKey(hDefKey, sKey, out arrKeys))
    {
        foreach (string subKey in arrKeys)
        {
            if (subKey.Length != iValidLength)
            {
                RegDeleteKey(hDefKey, $"{sKey}\\{subKey}\\");
            }
        }
    }
}

//-------------------------------------------------------------------------------
//   CleanOSPP
//
//   Clean out licenses from the Office Software Protection Platform
//-------------------------------------------------------------------------------
public void CleanOSPP()
{
    const string OfficeAppId = "0ff1ce15-a989-479d-af46-f275c6370663";  //Office 2013

    string sCleanOSPP = "x64\\CleanOSPP.exe";
    if (!f64) sCleanOSPP = "x86\\CleanOSPP.exe";
    if (File.Exists(sScriptDir + sCleanOSPP))
    {
        string sCmd = sScriptDir + sCleanOSPP;
        Log($"   Running: {sCmd}");
        try
        {
            if (!fDetectOnly)
            {
                var process = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = sCmd,
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
                process?.WaitForExit();
                int sRetVal = process?.ExitCode ?? 0;
                Log($"   Return value: {sRetVal}");
            }
        }
        catch { }
        return;
    }

    try
    {
        if (!(dicC2RSuite.Count > 0 || dicKeepSku.Count > 0))
        {
            Log("Skip CleanOSPP");
            return;
        }

        // Initialize the software protection platform object with a filter on Office 2013 products
        var scope = new System.Management.ManagementScope(@"\\.\root\cimv2");
        string queryString;
        if (iVersionNT > 601)
        {
            queryString = $"SELECT ID, ApplicationId, PartialProductKey, Name, ProductKeyID FROM SoftwareLicensingProduct WHERE ApplicationId = '{OfficeAppId}' AND PartialProductKey <> NULL";
        }
        else
        {
            queryString = $"SELECT ID, ApplicationId, PartialProductKey, Name, ProductKeyID FROM OfficeSoftwareProtectionProduct WHERE ApplicationId = '{OfficeAppId}' AND PartialProductKey <> NULL";
        }
        var query = new System.Management.ObjectQuery(queryString);
        using (var searcher = new System.Management.ManagementObjectSearcher(scope, query))
        {
            // Remove all licenses
            foreach (System.Management.ManagementObject pi in searcher.Get())
            {
                if (pi != null)
                {
                    try
                    {
                        pi.InvokeMethod("UninstallProductKey", new object[] { pi["ProductKeyID"] });
                    }
                    catch { }
                }
            }
        }
    }
    catch { }
}

//-------------------------------------------------------------------------------
//   ClearVNextLicCache
//
//   clear local license cache for vNext
//-------------------------------------------------------------------------------
public void ClearVNextLicCache()
{
    string sLocalAppData = ExpandEnv("%localappdata%");
    DeleteFolder($"{sLocalAppData}\\Microsoft\\Office\\Licenses");
}


//-------------------------------------------------------------------------------
//   DelSchtasks
//
//   Delete known scheduled tasks.
//-------------------------------------------------------------------------------
public void DelSchtasks()
{
    if ((iError & Constants.ERROR_USERCANCEL) != 0) return;

    LogH1("Remove scheduled tasks");

    LogOnly("FF_INTEGRATEDstreamSchedule");
    RunCommand("SCHTASKS /Delete /TN FF_INTEGRATEDstreamSchedule /F", false);
    System.Threading.Thread.Sleep(500);

    LogOnly("FF_INTEGRATEDUPDATEDETECTION");
    RunCommand("SCHTASKS /Delete /TN FF_INTEGRATEDUPDATEDETECTION /F", false);
    System.Threading.Thread.Sleep(500);

    LogOnly("C2RAppVLoggingStart");
    RunCommand("SCHTASKS /Delete /TN C2RAppVLoggingStart /F", false);
    System.Threading.Thread.Sleep(500);

    LogOnly("Office 15 Subscription Heartbeat");
    string sCmd = "SCHTASKS /Delete /TN \"" + "Office 15 Subscription Heartbeat" + "\" /F";
    RunCommand(sCmd, false);
    System.Threading.Thread.Sleep(500);

    LogOnly("Microsoft Office 15 Sync Maintenance");
    sCmd = "SCHTASKS /Delete /TN \"" + "Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}" + "\" /F";
    RunCommand(sCmd, false);
    System.Threading.Thread.Sleep(500);

    LogOnly("OfficeInventoryAgentFallBack");
    sCmd = "SCHTASKS /Delete /TN \"" + "\\Microsoft\\Office\\OfficeInventoryAgentFallBack" + "\" /F";
    RunCommand(sCmd, false);
    System.Threading.Thread.Sleep(500);

    LogOnly("OfficeTelemetryAgentFallBack");
    sCmd = "SCHTASKS /Delete /TN \"" + "\\Microsoft\\Office\\OfficeTelemetryAgentFallBack" + "\" /F";
    RunCommand(sCmd, false);
    System.Threading.Thread.Sleep(500);

    LogOnly("OfficeInventoryAgentLogOn");
    sCmd = "SCHTASKS /Delete /TN \"" + "\\Microsoft\\Office\\OfficeInventoryAgentLogOn" + "\" /F";
    RunCommand(sCmd, false);

    LogOnly("OfficeTelemetryAgentLogOn");
    sCmd = "SCHTASKS /Delete /TN \"" + "\\Microsoft\\Office\\OfficeTelemetryAgentLogOn" + "\" /F";
    RunCommand(sCmd, false);

    LogOnly("Office Background Streaming");
    sCmd = "SCHTASKS /Delete /TN \"" + "Office Background Streaming" + "\" /F";
    RunCommand(sCmd, false);
    System.Threading.Thread.Sleep(500);

    LogOnly("Office Automatic Updates");
    sCmd = "SCHTASKS /Delete /TN \"" + "\\Microsoft\\Office\\Office Automatic Updates" + "\" /F";
    RunCommand(sCmd, false);
    System.Threading.Thread.Sleep(500);

    LogOnly("Office ClickToRun Service Monitor");
    sCmd = "SCHTASKS /Delete /TN \"" + "\\Microsoft\\Office\\Office ClickToRun Service Monitor" + "\" /F";
    RunCommand(sCmd, false);
    System.Threading.Thread.Sleep(500);

    LogOnly("Office Subscription Maintenance");
    sCmd = "SCHTASKS /Delete /TN \"" + "Office Subscription Maintenance" + "\" /F";
    RunCommand(sCmd, false);
    System.Threading.Thread.Sleep(500);
}

//-------------------------------------------------------------------------------
//   CloseOfficeApps
//
//   End all running instances of applications that will be removed.
//-------------------------------------------------------------------------------
public void CloseOfficeApps()
{
    bool fWait = false;
    iProcCloseCnt = iProcCloseCnt + 1;
    if (fRerun) return;

    try
    {
        if (!fUserConsent)
        {
            // detect processes to allow a user warning
            string sUserWarn = "Please save all open documents and close all Office, IE and Windows Explorer applications before proceeding.\r\n" +
                              "When you click OK this removal process will terminate all running Office, IE and Windows Explorer processes and applications.\r\n\r\n" +
                              "Click \"Cancel\" to end this removal now.";
            var scope = new System.Management.ManagementScope(@"\\.\root\cimv2");

            foreach (string app in dicApps.Keys)
            {
                string sAppName = app.Replace(".", "%.");
                var query = new System.Management.ObjectQuery($"SELECT * FROM Win32_Process WHERE Name LIKE '{sAppName}'");
                using (var searcher = new System.Management.ManagementObjectSearcher(scope, query))
                {
                    foreach (System.Management.ManagementObject process in searcher.Get())
                    {
                        string processName = process["Name"]?.ToString() ?? "";
                        if (!sUserWarn.Contains(processName))
                        {
                            sUserWarn += "\r\n - " + processName;
                        }
                    }
                }
            }

            var query2 = new System.Management.ObjectQuery("SELECT * FROM Win32_Process");
            using (var searcher = new System.Management.ManagementObjectSearcher(scope, query2))
            {
                foreach (System.Management.ManagementObject process in searcher.Get())
                {
                    string executablePath = process["ExecutablePath"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(executablePath) && IsC2R(executablePath))
                    {
                        string processName = process["Name"]?.ToString() ?? "";
                        if (!sUserWarn.Contains(processName))
                        {
                            sUserWarn += "\r\n - " + processName;
                        }
                    }
                }
            }

            if (sUserWarn.Contains(" - ") && !fQuiet)
            {
                var result = MessageBox.Show(
                    sUserWarn,
                    "Save your unsaved work now!",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Warning
                );
                if (result == DialogResult.Cancel)
                {
                    SetError(Constants.ERROR_USERCANCEL);
                    ExitScript();
                    return;
                }
                else
                {
                    fUserConsent = true;
                }
            }
        }

        // end known processes first
        scope = new System.Management.ManagementScope(@"\\.\root\cimv2");
        foreach (string app in dicApps.Keys)
        {
            string sAppName = app.Replace(".", "%.");
            var query = new System.Management.ObjectQuery($"SELECT * FROM Win32_Process WHERE Name LIKE '{sAppName}'");
            using (var searcher = new System.Management.ManagementObjectSearcher(scope, query))
            {
                foreach (System.Management.ManagementObject process in searcher.Get())
                {
                    string processName = process["Name"]?.ToString() ?? "";
                    try
                    {
                        process.InvokeMethod("Terminate", null);
                        uint iRet = 0;
                        Log($"End process '{processName}' returned: {iRet}");
                        fWait = true;
                    }
                    catch (Exception ex)
                    {
                        Log($"Error terminating process {processName}: {ex.Message}");
                    }
                }
            }
        }

        // end running applications
        var query3 = new System.Management.ObjectQuery("SELECT * FROM Win32_Process");
        using (var searcher = new System.Management.ManagementObjectSearcher(scope, query3))
        {
            foreach (System.Management.ManagementObject process in searcher.Get())
            {
                string executablePath = process["ExecutablePath"]?.ToString() ?? "";
                if (!string.IsNullOrEmpty(executablePath) && IsC2R(executablePath))
                {
                    string processName = process["Name"]?.ToString() ?? "";
                    try
                    {
                        process.InvokeMethod("Terminate", null);
                        uint iRet = 0;
                        Log($"End process '{processName}' returned: {iRet}");
                        fWait = true;
                    }
                    catch (Exception ex)
                    {
                        Log($"Error terminating process {processName}: {ex.Message}");
                    }
                }
            }
        }

        if (fWait) System.Threading.Thread.Sleep(5000);
    }
    catch (Exception ex)
    {
        Log($"Error in CloseOfficeApps: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   Uninstall
//
//   Identify and invoke default uninstall command for a regular uninstall.
//-------------------------------------------------------------------------------
public void Uninstall()
{
    if ((iError & Constants.ERROR_USERCANCEL) != 0) return;

    try
    {
        int hDefKey;
        string sSubKeyName, sValue, Name, sPkgFld, sPkgGuid, sUninstallCmd, sMsiProp, sCmd;
        string[] arrNames, arrTypes;
        int sReturn;

        // check if OSE service is *installed, *not disabled, *running under System context.
        LogH2("Check state of OSE service");
        var scope = new System.Management.ManagementScope(@"\\.\root\cimv2");
        var query = new System.Management.ObjectQuery("SELECT * FROM Win32_Service WHERE Name LIKE 'ose%'");
        using (var searcher = new System.Management.ManagementObjectSearcher(scope, query))
        {
            foreach (System.Management.ManagementObject srvc in searcher.Get())
            {
                try
                {
                    if (srvc["StartMode"]?.ToString() == "Disabled")
                    {
                        uint result = (uint)srvc.InvokeMethod("ChangeStartMode", new object[] { "Manual" });
                        if (result != 0)
                        {
                            Log("Conflict detected: OSE service is disabled");
                        }
                    }
                    if (srvc["StartName"]?.ToString() != "LocalSystem")
                    {
                        uint result = (uint)srvc.InvokeMethod("Change", new object[] { null, null, null, null, null, null, "LocalSystem", "" });
                        if (result == 0)
                        {
                            Log("Conflict detected: OSE service not running as LocalSystem");
                        }
                    }
                }
                catch { }
            }
        }

        if (dicC2RSuite.Count == 0)
        {
            Log("No uninstallable C2R items registered in Uninstall");
        }

        // call odt based uninstall
        UninstallOfficeC2R();

        // remove the published component registration for C2R packages
        LogH2("Remove published component registration for C2R packages");
        // delete the manifest files
        for (int i = 1; i <= 4; i++)
        {
            switch (i)
            {
                case 1:
                    RegReadValue(Constants.HKLM, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun", "PackageFolder", out sPkgFld, "REG_SZ");
                    RegReadValue(Constants.HKLM, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun", "PackageGUID", out sPkgGuid, "REG_SZ");
                    break;
                case 2:
                    RegReadValue(Constants.HKLM, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun", "PackageFolder", out sPkgFld, "REG_SZ");
                    RegReadValue(Constants.HKLM, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun", "PackageGUID", out sPkgGuid, "REG_SZ");
                    break;
                case 3:
                    RegReadValue(Constants.HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun", "PackageFolder", out sPkgFld, "REG_SZ");
                    RegReadValue(Constants.HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun", "PackageGUID", out sPkgGuid, "REG_SZ");
                    break;
                case 4:
                    sPkgFld = sPackageFolder;
                    sPkgGuid = sPackageGuid;
                    break;
            }
            if (!string.IsNullOrEmpty(sPkgFld) && Directory.Exists(Path.Combine(sPkgFld, "root", "Integration")))
            {
                sCmd = $"cmd.exe /c del \"{sPkgFld}\\root\\Integration\\C2RManifest*.xml\"";
                Log($"   Run: {sCmd}");
                if (!fDetectOnly) sReturn = RunCommand(sCmd, true);
                Log($"   Return value: {sReturn}");
                string integratorPath = Path.Combine(sPkgFld, "root", "Integration", "integrator.exe");
                if (File.Exists(integratorPath))
                {
                    sCmd = $"\"{integratorPath}\" /U  /Extension PackageRoot=\"{sPkgFld}\\root\" PackageGUID={sPkgGuid}";
                    Log($"   Run: {sCmd}");
                    if (!fDetectOnly) sReturn = RunCommand(sCmd, true);
                    Log($"   Return value: {sReturn}");
                    sCmd = $"\"{integratorPath}\" /U";
                    Log($"   Run: {sCmd}");
                    if (!fDetectOnly) sReturn = RunCommand(sCmd, true);
                    Log($"   Return value: {sReturn}");
                }
                string programDataIntegrator = Path.Combine(sProgramData, "Microsoft", "ClickToRun", $"{{{sPkgGuid}}}", "integrator.exe");
                if (File.Exists(programDataIntegrator))
                {
                    sCmd = $"\"{programDataIntegrator}\" /U  /Extension PackageRoot=\"{sPkgFld}\\root\" PackageGUID={sPkgGuid}";
                    Log($"   Run: {sCmd}");
                    if (!fDetectOnly) sReturn = RunCommand(sCmd, true);
                    Log($"   Return value: {sReturn}");
                }
            }
        }

        // delete potential blocking registry keys for msiexec based tasks
        LogH2("Remove C2R and App-V registry data");
        foreach (string sku in dicC2RSuite.Keys)
        {
            // remove the ARP entry
            RegDeleteKey(Constants.HKLM, Constants.REG_ARP + sku);
        }
        RegDeleteKey(Constants.HKCU, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun");
        RegDeleteKey(Constants.HKCU, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun");
        RegDeleteKey(Constants.HKCU, @"SOFTWARE\Microsoft\Office\ClickToRun");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun");

        // AppV keys
        hDefKey = Constants.HKCU;
        sSubKeyName = @"SOFTWARE\Microsoft\AppV\ISV";
        do
        {
            LogOnly($"Scanning key: {sSubKeyName}");
            if (RegEnumValues(hDefKey, sSubKeyName, out arrNames, out arrTypes))
            {
                foreach (string nameItem in arrNames)
                {
                    Name = nameItem;
                    if (IsC2R(Name)) RegDeleteValue(hDefKey, sSubKeyName, Name, false);
                }
            }
            if (hDefKey == Constants.HKLM) break;
            hDefKey = Constants.HKLM;
        } while (true);

        hDefKey = Constants.HKCU;
        sSubKeyName = @"SOFTWARE\Microsoft\AppVISV";
        do
        {
            LogOnly($"Scanning key: {sSubKeyName}");
            if (RegEnumValues(hDefKey, sSubKeyName, out arrNames, out arrTypes))
            {
                foreach (string nameItem in arrNames)
                {
                    Name = nameItem;
                    if (IsC2R(Name)) RegDeleteValue(hDefKey, sSubKeyName, Name, false);
                }
            }
            if (hDefKey == Constants.HKLM) break;
            hDefKey = Constants.HKLM;
        } while (true);

        // msiexec based uninstall
        sMsiProp = " REBOOT=ReallySuppress NOREMOVESPAWN=True";
        LogH2("Detect Msi based products");
        // Note: oMsi.Products requires COM interop for WindowsInstaller.Installer
        // For now, we'll enumerate MSI products from registry
        try
        {
            string[] arrKeys;
            if (RegEnumKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products", out arrKeys))
            {
                foreach (string prod in arrKeys)
                {
                    if (CheckDelete(prod))
                    {
                        Log($"Call msiexec.exe to remove {prod}");
                        sUninstallCmd = $"msiexec.exe /x{prod}{sMsiProp}";
                        if (fQuiet)
                        {
                            sUninstallCmd += " /q";
                        }
                        else
                        {
                            sUninstallCmd += " /qb-!";
                        }
                        sUninstallCmd += $" /l*v \"{Path.Combine(sLogDir, $"Uninstall_{prod}.log")}\"";
                        CloseOfficeApps();
                        LogOnly($"Call msiexec with '{sUninstallCmd}'");
                        if (!fDetectOnly)
                        {
                            sReturn = RunCommand(sUninstallCmd, true);
                            Log($"msiexec returned: {SetupRetVal(sReturn)} ({sReturn})\r\n");
                            fRebootRequired = fRebootRequired || (sReturn == 3010);
                            if (fRebootRequired) SetError(Constants.ERROR_REBOOT_REQUIRED);
                            switch (sReturn)
                            {
                                case Constants.ERROR_SUCCESS:
                                case Constants.ERROR_SUCCESS_CONFIG_COMPLETE:
                                case Constants.ERROR_SUCCESS_REBOOT_REQUIRED:
                                    //success no action required
                                    break;
                                default:
                                    SetError(Constants.ERROR_STAGE1);
                                    break;
                            }
                        }
                    }
                    else
                    {
                        LogOnly($"Skip out of scope product: {prod}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Log($"Error enumerating MSI products: {ex.Message}");
        }

        if (!fDetectOnly) RunCommand("cmd.exe /c net stop msiserver", false);
    }
    catch (Exception ex)
    {
        Log($"Error in Uninstall: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   BuildRemoveXml
//
//-------------------------------------------------------------------------------
public void BuildRemoveXml()
{
    try
    {
        LogOnly("BuildRemoveXml");
        string sConfigRemoveAllXml;
        if (fQuiet)
        {
            sConfigRemoveAllXml = "<Configuration>\r\n" +
                                 "  <Remove All=\"TRUE\" />\r\n" +
                                 "  <Display Level=\"None\" />\r\n" +
                                 "</Configuration>";
        }
        else
        {
            sConfigRemoveAllXml = "<Configuration>\r\n" +
                                 "  <Remove All=\"TRUE\" />\r\n" +
                                 "</Configuration>";
        }

        // write out the config.xml
        string configPath = Path.Combine(sScrubDir, "RemoveAll.xml");
        File.WriteAllText(configPath, sConfigRemoveAllXml);
        LogOnly($"RemoveAll.xml:\r\n{sConfigRemoveAllXml}");
    }
    catch (Exception ex)
    {
        Log($"Error building RemoveAll.xml: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   HttpDownloadFile
//
//   Copy a file from a url to a local path using HttpClient
//-------------------------------------------------------------------------------
public bool HttpDownloadFile(string sUrl, string sLocalPath)
{
    try
    {
        Log($"Download {sUrl} to {sLocalPath}");

        using (var httpClient = new System.Net.Http.HttpClient())
        {
            var response = httpClient.GetAsync(sUrl).Result;
            if (response.IsSuccessStatusCode)
            {
                var fileBytes = response.Content.ReadAsByteArrayAsync().Result;
                File.WriteAllBytes(sLocalPath, fileBytes);
            }
            else
            {
                Log($"Download failed with status code: {response.StatusCode}");
                return false;
            }
        }

        bool fileExists = File.Exists(sLocalPath);
        Log($"Check download success. {sLocalPath} exists: {fileExists}");
        return fileExists;
    }
    catch (Exception ex)
    {
        Log($"Error downloading file: {ex.Message}");
        return false;
    }
}


//-------------------------------------------------------------------------------
//   UninstallOfficeC2R
//
//   Uninstall all of Office C2R through ODT
//-------------------------------------------------------------------------------
public void UninstallOfficeC2R()
{
    try
    {
        string sCmd, sODTFullPath = "", sKey, sFolder, sLeft, sRight;
        int iVerODT = 0;
        bool fCanUseOdtUninstall = false;

        if (RegValExists(Constants.HKLM, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\culture", "x-none"))
        {
            iVerODT = 15;
        }
        if (RegValExists(Constants.HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\Active\culture", "x-none") ||
            RegValExists(Constants.HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs", "ActiveConfiguration"))
        {
            iVerODT = 16;
        }

        if (RegValExists(Constants.HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\Active\culture", "x-none") ||
            RegValExists(Constants.HKLM, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\culture", "x-none") ||
            RegValExists(Constants.HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs", "ActiveConfiguration"))
        {
            LogH($"ODT Uninstall C2R {iVerODT}.0");

            if (!fForceArpUninstall)
            {
                //build the remove.xml
                BuildRemoveXml();

                //verify ODT is available
                string odtPath = Path.Combine(sScriptDir, $"ODT{iVerODT}", "setup.exe");
                if (File.Exists(odtPath))
                {
                    sODTFullPath = odtPath;
                    fCanUseOdtUninstall = true;
                }
                else
                {
                    //ODT not available. Try to download
                    if (!fOffline)
                    {
                        if (iVerODT == 15)
                        {
                            string downloadPath = Path.Combine(sScrubDir, "officedeploymenttool.exe");
                            if (HttpDownloadFile("https://download.microsoft.com/download/6/2/3/6230F7A2-D8A9-478B-AC5C-57091B632FCF/officedeploymenttool_x86_5031-1000.exe", downloadPath))
                            {
                                //Extract
                                sCmd = $"\"{downloadPath}\" /quiet /extract:\"{sScrubDir}\"";
                                Log($"Run silent extract: {sCmd}");
                                if (!fDetectOnly)
                                {
                                    int sReturn = RunCommand(sCmd, true);
                                }
                                sODTFullPath = Path.Combine(sScrubDir, "setup.exe");
                                if (File.Exists(sODTFullPath)) fCanUseOdtUninstall = true;
                            }
                        }
                        else
                        {
                            string setupPath = Path.Combine(sScrubDir, "setup.exe");
                            if (HttpDownloadFile("http://officecdn.microsoft.com/pr/wsus/setup.exe", setupPath))
                            {
                                sODTFullPath = setupPath;
                                fCanUseOdtUninstall = true;
                            }
                        }
                    }
                }

                Log($"Can use ODT based uninstall: {fCanUseOdtUninstall}");

                if (fCanUseOdtUninstall)
                {
                    //build uninstall command
                    string xmlPath = Path.Combine(sScrubDir, "RemoveAll.xml");
                    sCmd = $"\"{sODTFullPath}\" /configure \"{xmlPath}\"";
                    Log($"run uninstall: {sCmd}");
                    if (!fDetectOnly)
                    {
                        int sReturn = RunCommand(sCmd, true);
                        Log($"ODT uninstall for OfficeC2R returned with value: {sReturn}");
                    }
                }
                else
                {
                    //Can't use ODT for uninstall attempt. Use unified ARP uninstall command
                    if (!string.IsNullOrEmpty(sArpUninstallCmd))
                    {
                        sArpUninstallCmd = sArpUninstallCmd.Trim();
                        int productIndex = sArpUninstallCmd.IndexOf(" productstoremove=");
                        if (productIndex > 0)
                        {
                            sLeft = sArpUninstallCmd.Substring(0, productIndex);
                            sRight = sArpUninstallCmd.Substring(productIndex);
                            int spaceIndex = sRight.IndexOf(" ");
                            if (spaceIndex > 0) sRight = sRight.Substring(spaceIndex);
                            else sRight = "";
                        }
                        else
                        {
                            sLeft = sArpUninstallCmd;
                            sRight = "";
                        }
                        sCmd = $"{sLeft}productstoremove=\"AllProducts\"{sRight}";
                        if (fQuiet) sCmd += " displaylevel=\"false\"";
                        Log($"run uninstall: {sCmd}");
                        if (!fDetectOnly)
                        {
                            int sReturn = RunCommand(sCmd, true);
                            Log($"ARP uninstall for OfficeC2R returned with value: {sReturn}");
                        }
                    }
                }
            }
            else
            {
                Log("Skip ODT switch is active");
            }
        }
        else
        {
            Log("Uninstall for OfficeC2R not required");
        }

        //Log uninstall success
        Log("Log uninstall success");

        sKey = @"SOFTWARE\Microsoft\Office\ClickToRun";
        Log($"HKLM\\{sKey} still exists: {RegKeyExists(Constants.HKLM, sKey)}");

        sFolder = $"\"{ExpandEnv("%programfiles%")}\\Microsoft Office\\root\"";
        Log($"{sFolder} still exists: {Directory.Exists(ExpandEnv("%programfiles%") + "\\Microsoft Office\\root")}");

        if (f64)
        {
            sFolder = $"\"{ExpandEnv("%programfiles(x86)%")}\\Microsoft Office\\root\"";
            Log($"{sFolder} exists: {Directory.Exists(ExpandEnv("%programfiles(x86)%") + "\\Microsoft Office\\root")}");
        }
    }
    catch (Exception ex)
    {
        Log($"Error in UninstallOfficeC2R: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   RunCommand
//
//   Helper method to run a command and return exit code
//-------------------------------------------------------------------------------
private int RunCommand(string command, bool waitForExit)
{
    try
    {
        ProcessStartInfo processInfo;
        
        // Check if command already starts with cmd.exe
        if (command.StartsWith("cmd.exe", StringComparison.OrdinalIgnoreCase))
        {
            // Command already includes cmd.exe, parse it
            var parts = command.Split(new[] { ' ' }, 2, StringSplitOptions.RemoveEmptyEntries);
            processInfo = new ProcessStartInfo
            {
                FileName = parts[0],
                Arguments = parts.Length > 1 ? parts[1] : "",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
        }
        else
        {
            // Wrap in cmd.exe /c
            processInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/c " + command,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
        }

        using (var process = Process.Start(processInfo))
        {
            if (waitForExit && process != null)
            {
                process.WaitForExit();
                return process.ExitCode;
            }
            return 0;
        }
    }
    catch (Exception ex)
    {
        Log($"Error running command '{command}': {ex.Message}");
        return -1;
    }
}

//-------------------------------------------------------------------------------
//   RegWipe
//
//   Removal of left behind registry data
//-------------------------------------------------------------------------------
public void Regwipe()
{
    if ((iError & Constants.ERROR_USERCANCEL) != 0) return;

    try
    {
        int hDefKey;
        string item, name, value, sGuid, sSubKeyName, sValue;
        int i;
        string[] arrKeys, arrNames, arrTypes, arrTestNames, arrTestTypes;
        string[] arrMultiSzValues, arrMultiSzNewValues;
        bool fDelReg;

        LogH1("Registry CleanUp");

        //Moved to earlier timing to avoid reboot needs
        //if (dicKeepSku.Count == 0) ClearShellIntegrationReg();

        CloseOfficeApps();

        // Note: ARP entries have already been cleared in uninstall stage

        // HKCU Registration
        RegDeleteKey(Constants.HKCU, @"Software\Microsoft\Office\15.0\Registration");
        RegDeleteKey(Constants.HKCU, @"Software\Microsoft\Office\16.0\Registration");
        RegDeleteKey(Constants.HKCU, @"Software\Microsoft\Office\Registration");

        // C2R specifics
        // AppV key "SOFTWARE\Microsoft\AppV" has already been cleared in uninstall stage

        // Virtual InstallRoot
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Virtual");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Virtual");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\Common\InstallRoot\Virtual");

        // Mapi Search reg
        //O15
        if (dicKeepSku.Count == 0) RegDeleteKey(Constants.HKLM, @"SOFTWARE\Classes\CLSID\{2027FC3B-CF9D-4ec7-A823-38BA308625CC}");
        //O16
        //{F8E61EDD-EA25-484e-AC8A-7447F2AAE2A9}

        // C2R keys
        RegDeleteKey(Constants.HKCU, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\15.0\ClickToRun");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\15.0\ClickToRunStore");
        RegDeleteKey(Constants.HKCU, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\16.0\ClickToRun");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\16.0\ClickToRunStore");
        RegDeleteKey(Constants.HKCU, @"SOFTWARE\Microsoft\Office\ClickToRun");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\ClickToRun");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Office\ClickToRunStore");

        // Office key in HKLM
        if (dicKeepSku.Count == 0)
        {
            //double calls to ensure Wow6432 gets cleared out as well
            RegDeleteKey(Constants.HKLM, @"Software\Microsoft\Office\15.0");
            RegDeleteKey(Constants.HKLM, @"Software\Microsoft\Office\15.0");
            RegDeleteKey(Constants.HKLM, @"Software\Microsoft\Office\16.0");
            RegDeleteKey(Constants.HKLM, @"Software\Microsoft\Office\16.0");
        }
        ClearOfficeHKLM(@"SOFTWARE\Microsoft\Office");

        // Run key
        sSubKeyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        if (RegEnumValues(Constants.HKLM, sSubKeyName, out arrNames, out arrTypes))
        {
            foreach (string nameItem in arrNames)
            {
                name = nameItem;
                if (RegReadValue(Constants.HKLM, sSubKeyName, name, out sValue, "REG_SZ"))
                {
                    if (IsC2R(sValue)) RegDeleteValue(Constants.HKLM, sSubKeyName, name, false);
                }
            }
        }
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Lync15", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Lync16", false);

        // ARP
        // Note: configuration entries have already been removed
        // as part of the 'Uninstall' stage
        if (RegEnumKey(Constants.HKLM, Constants.REG_ARP, out arrKeys))
        {
            foreach (string itemKey in arrKeys)
            {
                item = itemKey;
                if (item.Length > 37)
                {
                    sGuid = item.Substring(0, 38).ToUpper();
                    if (CheckDelete(sGuid)) RegDeleteKey(Constants.HKLM, Constants.REG_ARP + item + "\\");
                }
            }
        }

        // UpgradeCodes, WI config, WI global config
        LogH2("Scan Windows Installer metadata for removeable UpgradeCodes");
        for (int iLoopCnt = 1; iLoopCnt <= 5; iLoopCnt++)
        {
            switch (iLoopCnt)
            {
                case 1:
                    sSubKeyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\";
                    hDefKey = Constants.HKLM;
                    break;
                case 2:
                    sSubKeyName = @"Installer\UpgradeCodes\";
                    hDefKey = Constants.HKCR;
                    break;
                case 3:
                    sSubKeyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\";
                    hDefKey = Constants.HKLM;
                    break;
                case 4:
                    sSubKeyName = @"Installer\Features\";
                    hDefKey = Constants.HKCR;
                    break;
                case 5:
                    sSubKeyName = @"Installer\Products\";
                    hDefKey = Constants.HKCR;
                    break;
            }
            if (RegEnumKey(hDefKey, sSubKeyName, out arrKeys))
            {
                foreach (string itemKey in arrKeys)
                {
                    item = itemKey;
                    // ensure the expected length for a compressed GUID
                    if (item.Length == 32)
                    {
                        // expand the GUID
                        sGuid = GetExpandedGuid(item);
                        // check if it's an Office key
                        if (CheckDelete(sGuid))
                        {
                            if (iLoopCnt < 3)
                            {
                                // enum all entries
                                if (RegEnumValues(hDefKey, sSubKeyName + item, out arrNames, out arrTypes))
                                {
                                    if (arrNames != null)
                                    {
                                        // delete entries within removal scope
                                        foreach (string nameItem in arrNames)
                                        {
                                            name = nameItem;
                                            if (name.Length == 32)
                                            {
                                                sGuid = GetExpandedGuid(name);
                                                if (CheckDelete(sGuid)) RegDeleteValue(hDefKey, sSubKeyName + item + "\\", name, true);
                                            }
                                            else
                                            {
                                                // invalid data -> delete the value
                                                RegDeleteValue(hDefKey, sSubKeyName + item + "\\", name, true);
                                            }
                                        }
                                    }
                                    // if all entries were removed - delete the key
                                    if (!RegEnumValues(hDefKey, sSubKeyName + item, out arrTestNames, out arrTestTypes))
                                    {
                                        RegDeleteKey(hDefKey, sSubKeyName + item + "\\");
                                    }
                                }
                            }
                            else //iLoopCnt >= 3
                            {
                                RegDeleteKey(hDefKey, sSubKeyName + item + "\\");
                            }
                        }
                    }
                }
            }
        }

        // Components in Global
        LogH2("Scan Windows Installer Global Components metadata");
        sSubKeyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\";
        hDefKey = Constants.HKLM;
        if (RegEnumKey(hDefKey, sSubKeyName, out arrKeys))
        {
            foreach (string itemKey in arrKeys)
            {
                item = itemKey;
                // ensure the expected length for a compressed GUID
                if (item.Length == 32)
                {
                    if (RegEnumValues(hDefKey, sSubKeyName + item, out arrNames, out arrTypes))
                    {
                        foreach (string nameItem in arrNames)
                        {
                            name = nameItem;
                            if (name.Length == 32)
                            {
                                sGuid = GetExpandedGuid(name);
                                if (CheckDelete(sGuid))
                                {
                                    RegDeleteValue(hDefKey, sSubKeyName + item + "\\", name, false);
                                    // if all entries were removed - delete the key
                                    if (!RegEnumValues(hDefKey, sSubKeyName + item, out arrTestNames, out arrTestTypes))
                                    {
                                        RegDeleteKey(hDefKey, sSubKeyName + item + "\\");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // Published Components
        LogH2("Scanning Windows Installer Published Components metadata");
        sSubKeyName = @"Installer\Components\";
        hDefKey = Constants.HKCR;
        if (RegEnumKey(hDefKey, sSubKeyName, out arrKeys))
        {
            foreach (string itemKey in arrKeys)
            {
                item = itemKey;
                // ensure the expected length for a compressed GUID
                if (item.Length == 32)
                {
                    if (RegEnumValues(hDefKey, sSubKeyName + item, out arrNames, out arrTypes))
                    {
                        foreach (string nameItem in arrNames)
                        {
                            name = nameItem;
                            if (RegReadValue(hDefKey, sSubKeyName + item, name, out sValue, "REG_MULTI_SZ"))
                            {
                                arrMultiSzValues = sValue.Split('\r');
                                if (arrMultiSzValues != null)
                                {
                                    var newValuesList = new List<string>();
                                    fDelReg = false;
                                    foreach (string valueItem in arrMultiSzValues)
                                    {
                                        value = valueItem;
                                        if (value.Length > 19)
                                        {
                                            sGuid = "";
                                            if (GetDecodedGuid(value.Substring(0, Math.Min(Constants.SQUISHED, value.Length)), out sGuid))
                                            {
                                                if (CheckDelete(sGuid))
                                                {
                                                    fDelReg = true;
                                                }
                                                else
                                                {
                                                    newValuesList.Add(value);
                                                }
                                            }
                                        }
                                    }
                                    if (newValuesList.Count > 0)
                                    {
                                        arrMultiSzNewValues = newValuesList.ToArray();
                                        if (arrMultiSzValues.Length != newValuesList.Count)
                                        {
                                            // Update registry with new values - requires COM interop or registry API
                                            // For now, just delete if all should be removed
                                            if (newValuesList.Count == 0)
                                            {
                                                RegDeleteValue(hDefKey, sSubKeyName + item + "\\", name, true);
                                                if (!RegEnumValues(hDefKey, sSubKeyName + item, out arrTestNames, out arrTestTypes))
                                                {
                                                    RegDeleteKey(hDefKey, sSubKeyName + item + "\\");
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (fDelReg)
                                        {
                                            RegDeleteValue(hDefKey, sSubKeyName + item + "\\", name, true);
                                            // if all entries were removed - delete the key
                                            if (!RegEnumValues(hDefKey, sSubKeyName + item, out arrTestNames, out arrTestTypes))
                                            {
                                                RegDeleteKey(hDefKey, sSubKeyName + item + "\\");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // Known Typelib Registration
        LogH2("Scanning known Office TypeLibs registration");
        RegWipeTypeLib();
    }
    catch (Exception ex)
    {
        Log($"Error in Regwipe: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   ClearShellIntegrationReg
//
//   Delete registry items that may cause Explorer / Windows Shell to have a lock
//   on files
//-------------------------------------------------------------------------------
public void ClearShellIntegrationReg()
{
    try
    {
        // Protocol Handlers
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Classes\Protocols\Handler\osf");

        // Context Menu Handlers
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Classes\CLSID\{573FFD05-2805-47C2-BCE0-5F19512BEB8D}");
        //RegDeleteKey(Constants.HKLM, @"SOFTWARE\Classes\CLSID\{4693FF15-B962-420A-9E5D-176F7D4B8321}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Classes\CLSID\{8BA85C75-763B-4103-94EB-9470F12FE0F7}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Classes\CLSID\{CD55129A-B1A1-438E-A425-CEBC7DC684EE}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Classes\CLSID\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Classes\CLSID\{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}");

        // Groove ShellIconOverlayIdentifiers
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)");

        // Shell extensions
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{B28AA736-876B-46DA-B3A8-84C5E30BA492}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{8B02D659-EBBB-43D7-9BBA-52CF22C5B025}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{42042206-2D85-11D3-8CFF-005004838597}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{506F4668-F13E-4AA1-BB04-B43203AB3CC0}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{D66DC78C-4F61-447F-942B-3FB6980118CF}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{46137B78-0EC3-426D-8B89-FF7C3A458B5E}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{8BA85C75-763B-4103-94EB-9470F12FE0F7}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{CD55129A-B1A1-438E-A425-CEBC7DC684EE}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}", false);
        RegDeleteValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}", false);

        // BHO
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}");

        // OneNote Namespace Extension for Desktop
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}");

        // Web Sites
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\Namespace\{B28AA736-876B-46DA-B3A8-84C5E30BA492}");
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\NetworkNeighborhood\Namespace\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}");

        // VolumeCaches
        RegDeleteKey(Constants.HKLM, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Microsoft Office Temp Files");

        RestoreExplorer();
        FreeObjects();
        System.Threading.Thread.Sleep(500);
        InitObjects();
    }
    catch (Exception ex)
    {
        Log($"Error in ClearShellIntegrationReg: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   RegWipeTypeLib
//
//   Clear out left behind Typelib registrations
//-------------------------------------------------------------------------------
//Clean out known typelib registration
public void RegWipeTypeLib()
{
    try
    {
        string sKey, sTLKey, sTLVerKey;
        string sTypeLibs, tl, k, sValue, sFilePath;
        string[] arrTypeLibs, arrKeys, arrKeys2;
        bool fClearTL, fCanDelete;

        sTypeLibs = "{000204EF-0000-0000-C000-000000000046};{000204EF-0000-0000-C000-000000000046};{00020802-0000-0000-C000-000000000046};{00020813-0000-0000-C000-000000000046};{00020905-0000-0000-C000-000000000046};{0002123C-0000-0000-C000-000000000046};{00024517-0000-0000-C000-000000000046};{0002E157-0000-0000-C000-000000000046};{00062FFF-0000-0000-C000-000000000046};{0006F062-0000-0000-C000-000000000046};{0006F080-0000-0000-C000-000000000046};{012F24C1-35B0-11D0-BF2D-0000E8D0D146};{06CA6721-CB57-449E-8097-E65B9F543A1A};{07B06096-5687-4D13-9E32-12B4259C9813};{0A2F2FC4-26E1-457B-83EC-671B8FC4C86D};{0AF7F3BE-8EA9-4816-889E-3ED22871FE05};{0D452EE1-E08F-101A-852E-02608C4D0BB4};{0EA692EE-BB50-4E3C-AEF0-356D91732725};{1F8E79BA-9268-4889-ADF3-6D2AABB3C32C};{2374F0B1-3220-4c71-B702-AF799F31ABB4};{238AA1AC-786F-4C17-BAAB-253670B449B9};{28DD2950-2D4A-42B5-ABBF-500AA42E7EC1};{2A59CA0A-4F1B-44DF-A216-CB2C831E5870};{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52};{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52};{2F7FC181-292B-11D2-A795-DFAA798E9148};{3120BA9F-4FC8-4A4F-AE1E-02114F421D0A};{31411197-A502-11D2-BBCA-00C04F8EC294};{3B514091-5A69-4650-87A3-607C4004C8F2};{47730B06-C23C-4FCA-8E86-42A6A1BC74F4};{49C40DDF-1B04-4868-B3B5-E49F120E4BFA};{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28};{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07};{4D95030A-A3A9-4C38-ACA8-D323A2267698};{55A108B0-73BB-43db-8C03-1BEF4E3D2FE4};{56D04F5D-964F-4DBF-8D23-B97989E53418};{5B87B6F0-17C8-11D0-AD41-00A0C90DC8D9};{66CDD37F-D313-4E81-8C31-4198F3E42C3C};{6911FD67-B842-4E78-80C3-2D48597C2ED0};{698BB59C-38F1-4CEF-92F9-7E3986E708D3};{6DDCE504-C0DC-4398-8BDB-11545AAA33EF};{6EFF1177-6974-4ED1-99AB-82905F931B87};{73720002-33A0-11E4-9B9A-00155D152105};{759EF423-2E8F-4200-ADF0-5B6177224BEE};{76F6F3F5-9937-11D2-93BB-00105A994D2C};{773F1B9A-35B9-4E95-83A0-A210F2DE3B37};{7D868ACD-1A5D-4A47-A247-F39741353012};{7E36E7CB-14FB-4F9E-B597-693CE6305ADC};{831FDD16-0C5C-11D2-A9FC-0000F8754DA1};{8404DD0E-7A27-4399-B1D9-6492B7DD7F7F};{8405D0DF-9FDD-4829-AEAD-8E2B0A18FEA4};{859D8CF5-7ADE-4DAB-8F7D-AF171643B934};{8E47F3A2-81A4-468E-A401-E1DEBBAE2D8D};{91493440-5A91-11CF-8700-00AA0060263B};{9A8120F2-2782-47DF-9B62-54F672075EA1};{9B7C3E2E-25D5-4898-9D85-71CEA8B2B6DD};{9B92EB61-CBC1-11D3-8C2D-00A0CC37B591};{9D58B963-654A-4625-86AC-345062F53232};{9DCE1FC0-58D3-471B-B069-653CE02DCE88};{A4D51C5D-F8BF-46CC-92CC-2B34D2D89716};{A717753E-C3A6-4650-9F60-472EB56A7061};{AA53E405-C36D-478A-BBFF-F359DF962E6D};{AAB9C2AA-6036-4AE1-A41C-A40AB7F39520};{AB54A09E-1604-4438-9AC7-04BE3E6B0320};{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4};{AC2DE821-36A2-11CF-8053-00AA006009FA};{B30CDC65-4456-4FAA-93E3-F8A79E21891C};{B8812619-BDB3-11D0-B19E-00A0C91E29D8};{B9164592-D558-4EE7-8B41-F1C9F66D683A};{B9AA1F11-F480-4054-A84E-B5D9277E40A8};{BA35B84E-A623-471B-8B09-6D72DD072F25};{BDEADE33-C265-11D0-BCED-00A0C90AB50F};{BDEADEF0-C265-11D0-BCED-00A0C90AB50F};{BDEADEF0-C265-11D0-BCED-00A0C90AB50F};{C04E4E5E-89E6-43C0-92BD-D3F2C7FBA5C4};{C3D19104-7A67-4EB0-B459-D5B2E734D430};{C78F486B-F679-4af5-9166-4E4D7EA1CEFC};{CA973FCA-E9C3-4B24-B864-7218FC1DA7BA};{CBA4EBC4-0C04-468d-9F69-EF3FEED03236};{CBBC4772-C9A4-4FE8-B34B-5EFBD68F8E27};{CD2194AA-11BE-4EFD-97A6-74C39C6508FF};{E0B12BAE-FC67-446C-AAE8-4FA1F00153A7};{E985809A-84A6-4F35-86D6-9B52119AB9D7};{ECD5307E-4419-43CF-8BDA-C9946AC375CF};{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C};{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C};{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C};{EDDCFF16-3AEE-4883-BD91-0F3978640DFB};{EE9CFA8C-F997-4221-BE2F-85A5F603218F};{F2A7EE29-8BF6-4a6d-83F1-098E366C709C};{F3685D71-1FC6-4CBD-B244-E60D8C89990B}";
        arrTypeLibs = sTypeLibs.Split(';');
        sTLKey = @"Software\Classes\TypeLib\";

        //iterate all known typelibs
        foreach (string tlItem in arrTypeLibs)
        {
            tl = tlItem;
            fClearTL = false;
            sKey = sTLKey + tl;
            if (RegKeyExists(Constants.HKLM, sKey))
            {
                //enum subkeys
                LogOnly($"Found registration for typelib {tl}");
                if (RegEnumKey(Constants.HKLM, sKey, out arrKeys))
                {
                    foreach (string kItem in arrKeys)
                    {
                        k = kItem;
                        sTLVerKey = sKey + "\\" + k;
                        fCanDelete = RegEnumKey(Constants.HKLM, sTLVerKey, out arrKeys2);
                        if (RegReadValue(Constants.HKLM, sTLVerKey + "\\0\\Win32\\", "", out sValue, "REG_SZ"))
                        {
                            LogOnly($"Found key HKLM\\{sTLVerKey}\\0\\Win32\\");
                            //get the safe filepath
                            int lastDot = sValue.LastIndexOf(".");
                            sFilePath = lastDot >= 0 ? sValue.Substring(0, Math.Min(lastDot + 3, sValue.Length)) : sValue;
                            LogOnly($"Found filepath: {sValue} - using filepath: {sFilePath}");
                            if (File.Exists(sFilePath))
                            {
                                fCanDelete = false;
                                fClearTL = false;
                                LogOnly("File target still in use. TypeLib registration will persisted.");
                            }
                            else
                            {
                                fClearTL = fCanDelete;
                                LogOnly("File target not found. Flagging for delete");
                            }
                        }
                        if (RegReadValue(Constants.HKLM, sTLVerKey + "\\9\\Win32\\", "", out sValue, "REG_SZ"))
                        {
                            LogOnly($"Found key HKLM\\{sTLVerKey}\\9\\Win32\\");
                            //get the safe filepath
                            int lastDot = sValue.LastIndexOf(".");
                            sFilePath = lastDot >= 0 ? sValue.Substring(0, Math.Min(lastDot + 3, sValue.Length)) : sValue;
                            LogOnly($"Found filepath: {sValue} - using filepath: {sFilePath}");
                            if (File.Exists(sFilePath))
                            {
                                fCanDelete = false;
                                fClearTL = false;
                                LogOnly("File target still in use. TypeLib registration will persisted.");
                            }
                            else
                            {
                                fClearTL = fCanDelete;
                                LogOnly("File target not found. Flagging for delete");
                            }
                        }
                        if (RegReadValue(Constants.HKLM, sTLVerKey + "\\0\\Win64\\", "", out sValue, "REG_SZ"))
                        {
                            LogOnly($"Found key HKLM\\{sTLVerKey}\\0\\Win64\\");
                            //get the safe filepath
                            int lastDot = sValue.LastIndexOf(".");
                            sFilePath = lastDot >= 0 ? sValue.Substring(0, Math.Min(lastDot + 3, sValue.Length)) : sValue;
                            LogOnly($"Found filepath: {sValue} - using filepath: {sFilePath}");
                            if (File.Exists(sFilePath))
                            {
                                fCanDelete = false;
                                fClearTL = false;
                                LogOnly("File target still in use. TypeLib registration will persisted.");
                            }
                            else
                            {
                                fClearTL = fCanDelete;
                                LogOnly("File target not found. Flagging for delete");
                            }
                        }
                        if (RegReadValue(Constants.HKLM, sTLVerKey + "\\9\\Win64\\", "", out sValue, "REG_SZ"))
                        {
                            LogOnly($"Found key HKLM\\{sTLVerKey}\\9\\Win64\\");
                            //get the safe filepath
                            int lastDot = sValue.LastIndexOf(".");
                            sFilePath = lastDot >= 0 ? sValue.Substring(0, Math.Min(lastDot + 3, sValue.Length)) : sValue;
                            LogOnly($"Found filepath: {sValue} - using filepath: {sFilePath}");
                            if (File.Exists(sFilePath))
                            {
                                fCanDelete = false;
                                fClearTL = false;
                                LogOnly("File target still in use. TypeLib registration will persisted.");
                            }
                            else
                            {
                                fClearTL = fCanDelete;
                                LogOnly("File target not found. Flagging for delete");
                            }
                        }
                        //remove the key if no valid usage references were found
                        if (fCanDelete)
                        {
                            LogOnly($"Removing version registration: HKLM\\{sTLVerKey}");
                            RegDeleteKey(Constants.HKLM, sTLVerKey);
                        }
                    }
                }
                //Re-evaluate if there are subkeys left to determine if the whole typelib reg should be removed
                if (!RegEnumKey(Constants.HKLM, sKey, out arrKeys))
                {
                    LogOnly("TypeLib registration obsolete - removing registration key");
                    RegDeleteKey(Constants.HKLM, sKey);
                }
            }
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RegWipeTypeLib: {ex.Message}");
    }
}


//-------------------------------------------------------------------------------
//   FileWipe
//
//   Removal of left behind services, files and shortcuts
//-------------------------------------------------------------------------------
public void FileWipe()
{
    if ((iError & Constants.ERROR_USERCANCEL) != 0) return;

    try
    {
        bool fDelFolders;

        LogH1("File Cleanup");

        fDelFolders = false;
        CloseOfficeApps();
        DelSchtasks();

        LogH1("Delete Services");
        // remove the OfficeSvc service
        LogH2("Delete OfficeSvc service");
        DeleteService("OfficeSvc");

        // SP1 addition / change
        // remove the ClickToRunSvc service
        LogH2("Delete ClickToRunSvc service");
        DeleteService("ClickToRunSvc");

        // adding additional processes for termination
        if (!dicApps.ContainsKey("explorer.exe")) dicApps.Add("explorer.exe", "explorer.exe");
        if (!dicApps.ContainsKey("msiexec.exe")) dicApps.Add("msiexec.exe", "msiexec.exe");
        if (!dicApps.ContainsKey("ose.exe")) dicApps.Add("ose.exe", "ose.exe");

        if (fC2R)
        {
            LogH1("Delete Files and Folders");
            // delete C2R package files
            LogH2("Delete C2R package files");
            if (FolderExists(sProgramFiles + "\\Microsoft Office 15") ||
                FolderExists(sProgramFiles + "\\Microsoft Office 16") ||
                FolderExists(ExpandEnv("%programfiles%") + "\\Microsoft Office\\PackageManifests") ||
                FolderExists(ExpandEnv("%programfiles(x86)%") + "\\Microsoft Office\\PackageManifests"))
            {
                fDelFolders = true;
                Log("   Attention: Now closing Explorer.exe for file delete operations");
                Log("   Explorer will automatically restart.");
                System.Threading.Thread.Sleep(2000);
                CloseOfficeApps();
            }
            // delete Office folders
            LogH2("Delete Office folders");
            DeleteFolder(sProgramFiles + "\\Microsoft Office 15");
            DeleteFolder(sProgramFiles + "\\Microsoft Office 16");
            if (f64)
            {
                DeleteFolder(sCommonProgramFilesX86 + "\\Microsoft Office 15");
                DeleteFolder(sCommonProgramFilesX86 + "\\Microsoft Office 16");
            }
            if (fDelFolders)
            {
                DeleteFolder(sProgramFiles + "\\Microsoft Office\\PackageManifests");
                DeleteFolder(sProgramFiles + "\\Microsoft Office\\PackageSunrisePolicies");
                DeleteFolder(sProgramFiles + "\\Microsoft Office\\root");
                DeleteFile(sProgramFiles + "\\Microsoft Office\\AppXManifest.xml");
                DeleteFile(sProgramFiles + "\\Microsoft Office\\FileSystemMetadata.xml");
                if (dicKeepSku.Count == 0)
                {
                    DeleteFolder(sProgramFiles + "\\Microsoft Office\\Office16");
                    DeleteFolder(sProgramFiles + "\\Microsoft Office\\Office15");
                }
                if (f64)
                {
                    DeleteFolder(sProgramFilesX86 + "\\Microsoft Office\\PackageManifests");
                    DeleteFolder(sProgramFilesX86 + "\\Microsoft Office\\PackageSunrisePolicies");
                    DeleteFolder(sProgramFilesX86 + "\\Microsoft Office\\root");
                    DeleteFile(sProgramFilesX86 + "\\Microsoft Office\\AppXManifest.xml");
                    DeleteFile(sProgramFilesX86 + "\\Microsoft Office\\FileSystemMetadata.xml");
                    if (dicKeepSku.Count == 0)
                    {
                        DeleteFolder(sProgramFilesX86 + "\\Microsoft Office\\Office16");
                        DeleteFolder(sProgramFilesX86 + "\\Microsoft Office\\Office15");
                    }
                }
            }

            DeleteFolder(sProgramData + "\\Microsoft\\ClickToRun");
            DeleteFolder(sCommonProgramFiles + "\\microsoft shared\\ClickToRun");
            DeleteFolder(sProgramData + "\\Microsoft\\office\\FFPackageLocker");
            DeleteFolder(sProgramData + "\\Microsoft\\office\\ClickToRunPackageLocker");
            if (File.Exists(sProgramData + "\\Microsoft\\office\\FFPackageLocker")) DeleteFile(sProgramData + "\\Microsoft\\office\\FFPackageLocker");
            if (File.Exists(sProgramData + "\\Microsoft\\office\\FFStatePBLocker")) DeleteFile(sProgramData + "\\Microsoft\\office\\FFStatePBLocker");
            if (dicKeepSku.Count == 0) DeleteFolder(sProgramData + "\\Microsoft\\office\\Heartbeat");
            DeleteFolder(ExpandEnv("%userprofile%") + "\\Microsoft Office");
            DeleteFolder(ExpandEnv("%userprofile%") + "\\Microsoft Office 15");
            DeleteFolder(ExpandEnv("%userprofile%") + "\\Microsoft Office 16");
        }

        // restore explorer.exe if needed
        RestoreExplorer();

        // delete shortcuts
        LogH2("Search and delete shortcuts");
        CleanShortcuts(sAllusersProfile, true, false);
        CleanShortcuts(sProfilesDirectory, true, false);

        // delete empty folder structures
        if (dicDelFolder.Count > 0)
        {
            LogH2("Remove empty folders");
            DeleteEmptyFolders();
        }

        // add the collected files in use for delete on reboot
        if (dicDelInUse.Count > 0) ScheduleDeleteEx();

        LogH2("File Cleanup complete");
    }
    catch (Exception ex)
    {
        Log($"Error in FileWipe: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   CleanShortcuts
//
//   Recursively search all profile folders for Office shortcuts in scope
//-------------------------------------------------------------------------------
public void CleanShortcuts(string sFolder, bool fDelete, bool fUnPin)
{
    if (fSkipSD) return;

    try
    {
        DirectoryInfo oFolder = new DirectoryInfo(sFolder);
        // exclude system protected link folders
        if ((oFolder.Attributes & FileAttributes.System) != 0) return;

        foreach (DirectoryInfo fld in oFolder.GetDirectories())
        {
            try
            {
                CleanShortcuts(fld.FullName, fDelete, fUnPin);
            }
            catch (Exception ex)
            {
                CheckError($"CleanShortcuts: \t{sFolder}");
                Log($"Error processing folder {fld.FullName}: {ex.Message}");
            }
        }

        foreach (FileInfo file in oFolder.GetFiles())
        {
            try
            {
                string filePath = file.FullName;
                if (filePath.ToLower().EndsWith(".lnk") && !filePath.ToLower().Contains("recentplaces"))
                {
                    bool fDeleteSC = false;
                    LogOnly($" check file: {filePath}");
                    // Use IWshRuntimeLibrary for shortcut handling - requires COM reference
                    // For now, we'll use a simplified approach
                    try
                    {
                        // Note: This requires COM interop for IWshShortcut
                        // Using dynamic to handle COM object
                        dynamic sc = null; // Would need: oWShell.CreateShortcut(filePath)
                        string targetPath = ""; // Would get from sc.TargetPath
                        
                        // Simplified check - would need actual COM interop implementation
                        // For now, checking file existence and path patterns
                        if (!string.IsNullOrEmpty(targetPath))
                        {
                            if (targetPath.Contains("{"))
                            {
                                //Handle Windows Installer shortcuts
                                int braceIndex = targetPath.IndexOf("{");
                                if (targetPath.Length >= braceIndex + 38)
                                {
                                    string guid = targetPath.Substring(braceIndex, 38);
                                    if (CheckDelete(guid)) fDeleteSC = true;
                                }
                            }
                            else
                            {
                                //Handle regular shortcuts
                                if (IsC2R(targetPath)) fDeleteSC = true;
                                if (!File.Exists(targetPath))
                                {
                                    // Shortcut target does not exist
                                    if (IsC2R(targetPath))
                                    {
                                        LogOnly($"remove Office shortcut with non-existent target: {filePath} - {targetPath}");
                                        fDeleteSC = true;
                                    }
                                }
                            }
                        }
                        else
                        {
                            // If we can't read shortcut, check if path suggests Office
                            if (IsC2R(filePath)) fDeleteSC = true;
                        }

                        if (fDeleteSC)
                        {
                            if (!dicDelFolder.ContainsKey(sFolder)) dicDelFolder.Add(sFolder, sFolder);
                            if (fUnPin || fDelete)
                            {
                                if (!string.IsNullOrEmpty(targetPath) && File.Exists(targetPath))
                                {
                                    // Target exists, keep as is
                                }
                                else
                                {
                                    // Would need COM interop to set sc.TargetPath = sNotepad and sc.Save()
                                    LogOnly($"linking empty shortcut to Notepad.exe as target: {filePath} - {targetPath}");
                                }
                                //Invoke new instance to UnPin file
                                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                                string sCmdLine = $"\"{exePath}\" \"UNPINSC\" \"{filePath}\"";
                                LogOnly($"Invoke UnPin handler for shortcut: {filePath}");
                                LogOnly($"UnPin command: {sCmdLine}");
                                if (!fDetectOnly)
                                {
                                    int sReturn = RunCommand(sCmdLine, true);
                                    LogOnly($"UnPin returned with: {sReturn}");
                                }
                            }
                            if (fDelete) DeleteFile(filePath);
                            fDeleteSC = false;
                            fClearTaskBand = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        CheckError($"CleanShortcutsSC: \t{sFolder}");
                        Log($"Error processing shortcut {filePath}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Error processing file {file.FullName}: {ex.Message}");
            }
        }
    }
    catch (Exception ex)
    {
        Log($"Error in CleanShortcuts: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   UnPin
//
//   Unpins a shortcut from the taskbar or start menu
//-------------------------------------------------------------------------------
public void Unpin(string sFilePath)
{
    try
    {
        // Requires COM interop for Shell.Application
        // Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
        // dynamic oShellAppUnPin = Activator.CreateInstance(shellAppType);
        // FileInfo file = new FileInfo(sFilePath);
        // dynamic folder = oShellAppUnPin.NameSpace(file.DirectoryName);
        // dynamic fldItem = folder.ParseName(file.Name);
        
        // For now, this is a placeholder - requires COM interop implementation
        // The actual implementation would iterate through fldItem.Verbs and call DoIt() on matching verbs
        LogOnly($"UnPin called for: {sFilePath}");
        // Note: Full implementation requires COM interop for Shell.Application
    }
    catch (Exception ex)
    {
        Log($"Error in Unpin: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   ClearTaskBand
//
//   Clears contents from the users taskband to get rid of pinned items
//-------------------------------------------------------------------------------
public void ClearTaskBand()
{
    try
    {
        string sid;
        string sTaskBand, sHKUTaskBand;
        string[] arrSid;

        sTaskBand = @"Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband\";
        RegDeleteValue(Constants.HKCU, sTaskBand, "Favorites", false);
        RegDeleteValue(Constants.HKCU, sTaskBand, "FavoritesRemovedChanges", false);
        RegDeleteValue(Constants.HKCU, sTaskBand, "FavoritesChanges", false);
        RegDeleteValue(Constants.HKCU, sTaskBand, "FavoritesResolve", false);
        RegDeleteValue(Constants.HKCU, sTaskBand, "FavoritesVersion", false);

        // enum all profiles in HKU
        LoadUsersReg();
        if (RegEnumKey(Constants.HKU, "", out arrSid))
        {
            foreach (string sidItem in arrSid)
            {
                sid = sidItem;
                sHKUTaskBand = sid + "\\" + sTaskBand;
                RegDeleteValue(Constants.HKCU, sHKUTaskBand, "Favorites", false);
                RegDeleteValue(Constants.HKCU, sHKUTaskBand, "FavoritesRemovedChanges", false);
                RegDeleteValue(Constants.HKCU, sHKUTaskBand, "FavoritesChanges", false);
                RegDeleteValue(Constants.HKCU, sHKUTaskBand, "FavoritesResolve", false);
                RegDeleteValue(Constants.HKCU, sHKUTaskBand, "FavoritesVersion", false);
            }
        }
    }
    catch (Exception ex)
    {
        Log($"Error in ClearTaskBand: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   LoadUsersReg
//
//   Loads the HKCU for all local users
//-------------------------------------------------------------------------------
public void LoadUsersReg()
{
    try
    {
        string sValue;

        LogH1("Load User Registry Profiles");

        if (RegReadValue(Constants.HKLM, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProfilesDirectory", out sValue, "REG_EXPAND_SZ"))
        {
            string profilesDir = ExpandEnv(sValue);
            if (Directory.Exists(profilesDir))
            {
                foreach (DirectoryInfo profilefolder in new DirectoryInfo(profilesDir).GetDirectories())
                {
                    string ntuserPath = Path.Combine(profilefolder.FullName, "ntuser.dat");
                    if (File.Exists(ntuserPath))
                    {
                        LogOnly($" load: {ntuserPath} as HKU\\{profilefolder.Name}");
                        string regCmd = $"reg load \"HKU\\{profilefolder.Name}\" \"{ntuserPath}\"";
                        RunCommand(regCmd, true);
                    }
                    //        if (File.Exists(Path.Combine(profilefolder.FullName, "Local Settings", "Application Data", "Microsoft", "Windows", "UsrClass.dat")))
                    //        {
                    //            LogOnly($" load: {Path.Combine(profilefolder.FullName, "Local Settings", "Application Data", "Microsoft", "Windows", "UsrClass.dat")} as HKU\\{profilefolder.Name}_Classes");
                    //            string regCmd2 = $"reg load \"HKU\\{profilefolder.Name}_Classes\" \"{Path.Combine(profilefolder.FullName, "Local Settings", "Application Data", "Microsoft", "Windows", "UsrClass.dat")}\"";
                    //            RunCommand(regCmd2, true);
                    //        }
                }
            }
        }
    }
    catch (Exception ex)
    {
        Log($"Error in LoadUsersReg: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   ClearOfficeHKLM
//
//   Recursively search and clear the HKLM Office key from references in scope
//-------------------------------------------------------------------------------
public void ClearOfficeHKLM(string sSubKeyName)
{
    try
    {
        string key, name;
        string sValue;
        string[] arrKeys, arrNames, arrTypes;

        // recursion
        if (RegEnumKey(Constants.HKLM, sSubKeyName, out arrKeys))
        {
            foreach (string keyItem in arrKeys)
            {
                key = keyItem;
                ClearOfficeHKLM(sSubKeyName + "\\" + key);
            }
        }

        // identify & clear removable entries
        if (RegEnumValues(Constants.HKLM, sSubKeyName, out arrNames, out arrTypes))
        {
            foreach (string nameItem in arrNames)
            {
                name = nameItem;
                if (RegReadValue(Constants.HKLM, sSubKeyName, name, out sValue, "REG_SZ"))
                {
                    if (IsC2R(sValue)) RegDeleteValue(Constants.HKLM, sSubKeyName, name, false);
                }
            }
        }

        // clear out empty keys
        if (!RegEnumValues(Constants.HKLM, sSubKeyName, out arrNames, out arrTypes) &&
            !RegEnumKey(Constants.HKLM, sSubKeyName, out arrKeys) &&
            dicKeepSku.Count == 0)
        {
            RegDeleteKey(Constants.HKLM, sSubKeyName);
        }
    }
    catch (Exception ex)
    {
        Log($"Error in ClearOfficeHKLM: {ex.Message}");
    }
}


'-------------------------------------------------------------------------------
'
'                                        Helper Functions
'
'-------------------------------------------------------------------------------

//-------------------------------------------------------------------------------
//   IsC2R
//
//   Check if the passed in string is related to C2R
//   Returns TRUE if in C2R scope
//-------------------------------------------------------------------------------
public bool IsC2R(string sValue)
{
    const string OREF = "\\ROOT\\OFFICE1";
    const string OREFROOT = "Microsoft Office\\Root\\";
    const string OREGREFC2R15 = "Microsoft Office 15";
    const string OREGREFC2R16 = "Microsoft Office 16";
    const string OCOMMON = "\\microsoft shared\\ClickToRun";
    const string OMANIFEST = "\\Microsoft Office\\PackageManifests";
    const string OSUNRISE = "\\Microsoft Office\\PackageSunrisePolicies";

    if (string.IsNullOrEmpty(sValue)) return false;

    string sValueLower = sValue.ToLower();
    return sValueLower.Contains(OREF.ToLower()) ||
           sValueLower.Contains(OREFROOT.ToLower()) ||
           sValueLower.Contains(OCOMMON.ToLower()) ||
           sValueLower.Contains(OMANIFEST.ToLower()) ||
           sValueLower.Contains(OSUNRISE.ToLower()) ||
           sValueLower.Contains(OREGREFC2R15.ToLower()) ||
           sValueLower.Contains(OREGREFC2R16.ToLower());
}

//-------------------------------------------------------------------------------
//   CheckRegPermissions
//
//   Test the permissions on some key registry locations to determine if
//   sufficient permissions are given.
//-------------------------------------------------------------------------------
public bool CheckRegPermissions()
{
    const int KEY_QUERY_VALUE = 0x0001;
    const int KEY_SET_VALUE = 0x0002;
    const int KEY_CREATE_SUB_KEY = 0x0004;
    const int DELETE = 0x00010000;

    try
    {
        string sSubKeyName = @"Software\Microsoft\Windows\";
        bool fReturn;

        // Note: This requires WMI StdRegProv CheckAccess method
        // For now, we'll use a simplified check using RegistryKey.OpenSubKey with RegistryRights
        using (var baseKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(sSubKeyName, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.QueryValues | System.Security.AccessControl.RegistryRights.SetValue | System.Security.AccessControl.RegistryRights.CreateSubKey | System.Security.AccessControl.RegistryRights.Delete))
        {
            return baseKey != null;
        }
    }
    catch
    {
        return false;
    }
}

//-------------------------------------------------------------------------------
//   GetMyProcessId
//
//   Returns the process id of the own process
//-------------------------------------------------------------------------------
public int GetMyProcessId()
{
    try
    {
        // In C#, we can directly get the current process ID
        return System.Diagnostics.Process.GetCurrentProcess().Id;
    }
    catch
    {
        // Fallback: try to get from WMI
        try
        {
            var scope = new System.Management.ManagementScope(@"\\.\root\cimv2");
            string queryString = $"SELECT * FROM Win32_Process WHERE Name='{System.Diagnostics.Process.GetCurrentProcess().ProcessName}.exe' AND CommandLine LIKE '%{Constants.SCRIPTNAME}%'";
            var query = new System.Management.ObjectQuery(queryString);
            using (var searcher = new System.Management.ManagementObjectSearcher(scope, query))
            {
                foreach (System.Management.ManagementObject process in searcher.Get())
                {
                    return Convert.ToInt32(process["ProcessId"]);
                }
            }
        }
        catch { }
        return 0;
    }
}

//-------------------------------------------------------------------------------
//   Delimiter
//
//   Returns the delimiter for a passed in string
//-------------------------------------------------------------------------------
public string Delimiter(string sVersion)
{
    if (string.IsNullOrEmpty(sVersion)) return " ";

    foreach (char c in sVersion)
    {
        int iAsc = (int)c;
        if (!(iAsc >= 48 && iAsc <= 57))
        {
            return c.ToString();
        }
    }
    return " ";
}

//-------------------------------------------------------------------------------
//   GetExpandedGuid
//
//   Returns the expanded string from a compressed GUID
//-------------------------------------------------------------------------------
public string GetExpandedGuid(string sGuid)
{
    //Ensure valid length
    if (sGuid == null || sGuid.Length != 32) return "";

    char[] chars = sGuid.ToCharArray();
    Array.Reverse(chars, 0, 8);
    string part1 = new string(chars, 0, 8);
    Array.Reverse(chars, 8, 4);
    string part2 = new string(chars, 8, 4);
    Array.Reverse(chars, 12, 4);
    string part3 = new string(chars, 12, 4);

    string part4 = "";
    for (int i = 17; i <= 20; i++)
    {
        // VBScript uses 1-based indexing: Mid(sGuid, i+1, 1) for odd, Mid(sGuid, i-1, 1) for even
        // Convert to 0-based: i+1 becomes i, i-1 becomes i-2
        if (i % 2 == 1)
        {
            part4 += sGuid[i]; // i+1 in 1-based = i in 0-based
        }
        else
        {
            part4 += sGuid[i - 2]; // i-1 in 1-based = i-2 in 0-based
        }
    }

    string part5 = "";
    for (int i = 21; i < 32; i++) // Changed to < 32 since sGuid is 32 chars (0-31)
    {
        // VBScript uses 1-based indexing: Mid(sGuid, i+1, 1) for odd, Mid(sGuid, i-1, 1) for even
        // Convert to 0-based: i+1 becomes i, i-1 becomes i-2
        if (i % 2 == 1)
        {
            part5 += sGuid[i]; // i+1 in 1-based = i in 0-based
        }
        else
        {
            part5 += sGuid[i - 2]; // i-1 in 1-based = i-2 in 0-based
        }
    }

    return "{" + part1 + "-" + part2 + "-" + part3 + "-" + part4 + "-" + part5 + "}";
}

//-------------------------------------------------------------------------------
//   GetCompressedGuid
//
//   Returns the compressed string for a GUID
//-------------------------------------------------------------------------------
public string GetCompressedGuid(string sGuid)
{
    //Ensure Valid Length
    if (sGuid == null || sGuid.Length != 38) return "";

    char[] part1 = sGuid.Substring(2, 8).ToCharArray();
    Array.Reverse(part1);
    char[] part2 = sGuid.Substring(11, 4).ToCharArray();
    Array.Reverse(part2);
    char[] part3 = sGuid.Substring(16, 4).ToCharArray();
    Array.Reverse(part3);

    string sCompGUID = new string(part1) + new string(part2) + new string(part3);

    for (int i = 21; i <= 24; i++)
    {
        if (i % 2 == 1)
        {
            sCompGUID += sGuid[i];
        }
        else
        {
            sCompGUID += sGuid[i - 2];
        }
    }

    for (int i = 26; i <= 37; i++)
    {
        if (i % 2 == 1)
        {
            sCompGUID += sGuid[i - 2];
        }
        else
        {
            sCompGUID += sGuid[i];
        }
    }

    return sCompGUID;
}

//-------------------------------------------------------------------------------
//   GetDecodedGuid
//
//   Returns the GUID from a squished format
//-------------------------------------------------------------------------------
public bool GetDecodedGuid(string sEncGuid, out string sGuid)
{
    sGuid = "";
    if (string.IsNullOrEmpty(sEncGuid) || sEncGuid.Length < 20) return false;

    string sTable = "0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff," +
                    "0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff," +
                    "0xff,0x00,0xff,0xff,0x01,0x02,0x03,0x04,0x05,0x06,0x07,0x08,0x09,0x0a,0x0b,0xff," +
                    "0x0c,0x0d,0x0e,0x0f,0x10,0x11,0x12,0x13,0x14,0x15,0xff,0xff,0xff,0x16,0xff,0x17," +
                    "0x18,0x19,0x1a,0x1b,0x1c,0x1d,0x1e,0x1f,0x20,0x21,0x22,0x23,0x24,0x25,0x26,0x27," +
                    "0x28,0x29,0x2a,0x2b,0x2c,0x2d,0x2e,0x2f,0x30,0x31,0x32,0x33,0xff,0x34,0x35,0x36," +
                    "0x37,0x38,0x39,0x3a,0x3b,0x3c,0x3d,0x3e,0x3f,0x40,0x41,0x42,0x43,0x44,0x45,0x46," +
                    "0x47,0x48,0x49,0x4a,0x4b,0x4c,0x4d,0x4e,0x4f,0x50,0x51,0x52,0xff,0x53,0x54,0xff";
    string[] arrTable = sTable.Split(',');
    long lTotal = 0;
    long pow85 = 1;
    string sDecode = "";
    bool fFailed = false;

    for (int i = 0; i < 20 && i < sEncGuid.Length; i++)
    {
        fFailed = true;
        if (i % 5 == 0)
        {
            lTotal = 0;
            pow85 = 1;
        }
        int iAsc = (int)sEncGuid[i];
        if (iAsc >= 128) break;
        if (iAsc >= arrTable.Length) break;
        string sHex = arrTable[iAsc];
        if (sHex == "0xff") break;
        int iChr = Convert.ToInt32(sHex.Substring(2), 16);
        lTotal = lTotal + (iChr * pow85);
        if (i % 5 == 4) sDecode = sDecode + DecToHex(lTotal);
        pow85 = pow85 * 85;
        fFailed = false;
    }

    if (!fFailed && sDecode.Length >= 32)
    {
        sGuid = "{" + sDecode.Substring(0, 8) + "-" +
                sDecode.Substring(12, 4) + "-" +
                sDecode.Substring(8, 4) + "-" +
                sDecode.Substring(22, 2) + sDecode.Substring(20, 2) + "-" +
                sDecode.Substring(18, 2) + sDecode.Substring(16, 2) + sDecode.Substring(30, 2) + sDecode.Substring(28, 2) + sDecode.Substring(26, 2) + sDecode.Substring(24, 2) + "}";
    }

    return !fFailed;
}

//-------------------------------------------------------------------------------
//   DecToHex
//
//   Convert a long decimal to hex
//-------------------------------------------------------------------------------
public string DecToHex(long lDec)
{
    string[] arrChr = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F" };
    string sHex = "";
    long lVal = lDec;
    long lExp = (long)Math.Pow(16, 10);

    while (lExp >= 1)
    {
        if (lVal >= lExp)
        {
            int index = (int)(lVal / lExp);
            sHex = sHex + arrChr[index];
            lVal = lVal - lExp * index;
        }
        else
        {
            sHex = sHex + "0";
            if (sHex == "0") sHex = "";
        }
        lExp = lExp / 16;
    }

    int iLen = 8 - sHex.Length;
    if (iLen > 0) sHex = new string('0', iLen) + sHex;
    return sHex;
}

//-------------------------------------------------------------------------------
//   RelaunchAs64Host
//
//   Relaunch self with 64 bit CScript host
//-------------------------------------------------------------------------------
public void RelaunchAs64Host()
{
    try
    {
        bool fQuietRelaunch = false;
        string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string sysDir = Path.GetDirectoryName(exePath);
        // Replace syswow64 with sysnative (case-insensitive)
        if (sysDir.IndexOf("syswow64", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            sysDir = sysDir.Substring(0, sysDir.IndexOf("syswow64", StringComparison.OrdinalIgnoreCase)) + "sysnative" + sysDir.Substring(sysDir.IndexOf("syswow64", StringComparison.OrdinalIgnoreCase) + 8);
        }
        string sCmd = sysDir + "\\cscript.exe \"" + exePath + "\"";

        if (fQuiet) fQuietRelaunch = true;

        var args = Environment.GetCommandLineArgs();
        for (int i = 1; i < args.Length; i++)
        {
            string argument = args[i];
            sCmd = sCmd + " \"" + argument + "\"";
            switch (argument.ToUpper())
            {
                case "/Q":
                case "/QUIET":
                    fQuietRelaunch = true;
                    break;
            }
        }

        sCmd = sCmd + " /ChangedHostBitness";
        if (fQuietRelaunch)
        {
            // Replace cscript.exe with wscript.exe (case-insensitive)
            int idx = sCmd.IndexOf("\\cscript.exe", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                sCmd = sCmd.Substring(0, idx) + "\\wscript.exe" + sCmd.Substring(idx + 13);
            }
            int exitCode = RunCommand(sCmd, true);
            Environment.Exit(exitCode);
        }
        else
        {
            int exitCode = RunCommand(sCmd, true);
            Environment.Exit(exitCode);
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RelaunchAs64Host: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   RelaunchElevated
//
//   Relaunch the script with elevated permissions
//-------------------------------------------------------------------------------
public void RelaunchElevated()
{
    try
    {
        SetError(Constants.ERROR_RELAUNCH);
        // Shell object for relaunch - requires COM interop
        // Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
        // dynamic oShell = Activator.CreateInstance(shellAppType);

        // Note: Command line has not been parsed at this point
        // build command line for relaunch
        string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string sCmdLine = "\"" + exePath + "\"";

        var args = Environment.GetCommandLineArgs();
        for (int i = 1; i < args.Length; i++)
        {
            string argument = args[i];
            switch (argument.ToUpper())
            {
                case "/Q":
                case "/QUIET":
                    //Don't try to relaunch in quiet mode
                    SetError(Constants.ERROR_ELEVATION);
                    return;
                case "UAC":
                    //Already tried elevated relaunch
                    SetError(Constants.ERROR_ELEVATION);
                    return;
                default:
                    sCmdLine = sCmdLine + " \"" + argument + "\"";
                    break;
            }
        }

        // prep work to get the return value from the elevated process
        int iParentProcessId = GetMyProcessId();

        //    // make user aware of elevation attempt after reboot
        //    if (RegReadValue(Constants.HKCU, @"SOFTWARE\Microsoft\Office\15.0\CleanC2R", "Rerun", out sValue, "REG_DWORD"))
        //    {
        //        MessageBox.Show("System reboot complete. OffScrub will now prompt for elevation!", Constants.SCRIPTNAME + " - NOTE!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }

        // launch the elevated instance
        // Requires COM interop: oShell.ShellExecute("cscript.exe", sCmdLine + " /NoElevate UAC", "", "runas", 1);
        // For now, using ProcessStartInfo with Verb = "runas"
        var startInfo = new ProcessStartInfo
        {
            FileName = "cscript.exe",
            Arguments = sCmdLine + " /NoElevate UAC",
            Verb = "runas",
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Normal
        };

        Process elevatedProcess = Process.Start(startInfo);
        if (elevatedProcess != null)
        {
            int iSpawnedProcessId = elevatedProcess.Id;
            // monitor the process to detect the end
            while (!elevatedProcess.HasExited)
            {
                System.Threading.Thread.Sleep(3000);
            }
            // get the return value from the file
            int retVal = GetRetValFromFile();
            Environment.Exit(retVal);
        }
        else
        {
            // elevation failed (user declined)
            SetError(Constants.ERROR_ELEVATION_USERDECLINED);
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RelaunchElevated: {ex.Message}");
        SetError(Constants.ERROR_ELEVATION);
    }
}

//-------------------------------------------------------------------------------
//   RelaunchAsCScript
//
//   Relaunch self with Cscript as host
//-------------------------------------------------------------------------------
public void RelaunchAsCScript()
{
    try
    {
        bool fQuietNoCScript = false;
        SetError(Constants.ERROR_RELAUNCH);
        string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string sysDir = Path.GetDirectoryName(exePath);
        string sCmdLine = "cmd.exe /c " + sysDir + "\\cscript.exe //NOLOGO \"" + exePath + "\"";

        var args = Environment.GetCommandLineArgs();
        for (int i = 1; i < args.Length; i++)
        {
            string argument = args[i];
            sCmdLine = sCmdLine + " \"" + argument + "\"";
            switch (argument.ToUpper())
            {
                case "/Q":
                case "/QUIET":
                    fQuietNoCScript = true;
                    ClearError(Constants.ERROR_RELAUNCH);
                    break;
            }
        }

        sCmdLine = sCmdLine + " \"/ChangedScriptHost\"";

        if (!fQuietNoCScript)
        {
            int exitCode = RunCommand(sCmdLine, true);
            Environment.Exit(exitCode);
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RelaunchAsCScript: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   SetError
//
//   Set error bit(s)
//-------------------------------------------------------------------------------
public void SetError(int ErrorBit)
{
    iError = iError | ErrorBit;
    switch (ErrorBit)
    {
        case Constants.ERROR_DCAF_FAILURE:
        case Constants.ERROR_STAGE2:
        case Constants.ERROR_ELEVATION_USERDECLINED:
        case Constants.ERROR_ELEVATION:
        case Constants.ERROR_SCRIPTINIT:
            iError = iError | Constants.ERROR_FAIL;
            break;
    }
}

//-------------------------------------------------------------------------------
//   ClearError
//
//   Unset error bit(s)
//-------------------------------------------------------------------------------
public void ClearError(int ErrorBit)
{
    iError = iError & (Constants.ERROR_ALL - ErrorBit);
    switch (ErrorBit)
    {
        case Constants.ERROR_ELEVATION_USERDECLINED:
        case Constants.ERROR_ELEVATION:
        case Constants.ERROR_SCRIPTINIT:
            iError = iError & (Constants.ERROR_ALL - Constants.ERROR_FAIL);
            break;
    }
}

//-------------------------------------------------------------------------------
//   SetRetVal
//
//   Write return value to file
//-------------------------------------------------------------------------------
public void SetRetVal(int errorValue)
{
    //don't fail script execution if writing the return value to file fails
    try
    {
        string retValPath = Path.Combine(sScrubDir, Constants.RETVALFILE);
        File.WriteAllText(retValPath, errorValue.ToString());
    }
    catch (Exception ex)
    {
        Log($"Error writing return value to file: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   GetRetValFromFile
//
//   Read return value from file.
//   Used to ensure return value can get obtained from an elevated process
//-------------------------------------------------------------------------------
public int GetRetValFromFile()
{
    //don't fail script execution when getting the return value from file fails
    try
    {
        string retValPath = Path.Combine(sScrubDir, Constants.RETVALFILE);
        if (File.Exists(retValPath))
        {
            string content = File.ReadAllText(retValPath);
            if (int.TryParse(content, out int iRetValFromFile))
            {
                return iRetValFromFile;
            }
        }
    }
    catch (Exception ex)
    {
        Log($"Error reading return value from file: {ex.Message}");
    }
    return Constants.ERROR_UNKNOWN;
}

//-------------------------------------------------------------------------------
//   CreateLog
//
//   Create the removal log file
//-------------------------------------------------------------------------------
public void CreateLog()
{
    try
    {
        // create the log file
        string computerName = Environment.MachineName;
        string dateTimeStr = DateTime.Now.ToString("yyyyMMddHHmmss");
        string sLogName = Path.Combine(sLogDir, $"{computerName}_{dateTimeStr}_ScrubLog.txt");

        try
        {
            LogStream = new StreamWriter(sLogName, false, Encoding.UTF8);
        }
        catch
        {
            sLogDir = sScrubDir;
            sLogName = Path.Combine(sLogDir, $"{computerName}_{dateTimeStr}_ScrubLog.txt");
            LogStream = new StreamWriter(sLogName, false, Encoding.UTF8);
        }

        LogH2($"Microsoft Customer Support Services - {Constants.ONAME} Removal Utility\r\n\r\n" +
              $"Version:\t{Constants.SCRIPTVERSION}\r\n" +
              $"64 bit OS:\t{f64}\r\n" +
              $"Removal start:\t{DateTime.Now:HH:mm:ss}");
        LogH2($"OS Details: {sOSinfo}\r\n");
        fLogInitialized = true;
    }
    catch (Exception ex)
    {
        Log($"Error in CreateLog: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   HiveString
//
//   Translates the numeric constant into the human readable registry hive string
//-------------------------------------------------------------------------------
public string HiveString(int hDefKey)
{
    switch (hDefKey)
    {
        case Constants.HKCR:
            return "HKEY_CLASSES_ROOT";
        case Constants.HKCU:
            return "HKEY_CURRENT_USER";
        case Constants.HKLM:
            return "HKEY_LOCAL_MACHINE";
        case Constants.HKU:
            return "HKEY_USERS";
        default:
            return hDefKey.ToString();
    }
}

//-------------------------------------------------------------------------------
//   RegKeyExists
//
//   Returns a boolean for the test on existence of a given registry key
//-------------------------------------------------------------------------------
public bool RegKeyExists(int hDefKey, string sSubKeyName)
{
    string[] arrKeys;
    return RegEnumKey(hDefKey, sSubKeyName, out arrKeys);
}

//-------------------------------------------------------------------------------
//   RegValExists
//
//   Returns a boolean for the test on existence of a given registry value
//-------------------------------------------------------------------------------
public bool RegValExists(int hDefKey, string sSubKeyName, string sName)
{
    if (!RegKeyExists(hDefKey, sSubKeyName)) return false;

    string[] arrValueNames, arrValueTypes;
    if (RegEnumValues(hDefKey, sSubKeyName, out arrValueNames, out arrValueTypes) && arrValueNames != null)
    {
        for (int i = 0; i < arrValueNames.Length; i++)
        {
            if (arrValueNames[i].Trim().Equals(sName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
    }
    return false;
}

//-------------------------------------------------------------------------------
//   RegReadValue
//
//   Read the value of a given registry entry
//   The correct type has to be passed in as argument
//-------------------------------------------------------------------------------
public bool RegReadValue(int hDefKey, string sSubKeyName, string sName, out string sValue, string sType)
{
    sValue = "";
    int RetVal = -1;
    string[] arrValues;

    try
    {
        Microsoft.Win32.RegistryKey baseKey = GetRegistryKey(hDefKey);
        if (baseKey == null) return false;

        Microsoft.Win32.RegistryKey key = baseKey.OpenSubKey(sSubKeyName);
        if (key == null && f64)
        {
            key = baseKey.OpenSubKey(Wow64Key(hDefKey, sSubKeyName));
        }

        if (key == null) return false;

        switch (sType.ToUpper())
        {
            case "1":
            case "REG_SZ":
                object val = key.GetValue(sName);
                if (val != null)
                {
                    sValue = val.ToString();
                    RetVal = 0;
                }
                break;
            case "2":
            case "REG_EXPAND_SZ":
                val = key.GetValue(sName);
                if (val != null)
                {
                    sValue = Environment.ExpandEnvironmentVariables(val.ToString());
                    RetVal = 0;
                }
                break;
            case "3":
            case "REG_BINARY":
                byte[] binaryVal = (byte[])key.GetValue(sName);
                if (binaryVal != null)
                {
                    sValue = BitConverter.ToString(binaryVal).Replace("-", "");
                    RetVal = 0;
                }
                break;
            case "4":
            case "REG_DWORD":
                object dwordVal = key.GetValue(sName);
                if (dwordVal != null)
                {
                    sValue = Convert.ToInt32(dwordVal).ToString();
                    RetVal = 0;
                }
                break;
            case "7":
            case "REG_MULTI_SZ":
                string[] multiVal = (string[])key.GetValue(sName);
                if (multiVal != null)
                {
                    sValue = string.Join("\r", multiVal);
                    RetVal = 0;
                }
                break;
            default:
                RetVal = -1;
                break;
        }
        key.Close();
        baseKey.Close();
    }
    catch (Exception ex)
    {
        Log($"Error in RegReadValue: {ex.Message}");
        RetVal = -1;
    }

    return (RetVal == 0);
}

private Microsoft.Win32.RegistryKey GetRegistryKey(int hDefKey)
{
    switch (hDefKey)
    {
        case Constants.HKCR:
            return Microsoft.Win32.Registry.ClassesRoot;
        case Constants.HKCU:
            return Microsoft.Win32.Registry.CurrentUser;
        case Constants.HKLM:
            return Microsoft.Win32.Registry.LocalMachine;
        case Constants.HKU:
            return Microsoft.Win32.Registry.Users;
        default:
            return null;
    }
}

//-------------------------------------------------------------------------------
//   RegEnumValues
//
//   Enumerate a registry key to return all values
//-------------------------------------------------------------------------------
public bool RegEnumValues(int hDefKey, string sSubKeyName, out string[] arrNames, out string[] arrTypes)
{
    arrNames = null;
    arrTypes = null;
    bool RetVal = false;
    bool RetVal64 = false;
    string[] arrNames32 = null, arrNames64 = null;
    string[] arrTypes32 = null, arrTypes64 = null;

    try
    {
        if (f64)
        {
            RetVal = EnumRegistryValues(hDefKey, sSubKeyName, out arrNames32, out arrTypes32);
            RetVal64 = EnumRegistryValues(hDefKey, Wow64Key(hDefKey, sSubKeyName), out arrNames64, out arrTypes64);

            if (RetVal && !RetVal64 && arrNames32 != null && arrTypes32 != null)
            {
                arrNames = arrNames32;
                arrTypes = arrTypes32;
            }
            else if (!RetVal && RetVal64 && arrNames64 != null && arrTypes64 != null)
            {
                arrNames = arrNames64;
                arrTypes = arrTypes64;
            }
            else if (RetVal && RetVal64 && arrNames32 != null && arrNames64 != null && arrTypes32 != null && arrTypes64 != null)
            {
                List<string> combinedNames = new List<string>(arrNames32);
                combinedNames.AddRange(arrNames64);
                List<string> combinedTypes = new List<string>(arrTypes32);
                combinedTypes.AddRange(arrTypes64);
                arrNames = RemoveDuplicates(combinedNames.ToArray());
                arrTypes = RemoveDuplicates(combinedTypes.ToArray());
            }
        }
        else
        {
            RetVal = EnumRegistryValues(hDefKey, sSubKeyName, out arrNames, out arrTypes);
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RegEnumValues: {ex.Message}");
    }

    return (RetVal || RetVal64) && arrNames != null && arrTypes != null;
}

private bool EnumRegistryValues(int hDefKey, string sSubKeyName, out string[] arrNames, out string[] arrTypes)
{
    arrNames = null;
    arrTypes = null;
    try
    {
        Microsoft.Win32.RegistryKey baseKey = GetRegistryKey(hDefKey);
        if (baseKey == null) return false;

        Microsoft.Win32.RegistryKey key = baseKey.OpenSubKey(sSubKeyName);
        if (key == null) return false;

        List<string> names = new List<string>();
        List<string> types = new List<string>();

        foreach (string valueName in key.GetValueNames())
        {
            names.Add(valueName);
            Microsoft.Win32.RegistryValueKind kind = key.GetValueKind(valueName);
            types.Add(kind.ToString());
        }

        arrNames = names.ToArray();
        arrTypes = types.ToArray();
        key.Close();
        baseKey.Close();
        return true;
    }
    catch
    {
        return false;
    }
}

//-------------------------------------------------------------------------------
//   RegEnumKey
//
//   Enumerate a registry key to return all subkeys
//-------------------------------------------------------------------------------
public bool RegEnumKey(int hDefKey, string sSubKeyName, out string[] arrKeys)
{
    arrKeys = null;
    bool RetVal = false;
    bool RetVal64 = false;
    string[] arrKeys32 = null, arrKeys64 = null;

    try
    {
        if (f64)
        {
            RetVal = EnumRegistryKeys(hDefKey, sSubKeyName, out arrKeys32);
            RetVal64 = EnumRegistryKeys(hDefKey, Wow64Key(hDefKey, sSubKeyName), out arrKeys64);

            if (RetVal && !RetVal64 && arrKeys32 != null)
            {
                arrKeys = arrKeys32;
            }
            else if (!RetVal && RetVal64 && arrKeys64 != null)
            {
                arrKeys = arrKeys64;
            }
            else if (RetVal && RetVal64)
            {
                if (arrKeys32 != null && arrKeys64 != null)
                {
                    List<string> combined = new List<string>(arrKeys32);
                    combined.AddRange(arrKeys64);
                    arrKeys = RemoveDuplicates(combined.ToArray());
                }
                else if (arrKeys64 != null)
                {
                    arrKeys = arrKeys64;
                }
                else
                {
                    arrKeys = arrKeys32;
                }
            }
        }
        else
        {
            RetVal = EnumRegistryKeys(hDefKey, sSubKeyName, out arrKeys);
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RegEnumKey: {ex.Message}");
    }

    return (RetVal || RetVal64) && arrKeys != null;
}

private bool EnumRegistryKeys(int hDefKey, string sSubKeyName, out string[] arrKeys)
{
    arrKeys = null;
    try
    {
        Microsoft.Win32.RegistryKey baseKey = GetRegistryKey(hDefKey);
        if (baseKey == null) return false;

        Microsoft.Win32.RegistryKey key = baseKey.OpenSubKey(sSubKeyName);
        if (key == null) return false;

        List<string> keys = new List<string>();
        foreach (string subKeyName in key.GetSubKeyNames())
        {
            keys.Add(subKeyName);
        }

        arrKeys = keys.ToArray();
        key.Close();
        baseKey.Close();
        return true;
    }
    catch
    {
        return false;
    }
}

//-------------------------------------------------------------------------------
//   RegDeleteValue
//
//   Wrapper around oReg.DeleteValue to handle 64 bit
//-------------------------------------------------------------------------------
public void RegDeleteValue(int hDefKey, string sSubKeyName, string sName, bool fRegMultiSZ)
{
    try
    {
        string sDelKeyName, sValue;
        int iRetVal = 0;
        bool fKeep;

        // ensure trailing "\"
        if (!sSubKeyName.EndsWith("\\")) sSubKeyName = sSubKeyName + "\\";
        while (sSubKeyName.Contains("\\\\"))
        {
            sSubKeyName = sSubKeyName.Replace("\\\\", "\\");
        }

        fKeep = dicKeepReg.ContainsKey((sSubKeyName + sName).ToLower());
        if (!fKeep && f64) fKeep = dicKeepReg.ContainsKey((Wow64Key(hDefKey, sSubKeyName) + sName).ToLower());

        if (fKeep)
        {
            LogOnly($"Disallowing the delete of still required keypath element: {HiveString(hDefKey)}\\{sSubKeyName}{sName}");
            if (!fForce) return;
        }

        // check on forced delete
        if (fKeep)
        {
            LogOnly($"Enforced delete of still required keypath element: {HiveString(hDefKey)}\\{sSubKeyName}{sName}");
            LogOnly("   Remaining applications will need a repair!");
        }

        // ensure value exists
        if (RegValExists(hDefKey, sSubKeyName, sName))
        {
            sDelKeyName = sSubKeyName;
        }
        else if (RegValExists(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName))
        {
            sDelKeyName = Wow64Key(hDefKey, sSubKeyName);
        }
        else
        {
            LogOnly($"Value not found. Cannot delete value: {HiveString(hDefKey)}\\{sSubKeyName}{sName}");
            return;
        }

        // prevent unintentional, unsafe REG_MULTI_SZ delete
        if (RegReadValue(hDefKey, sDelKeyName, sName, out sValue, "REG_MULTI_SZ") && !fRegMultiSZ)
        {
            LogOnly($"Disallowing unsafe delete of REG_MULTI_SZ: {HiveString(hDefKey)}\\{sDelKeyName}{sName}");
            return;
        }

        // execute delete operation
        if (!fDetectOnly)
        {
            LogOnly($"Delete registry value: {HiveString(hDefKey)}\\{sDelKeyName} -> {sName}");
            iRetVal = DeleteRegistryValue(hDefKey, sDelKeyName, sName);
            CheckError("RegDeleteValue");
            if (iRetVal != 0)
            {
                LogOnly($"     Delete failed. Return value: {iRetVal}");
                SetError(Constants.ERROR_STAGE2);
            }
        }
        else
        {
            LogOnly($"Preview mode. Disallowing delete registry value: {HiveString(hDefKey)}\\{sDelKeyName} -> {sName}");
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RegDeleteValue: {ex.Message}");
        CheckError("RegDeleteValue");
    }
}

private int DeleteRegistryValue(int hDefKey, string sSubKeyName, string sName)
{
    try
    {
        Microsoft.Win32.RegistryKey baseKey = GetRegistryKey(hDefKey);
        if (baseKey == null) return -1;

        Microsoft.Win32.RegistryKey key = baseKey.OpenSubKey(sSubKeyName, true);
        if (key == null) return -1;

        key.DeleteValue(sName, false);
        key.Close();
        baseKey.Close();
        return 0;
    }
    catch
    {
        return -1;
    }
}

//-------------------------------------------------------------------------------
//   RegDeleteKey
//
//   Wrappper around RegDeleteKeyEx to handle 64bit
//-------------------------------------------------------------------------------
public void RegDeleteKey(int hDefKey, string sSubKeyName)
{
    try
    {
        string sDelKeyName;
        bool fKeep;

        // ensure trailing "\"
        if (!sSubKeyName.EndsWith("\\")) sSubKeyName = sSubKeyName + "\\";
        while (sSubKeyName.Contains("\\\\"))
        {
            sSubKeyName = sSubKeyName.Replace("\\\\", "\\");
        }

        fKeep = dicKeepReg.ContainsKey(sSubKeyName.ToLower());
        if (!fKeep && f64) fKeep = dicKeepReg.ContainsKey(Wow64Key(hDefKey, sSubKeyName).ToLower());

        if (fKeep)
        {
            LogOnly($"Disallowing the delete of still required keypath element: {HiveString(hDefKey)}\\{sSubKeyName}");
            if (!fForce) return;
        }

        // check on forced delete
        if (fKeep)
        {
            LogOnly($"Enforced delete of still required keypath element: {HiveString(hDefKey)}\\{sSubKeyName}");
            LogOnly("   Remaining applications will need a repair!");
        }

        if (sSubKeyName.Length > 1)
        {
            //Strip of trailing "\"
            sSubKeyName = sSubKeyName.Substring(0, sSubKeyName.Length - 1);
        }

        // ensure key exists
        if (RegKeyExists(hDefKey, sSubKeyName))
        {
            sDelKeyName = sSubKeyName;
        }
        else if (f64 && RegKeyExists(hDefKey, Wow64Key(hDefKey, sSubKeyName)))
        {
            sDelKeyName = Wow64Key(hDefKey, sSubKeyName);
        }
        else
        {
            LogOnly($"Key not found. Cannot delete key: {HiveString(hDefKey)}\\{sSubKeyName}");
            return;
        }

        // execute delete
        if (!fDetectOnly)
        {
            LogOnly($"Delete registry key: {HiveString(hDefKey)}\\{sDelKeyName}");
            try
            {
                RegDeleteKeyEx(hDefKey, sDelKeyName);
            }
            catch (Exception ex)
            {
                Log($"Error in RegDeleteKeyEx: {ex.Message}");
            }
        }
        else
        {
            LogOnly($"Preview mode. Disallowing delete of registry key: {HiveString(hDefKey)}\\{sSubKeyName}");
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RegDeleteKey: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   RegDeleteKeyEx
//
//   Recursively delete a registry structure
//-------------------------------------------------------------------------------
public void RegDeleteKeyEx(int hDefKey, string sSubKeyName)
{
    try
    {
        string[] arrSubkeys;
        string sSubkey;
        int iRetVal = 0;

        //Strip of trailing "\"
        if (sSubKeyName.Length > 1)
        {
            if (sSubKeyName.EndsWith("\\")) sSubKeyName = sSubKeyName.Substring(0, sSubKeyName.Length - 1);
        }

        // exception handler
        if (hDefKey == Constants.HKLM && sSubKeyName == @"SOFTWARE\Microsoft\Office\15.0\ClickToRun")
        {
            if (!fDetectOnly) iRetVal = RunCommand("reg delete HKLM\\SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun /f", true);
            return;
        }

        // regular recursion
        if (RegEnumKey(hDefKey, sSubKeyName, out arrSubkeys) && arrSubkeys != null)
        {
            foreach (string subkeyItem in arrSubkeys)
            {
                sSubkey = subkeyItem;
                RegDeleteKeyEx(hDefKey, sSubKeyName + "\\" + sSubkey);
            }
        }

        if (!fDetectOnly)
        {
            iRetVal = DeleteRegistryKey(hDefKey, sSubKeyName);
            if (iRetVal != 0) LogOnly($"     Delete failed. Return value: {iRetVal}");
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RegDeleteKeyEx: {ex.Message}");
    }
}

private int DeleteRegistryKey(int hDefKey, string sSubKeyName)
{
    try
    {
        Microsoft.Win32.RegistryKey baseKey = GetRegistryKey(hDefKey);
        if (baseKey == null) return -1;

        int lastSlash = sSubKeyName.LastIndexOf('\\');
        if (lastSlash > 0)
        {
            string parentPath = sSubKeyName.Substring(0, lastSlash);
            string keyName = sSubKeyName.Substring(lastSlash + 1);
            Microsoft.Win32.RegistryKey parentKey = baseKey.OpenSubKey(parentPath, true);
            if (parentKey != null)
            {
                parentKey.DeleteSubKeyTree(keyName, false);
                parentKey.Close();
            }
        }
        else
        {
            baseKey.DeleteSubKeyTree(sSubKeyName, false);
        }
        baseKey.Close();
        return 0;
    }
    catch
    {
        return -1;
    }
}

//-------------------------------------------------------------------------------
//   Wow64Key
//
//   Return the 32bit regkey location on a 64bit environment
//-------------------------------------------------------------------------------
public string Wow64Key(int hDefKey, string sSubKeyName)
{
    int iPos;

    switch (hDefKey)
    {
        case Constants.HKCU:
            if (sSubKeyName.StartsWith("Software\\Classes\\"))
            {
                return sSubKeyName.Substring(0, 17) + "Wow6432Node\\" + sSubKeyName.Substring(17);
            }
            else
            {
                iPos = sSubKeyName.IndexOf('\\');
                if (iPos > 0)
                {
                    return sSubKeyName.Substring(0, iPos) + "\\Wow6432Node\\" + sSubKeyName.Substring(iPos + 1);
                }
                return "Wow6432Node\\" + sSubKeyName;
            }
        case Constants.HKLM:
            if (sSubKeyName.StartsWith("Software\\Classes\\"))
            {
                return sSubKeyName.Substring(0, 17) + "Wow6432Node\\" + sSubKeyName.Substring(17);
            }
            else
            {
                iPos = sSubKeyName.IndexOf('\\');
                if (iPos > 0)
                {
                    return sSubKeyName.Substring(0, iPos) + "\\Wow6432Node\\" + sSubKeyName.Substring(iPos + 1);
                }
                return "Wow6432Node\\" + sSubKeyName;
            }
        default:
            return "Wow6432Node\\" + sSubKeyName;
    }
}

//-------------------------------------------------------------------------------
//   RemoveDuplicates
//
//   Remove duplicate entries from a one dimensional array
//-------------------------------------------------------------------------------
public string[] RemoveDuplicates(string[] array)
{
    if (array == null) return null;

    HashSet<string> dicNoDupes = new HashSet<string>();
    foreach (string item in array)
    {
        if (!dicNoDupes.Contains(item))
        {
            dicNoDupes.Add(item);
        }
    }
    return dicNoDupes.ToArray();
}

//-------------------------------------------------------------------------------
//   CheckError
//
//   Checks the status of 'Err' and logs the error details if <> 0
//-------------------------------------------------------------------------------
public void CheckError(string sModule)
{
    // In C#, we check for exceptions differently - this method is called from catch blocks
    // For VBScript compatibility, we'll log if there's a last exception
    // Note: This is typically called from catch blocks where exception details are available
    // For now, we'll just log the module name - actual error details should be logged in catch blocks
    // This maintains compatibility with existing code that calls CheckError
}

//-------------------------------------------------------------------------------
//   LogH
//
//   Write a header log string to the log file
//-------------------------------------------------------------------------------
public void LogH(string sLog)
{
    if (LogStream != null)
    {
        LogStream.WriteLine("");
        sLog = sLog + "\r\n" + new string('=', sLog.Length);
        if (!fQuiet && fCScript)
        {
            Console.WriteLine("");
            Console.WriteLine(sLog);
        }
        LogStream.WriteLine(sLog);
    }
}

//-------------------------------------------------------------------------------
//   LogH1
//
//   Write a header log string to the log file
//-------------------------------------------------------------------------------
public void LogH1(string sLog)
{
    if (LogStream != null)
    {
        LogStream.WriteLine("");
        sLog = sLog + "\r\n" + new string('-', sLog.Length);
        if (!fQuiet && fCScript)
        {
            Console.WriteLine("");
            Console.WriteLine(sLog);
        }
        LogStream.WriteLine(sLog);
    }
}

//-------------------------------------------------------------------------------
//   LogH2
//
//   Write w/o indent Cmd window and the log file
//-------------------------------------------------------------------------------
public void LogH2(string sLog)
{
    if (!fQuiet && fCScript)
    {
        Console.WriteLine(sLog);
    }
    if (LogStream != null)
    {
        LogStream.WriteLine("");
        LogStream.WriteLine(sLog);
    }
}

//-------------------------------------------------------------------------------
//   Log
//
//   Echos the log string to the Cmd window and the log file
//-------------------------------------------------------------------------------
public void Log(string sLog)
{
    if (!fQuiet && fCScript)
    {
        Console.WriteLine(sLog);
    }
    if (LogStream != null)
    {
        if (string.IsNullOrEmpty(sLog))
        {
            LogStream.WriteLine();
        }
        else
        {
            LogStream.WriteLine($"   {DateTime.Now:HH:mm:ss}: {sLog}");
        }
    }
}

//-------------------------------------------------------------------------------
//   LogOnly
//
//   Commits the log string to the log file
//-------------------------------------------------------------------------------
public void LogOnly(string sLog)
{
    if (LogStream != null)
    {
        if (string.IsNullOrEmpty(sLog))
        {
            LogStream.WriteLine();
        }
        else
        {
            LogStream.WriteLine($"   {DateTime.Now:HH:mm:ss}: {sLog}");
        }
    }
}

public void LogY(string sLog)
{
    LogPipe(sLog);
}

public void LogPipe(string sLog)
{
    try
    {
        // Named pipe implementation - requires System.IO.Pipes
        // For now, using file-based approach or placeholder
        // Note: Named pipes require NamedPipeClientStream or similar
        using (var pipeStream = new System.IO.FileStream(@"\\.\pipe\offscrub_pipe", FileMode.Create, FileAccess.Write, FileShare.Read))
        {
            using (var writer = new StreamWriter(pipeStream))
            {
                writer.WriteLine(sLog);
            }
        }
        System.Threading.Thread.Sleep(5000);
    }
    catch (Exception ex)
    {
        // Silently fail - don't interrupt script execution
        // Log($"Error in LogPipe: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   InScope
//
//   Check if ProductCode is in scope for removal
//-------------------------------------------------------------------------------
//Check if ProductCode is in scope
public bool InScope(string sProductCode)
{
    const string OFFICEID = "0000000FF1CE}";
    bool fInScope = false;

    try
    {
        //LogOnly($"Now checking scope of: {sProductCode}");
        if (sProductCode != null && sProductCode.Length == 38)
        {
            //LogOnly("GUID length validated to be 38 characters");
            string sProd = sProductCode.ToUpper();
            if (sProd.EndsWith(OFFICEID))
            {
                //LogOnly($"Pattern matches {OFFICEID}");
                if (sProd.Length >= 6)
                {
                    string versionStr = sProd.Substring(3, 2);
                    if (int.TryParse(versionStr, out int versionMajor) && versionMajor > 14)
                    {
                        //LogOnly("VersionMajor confirmed to be > 14");
                        if (sProd.Length >= 15)
                        {
                            string skuFilter = sProd.Substring(10, 4);
                            switch (skuFilter)
                            {
                                case "007E":
                                case "008F":
                                case "008C":
                                case "24E1":
                                case "237A":
                                case "00DD":
                                    //LogOnly("SKUFilter matches scope");
                                    fInScope = true;
                                    break;
                                default:
                                    //LogOnly($"SKU {skuFilter} doesn't match known integration products scope");
                                    break;
                            }
                        }
                    }
                }
            }
            // Microsoft Online Services Sign-in Assistant (x64 ship and x86 ship)
            if (sProd == "{6C1ADE97-24E1-4AE4-AEDD-86D3A209CE60}") fInScope = true;
            if (sProd == "{9520DDEB-237A-41DB-AA20-F2EF2360DCEB}") fInScope = true;
            if (!string.IsNullOrEmpty(sPackageGuid) && sProd == sPackageGuid.ToUpper()) fInScope = true;
            if (sProd == "{9AC08E99-230B-47E8-9721-4577B7F124EA}") fInScope = true;
        }
    }
    catch
    {
        // Silently fail
    }

    return fInScope;
}

//-------------------------------------------------------------------------------
//   CheckDelete
//
//   Check a ProductCode is known to stay installed
//-------------------------------------------------------------------------------
public bool CheckDelete(string sProductCode)
{
    // ensure valid GUID length
    if (sProductCode == null || sProductCode.Length != 38) return false;
    // only care if it's in the expected ProductCode pattern
    if (!InScope(sProductCode)) return false;
    // check if it's a known product that should be kept
    if (dicKeepSku.ContainsKey(sProductCode.ToUpper())) return false;

    return true;
}

//-------------------------------------------------------------------------------
//   DeleteService
//
//   Delete a service
//-------------------------------------------------------------------------------
//Delete a service
public void DeleteService(string sName)
{
    try
    {
        string sStates = "STARTED;RUNNING";
        string sQuery = $"SELECT * FROM Win32_Service WHERE Name='{sName}'";
        var scope = new ManagementScope(@"\\.\root\cimv2");
        var query = new ObjectQuery(sQuery);
        using (var searcher = new ManagementObjectSearcher(scope, query))
        {
            ManagementObjectCollection services = searcher.Get();

            // stop and delete the service
            foreach (ManagementObject srvc in services)
            {
                string displayName = srvc["DisplayName"]?.ToString() ?? "";
                string state = srvc["State"]?.ToString() ?? "";
                Log($"   Found service {sName} ({displayName}) in state {state}");

                // get the process name
                string pathName = srvc["PathName"]?.ToString() ?? "";
                int lastSlash = pathName.LastIndexOf('\\');
                string sProcessName = lastSlash >= 0 ? pathName.Substring(lastSlash + 1).Trim('"').Trim() : pathName.Trim('"').Trim();

                // stop the service
                if (sStates.Contains(state.ToUpper()))
                {
                    ManagementBaseObject result = srvc.InvokeMethod("StopService", null);
                    uint iRet = result != null ? Convert.ToUInt32(result["ReturnValue"]) : 999;
                    LogOnly($" attempt to stop service {sName} returned: {iRet}");
                }

                // ensure no more instances of the service are running
                string processQuery = $"SELECT * FROM Win32_Process WHERE Name='{sProcessName}'";
                var processQueryObj = new ObjectQuery(processQuery);
                using (var processSearcher = new ManagementObjectSearcher(scope, processQueryObj))
                {
                    foreach (ManagementObject process in processSearcher.Get())
                    {
                        process.InvokeMethod("Terminate", null);
                    }
                }

                if (fDetectOnly)
                {
                    Log($"   Not deleting service {sName} in preview mode");
                    return;
                }

                ManagementBaseObject deleteResult = srvc.InvokeMethod("Delete", null);
                uint deleteRet = deleteResult != null ? Convert.ToUInt32(deleteResult["ReturnValue"]) : 999;
                Log($"   Delete service {sName} returned: {deleteRet}");
            }

            // check if service got deleted
            services = searcher.Get();
            foreach (ManagementObject srvc in services)
            {
                // failed to delete service. retry with 'sc' command
                Log("Delete service " + sName + " failed.");
                Log("Retry delete using 'SC' command");
                string sCmd = "sc delete " + sName;
                if (!fDetectOnly)
                {
                    int iRet = RunCommand(sCmd, true);
                }
            }
        }
    }
    catch (Exception ex)
    {
        Log($"Error in DeleteService: {ex.Message}");
        CheckError("DeleteService");
    }
}


//-------------------------------------------------------------------------------
//   SetupRetVal
//
//   Translation for known uninstall return values
//-------------------------------------------------------------------------------
public string SetupRetVal(int RetVal)
{
    switch (RetVal)
    {
        case 0: return "Success";
        //msiexec return values
        case 1259: return "APPHELP_BLOCK";
        case 1601: return "INSTALL_SERVICE_FAILURE";
        case 1602: return "INSTALL_USEREXIT";
        case 1603: return "INSTALL_FAILURE";
        case 1604: return "INSTALL_SUSPEND";
        case 1605: return "UNKNOWN_PRODUCT";
        case 1606: return "UNKNOWN_FEATURE";
        case 1607: return "UNKNOWN_COMPONENT";
        case 1608: return "UNKNOWN_PROPERTY";
        case 1609: return "INVALID_HANDLE_STATE";
        case 1610: return "BAD_CONFIGURATION";
        case 1611: return "INDEX_ABSENT";
        case 1612: return "INSTALL_SOURCE_ABSENT";
        case 1613: return "INSTALL_PACKAGE_VERSION";
        case 1614: return "PRODUCT_UNINSTALLED";
        case 1615: return "BAD_QUERY_SYNTAX";
        case 1616: return "INVALID_FIELD";
        case 1618: return "INSTALL_ALREADY_RUNNING";
        case 1619: return "INSTALL_PACKAGE_OPEN_FAILED";
        case 1620: return "INSTALL_PACKAGE_INVALID";
        case 1621: return "INSTALL_UI_FAILURE";
        case 1622: return "INSTALL_LOG_FAILURE";
        case 1623: return "INSTALL_LANGUAGE_UNSUPPORTED";
        case 1624: return "INSTALL_TRANSFORM_FAILURE";
        case 1625: return "INSTALL_PACKAGE_REJECTED";
        case 1626: return "FUNCTION_NOT_CALLED";
        case 1627: return "FUNCTION_FAILED";
        case 1628: return "INVALID_TABLE";
        case 1629: return "DATATYPE_MISMATCH";
        case 1630: return "UNSUPPORTED_TYPE";
        case 1631: return "CREATE_FAILED";
        case 1632: return "INSTALL_TEMP_UNWRITABLE";
        case 1633: return "INSTALL_PLATFORM_UNSUPPORTED";
        case 1634: return "INSTALL_NOTUSED";
        case 1635: return "PATCH_PACKAGE_OPEN_FAILED";
        case 1636: return "PATCH_PACKAGE_INVALID";
        case 1637: return "PATCH_PACKAGE_UNSUPPORTED";
        case 1638: return "PRODUCT_VERSION";
        case 1639: return "INVALID_COMMAND_LINE";
        case 1640: return "INSTALL_REMOTE_DISALLOWED";
        case 1641: return "SUCCESS_REBOOT_INITIATED";
        case 1642: return "PATCH_TARGET_NOT_FOUND";
        case 1643: return "PATCH_PACKAGE_REJECTED";
        case 1644: return "INSTALL_TRANSFORM_REJECTED";
        case 1645: return "INSTALL_REMOTE_PROHIBITED";
        case 1646: return "PATCH_REMOVAL_UNSUPPORTED";
        case 1647: return "UNKNOWN_PATCH";
        case 1648: return "PATCH_NO_SEQUENCE";
        case 1649: return "PATCH_REMOVAL_DISALLOWED";
        case 1650: return "INVALID_PATCH_XML";
        case 3010: return "SUCCESS_REBOOT_REQUIRED";
        default: return "Unknown Return Value";
    }
}

//-------------------------------------------------------------------------------
//   DeleteFile
//
//   Wrapper to delete a file
//-------------------------------------------------------------------------------
public void DeleteFile(string sFile)
{
    try
    {
        string sDelFile;
        bool fKeep;

        fKeep = dicKeepFolder.ContainsKey(sFile.ToLower());
        if (!fKeep && f64) fKeep = dicKeepFolder.ContainsKey(Wow64Folder(sFile).ToLower());

        if (fKeep)
        {
            LogOnly($"Disallowing the delete of still required keypath element: {sFile}");
            if (!fForce) return;
        }

        // check on forced delete
        if (fKeep)
        {
            LogOnly($"Enforced delete of still required keypath element: {sFile}");
            LogOnly("   Remaining applications will need a repair!");
        }

        if (File.Exists(sFile))
        {
            sDelFile = sFile;
        }
        else if (f64 && File.Exists(Wow64Folder(sFile)))
        {
            sDelFile = Wow64Folder(sFile);
        }
        else
        {
            LogOnly($"Path not found. Cannot delete file: {sFile}");
            return;
        }

        if (!fDetectOnly)
        {
            LogOnly($"Delete file: {sDelFile}");
            FileInfo fileInfo = new FileInfo(sDelFile);
            // ensure read-only flag is not set
            if ((fileInfo.Attributes & FileAttributes.ReadOnly) != 0)
            {
                fileInfo.Attributes = fileInfo.Attributes & ~FileAttributes.ReadOnly;
            }
            // add folder to empty folder cleanup list
            string parentFolder = fileInfo.DirectoryName;
            if (!dicDelFolder.ContainsKey(parentFolder))
            {
                dicDelFolder.Add(parentFolder, parentFolder);
            }
            // delete the file
            try
            {
                File.Delete(sDelFile);
            }
            catch (Exception ex)
            {
                CheckError("DeleteFile");
                Log($"Error deleting file: {ex.Message}");
                // schedule file for delete on next reboot
                ScheduleDeleteFile(sDelFile);
            }
        }
        else
        {
            LogOnly($"Preview mode. Disallowing delete for file: {sDelFile}");
        }
    }
    catch (Exception ex)
    {
        Log($"Error in DeleteFile: {ex.Message}");
        CheckError("DeleteFile");
    }
}

//-------------------------------------------------------------------------------
//   DeleteFolder
//
//   Wrapper to delete a folder
//-------------------------------------------------------------------------------
public void DeleteFolder(string sFolder)
{
    try
    {
        DirectoryInfo fld;
        string sDelFolder, sCmd;
        bool fKeep;

        // ensure trailing "\"
        // trailing \ is required for dicKeepFolder comparisons
        if (!sFolder.EndsWith("\\")) sFolder = sFolder + "\\";
        while (sFolder.Contains("\\\\"))
        {
            sFolder = sFolder.Replace("\\\\", "\\");
        }

        // prevent delete of folders that are known to be still required
        fKeep = dicKeepFolder.ContainsKey(sFolder.ToLower());
        if (!fKeep && f64) fKeep = dicKeepFolder.ContainsKey(Wow64Folder(sFolder).ToLower());

        if (fKeep)
        {
            LogOnly($"Disallowing the delete of still required keypath element: {sFolder}");
            if (!fForce) return;
        }

        // check on forced delete
        if (fKeep)
        {
            LogOnly($"Enforced delete of still required keypath element: {sFolder}");
            LogOnly("   Remaining applications will need a repair!");
        }

        // strip trailing "\"
        if (sFolder.Length > 1)
        {
            sFolder = sFolder.Substring(0, sFolder.Length - 1);
        }

        if (Directory.Exists(sFolder))
        {
            sDelFolder = sFolder;
        }
        else if (f64 && Directory.Exists(Wow64Folder(sFolder)))
        {
            sDelFolder = Wow64Folder(sFolder);
        }
        else
        {
            LogOnly($"Path not found. Cannot delete folder: {sFolder}");
            return;
        }

        if (!fDetectOnly)
        {
            LogOnly($"Delete folder: {sDelFolder}");
            DirectoryInfo folder = new DirectoryInfo(sDelFolder);
            // ensure to remove read only flag
            if ((folder.Attributes & FileAttributes.ReadOnly) != 0)
            {
                folder.Attributes = folder.Attributes & ~FileAttributes.ReadOnly;
            }
            // add to empty folder cleanup list
            if (!dicDelFolder.ContainsKey(folder.FullName))
            {
                dicDelFolder.Add(folder.FullName, folder.FullName);
            }
            // delete the folder
            // for performance reasons try 'rd' first
            sCmd = $"cmd.exe /c rd /s \"{sDelFolder}\" /q";
            RunCommand(sCmd, true);
            if (!Directory.Exists(sDelFolder)) return;

            // rd didn't work check with Directory.Delete
            try
            {
                Directory.Delete(sDelFolder, true);
            }
            catch (UnauthorizedAccessException)
            {
                // Access Denied
                // Retry after closing running processes
                CheckError("DeleteFolder");
                if (!fRerun)
                {
                    CloseOfficeApps();
                    // attempt 'rd' command
                    LogOnly("   Attempt to remove with 'rd' command");
                    sCmd = $"cmd.exe /c rd /s \"{sDelFolder}\" /q";
                    RunCommand(sCmd, true);
                    if (!Directory.Exists(sDelFolder)) return;
                }
            }
            catch (DirectoryNotFoundException)
            {
                // check on invalid path length issues
                // attempt 'rd' command
                CheckError("DeleteFolder");
                LogOnly("   Attempt to remove with 'rd' command");
                sCmd = $"cmd.exe /c rd /s \"{sDelFolder}\" /q";
                RunCommand(sCmd, true);
                if (!Directory.Exists(sDelFolder)) return;
            }
            catch (Exception ex)
            {
                // still failed!
                Log($"   Failed to delete folder: {sDelFolder}");
                CheckError("DeleteFolder");

                // try to delete as many folder contents as possible
                // before the recursive error handling is called
                folder = new DirectoryInfo(sDelFolder);
                foreach (DirectoryInfo subFolder in folder.GetDirectories())
                {
                    fld = subFolder;
                    sCmd = $"cmd.exe /c rd /s \"{fld.FullName}\" /q";
                    RunCommand(sCmd, true);
                }
                if (folder.GetDirectories().Length > 0)
                {
                    DirectoryInfo lastFolder = folder.GetDirectories()[folder.GetDirectories().Length - 1];
                    sCmd = $"cmd.exe /c del \"{lastFolder.FullName}\\*.*\"";
                    RunCommand(sCmd, true);
                }

                // schedule an additional run of the tool after reboot
                if (!fRerun) Rerun();

                // schedule folder for delete on next reboot
                ScheduleDeleteFolder(sDelFolder);
            }
        }
        else
        {
            LogOnly($"Preview mode. Disallowing delete of folder: {sDelFolder}");
        }
    }
    catch (Exception ex)
    {
        Log($"Error in DeleteFolder: {ex.Message}");
        CheckError("DeleteFolder");
    }
}

public void DeleteFolder_WMI(string sFolder)
{
    try
    {
        string sWqlFolder = sFolder.Replace("\\", "\\\\");
        var scope = new ManagementScope(@"\\.\root\cimv2");
        string queryString = $"SELECT * FROM Win32_Directory WHERE Name='{sWqlFolder}'";
        var query = new ObjectQuery(queryString);
        using (var searcher = new ManagementObjectSearcher(scope, query))
        {
            foreach (ManagementObject folder in searcher.Get())
            {
                ManagementBaseObject result = folder.InvokeMethod("Delete", null);
                uint iRet = result != null ? Convert.ToUInt32(result["ReturnValue"]) : 999;
                LogOnly($"   Delete (wmi) for folder {sFolder} returned: {iRet}");
            }
        }
    }
    catch (Exception ex)
    {
        Log($"Error in DeleteFolder_WMI: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   Wow64Folder
//
//   Returns the WOW folder structure to handle folder-path operations on
//   64 bit environments
//-------------------------------------------------------------------------------
public string Wow64Folder(string sFolder)
{
    string system32Path = sWinDir + "\\System32";
    if (sFolder.StartsWith(system32Path, StringComparison.OrdinalIgnoreCase))
    {
        return sWinDir + "\\syswow64" + sFolder.Substring(system32Path.Length);
    }
    else if (sFolder.StartsWith(sProgramFiles, StringComparison.OrdinalIgnoreCase))
    {
        return sProgramFilesX86 + sFolder.Substring(sProgramFiles.Length);
    }
    else
    {
        return "?"; //Return invalid string to ensure the folder cannot exist
    }
}

//-------------------------------------------------------------------------------
//   ScheduleDeleteFile
//
//   Adds a file to the list of items to delete on reboot
//-------------------------------------------------------------------------------
public void ScheduleDeleteFile(string sFile)
{
    if (!dicDelInUse.ContainsKey(sFile))
    {
        dicDelInUse.Add(sFile, sFile);
        LogOnly($"Add file in use for delete on reboot: {sFile}");
        fRebootRequired = true;
        SetError(Constants.ERROR_REBOOT_REQUIRED);
    }
}

//-------------------------------------------------------------------------------
//   ScheduleDeleteFolder
//
//   Recursively adds a folder and its contents to the list of
//   items to delete on reboot
//-------------------------------------------------------------------------------
public void ScheduleDeleteFolder(string sFolder)
{
    try
    {
        DirectoryInfo oFolder = new DirectoryInfo(sFolder);
        // exclude hidden system folders
        if ((oFolder.Attributes & (FileAttributes.Hidden | FileAttributes.System)) != 0) return;

        foreach (DirectoryInfo fld in oFolder.GetDirectories())
        {
            DeleteFolder(fld.FullName);
        }
        foreach (FileInfo file in oFolder.GetFiles())
        {
            DeleteFile(file.FullName);
        }
        if (!dicDelInUse.ContainsKey(oFolder.FullName))
        {
            dicDelInUse.Add(oFolder.FullName, "");
            LogOnly($"Add folder for delete on reboot: {oFolder.FullName}");
            fRebootRequired = true;
            SetError(Constants.ERROR_REBOOT_REQUIRED);
        }
    }
    catch (Exception ex)
    {
        Log($"Error in ScheduleDeleteFolder: {ex.Message}");
    }
}


//-------------------------------------------------------------------------------
//   ScheduleDeleteEx
//
//   Schedules the delete of files/folders in use on next reboot by adding
//   affected files/folders to the PendingFileRenameOperations registry entry
//-------------------------------------------------------------------------------
public void ScheduleDeleteEx()
{
    try
    {
        int hDefKey = Constants.HKLM;
        string sKeyName = @"SYSTEM\CurrentControlSet\Control\Session Manager";
        string sValueName = "PendingFileRenameOperations";

        LogH2($"Add {dicDelInUse.Count} PendingFileRenameOperations");
        List<string> arrData = new List<string>();

        if (RegValExists(hDefKey, sKeyName, sValueName))
        {
            string existingValue;
            if (RegReadValue(hDefKey, sKeyName, sValueName, out existingValue, "REG_MULTI_SZ"))
            {
                arrData.AddRange(existingValue.Split('\r'));
            }
        }

        foreach (string key in dicDelInUse.Keys)
        {
            LogOnly($"   {key}");
            arrData.Add("\\??\\" + key);
            arrData.Add("");
        }

        // Write REG_MULTI_SZ value - requires registry API
        Microsoft.Win32.RegistryKey baseKey = Microsoft.Win32.Registry.LocalMachine;
        Microsoft.Win32.RegistryKey key = baseKey.OpenSubKey(sKeyName, true);
        if (key != null)
        {
            key.SetValue(sValueName, arrData.ToArray(), Microsoft.Win32.RegistryValueKind.MultiString);
            key.Close();
        }
        baseKey.Close();
    }
    catch (Exception ex)
    {
        Log($"Error in ScheduleDeleteEx: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   DeleteEmptyFolder
//
//   Deletes an individual folder structure if empty
//-------------------------------------------------------------------------------
public void DeleteEmptyFolder(string sFolder)
{
    // cosmetic task don't fail on error
    try
    {
        if (Directory.Exists(sFolder))
        {
            DirectoryInfo folder = new DirectoryInfo(sFolder);
            if (folder.GetDirectories().Length == 0 && folder.GetFiles().Length == 0)
            {
                SmartDeleteFolder(sFolder);
            }
        }
    }
    catch (Exception ex)
    {
        CheckError("DeleteEmptyFolder");
        Log($"Error in DeleteEmptyFolder: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   DeleteEmptyFolders
//
//   Delete an empty folder structure
//-------------------------------------------------------------------------------
public void DeleteEmptyFolders()
{
    // cosmetic task don't fail on error
    try
    {
        DeleteEmptyFolder(sCommonProgramFiles + "\\Microsoft Shared\\Office15");
        DeleteEmptyFolder(sCommonProgramFiles + "\\Microsoft Shared\\Office16");
        DeleteEmptyFolder(sCommonProgramFiles + "\\Microsoft Shared\\");
        DeleteEmptyFolder(sProgramFiles + "\\Microsoft Office\\Office15");
        DeleteEmptyFolder(sProgramFiles + "\\Microsoft Office\\Office16");

        foreach (string sFolder in dicDelFolder.Keys)
        {
            if (Directory.Exists(sFolder))
            {
                DirectoryInfo folder = new DirectoryInfo(sFolder);
                if (folder.GetDirectories().Length == 0 && folder.GetFiles().Length == 0)
                {
                    SmartDeleteFolder(sFolder);
                }
            }
        }
    }
    catch (Exception ex)
    {
        CheckError("DeleteEmptyFolders");
        Log($"Error in DeleteEmptyFolders: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   SmartDeleteFolder
//
//   Wrapper to delete a folder and the empty parent folder structure
//-------------------------------------------------------------------------------
public void SmartDeleteFolder(string sFolder)
{
    string sDelFolder;

    if (Directory.Exists(sFolder))
    {
        sDelFolder = sFolder;
    }
    else if (f64 && Directory.Exists(Wow64Folder(sFolder)))
    {
        sDelFolder = Wow64Folder(sFolder);
    }
    else
    {
        return;
    }

    if (!fDetectOnly)
    {
        LogOnly($"Request SmartDelete for folder: {sDelFolder}");
        SmartDeleteFolderEx(sDelFolder);
    }
    else
    {
        LogOnly($"Preview mode. Disallowing SmartDelete request for folder: {sDelFolder}");
    }
}

//-------------------------------------------------------------------------------
//   SmartDeleteFolderEx
//
//   Executes the folder delete operation(s)
//-------------------------------------------------------------------------------
public void SmartDeleteFolderEx(string sFolder)
{
    try
    {
        DeleteFolder(sFolder);
        CheckError("SmartDeleteFolderEx");
        DirectoryInfo parentFolder = Directory.GetParent(sFolder);
        if (parentFolder != null && parentFolder.GetDirectories().Length == 0 && parentFolder.GetFiles().Length == 0)
        {
            SmartDeleteFolderEx(parentFolder.FullName);
        }
    }
    catch (Exception ex)
    {
        Log($"Error in SmartDeleteFolderEx: {ex.Message}");
        CheckError("SmartDeleteFolderEx");
    }
}

//-------------------------------------------------------------------------------
//   RestoreExplorer
//
//   Ensure Windows Explorer is restarted if needed
//-------------------------------------------------------------------------------
public void RestoreExplorer()
{
    //Non critical routine. Don't fail on error
    try
    {
        System.Threading.Thread.Sleep(1000);
        var scope = new ManagementScope(@"\\.\root\cimv2");
        string queryString = "SELECT * FROM Win32_Process WHERE Name='explorer.exe'";
        var query = new ObjectQuery(queryString);
        using (var searcher = new ManagementObjectSearcher(scope, query))
        {
            if (searcher.Get().Count < 1)
            {
                Process.Start("explorer.exe");
                //To handle this in case of System context, schedule and run as interactive task
                RunCommand("SCHTASKS /Create /TN OffScrEx /TR explorer /SC ONCE /ST 12:00 /IT", true);
                RunCommand("SCHTASKS /Run /TN OffScrEx", true);
                RunCommand("SCHTASKS /Delete /TN OffScrEx /F", false);
            }
        }
    }
    catch (Exception ex)
    {
        Log($"Error in RestoreExplorer: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   MyJoin
//
//   Replacement function to the internal Join function to prevent failures
//   that were seen in some instances
//-------------------------------------------------------------------------------
public string MyJoin(string[] arrToJoin, string sSeparator)
{
    string sJoined = "";
    if (arrToJoin != null)
    {
        foreach (string item in arrToJoin)
        {
            sJoined = sJoined + item + sSeparator;
        }
    }
    if (sJoined.Length > 1) sJoined = sJoined.Substring(0, sJoined.Length - 1);
    return sJoined;
}

//-------------------------------------------------------------------------------
//   Rerun
//
//   Flag need for reboot and schedule autorun to run the tool again on reboot.
//-------------------------------------------------------------------------------
public void Rerun()
{
    try
    {
        string sValue;

        // check if Rerun has already been called
        if (fRerun) return;

        // set Rerun flag
        fRerun = true;

        // check if the previous run already initiated the Rerun
        if (RegReadValue(Constants.HKCU, @"SOFTWARE\Microsoft\Office\15.0\CleanC2R", "Rerun", out sValue, "REG_DWORD"))
        {
            // Rerun has already been tried
            LogH2("Error: Removal failed");
            SetError(Constants.ERROR_DCAF_FAILURE);
            return;
        }

        fRebootRequired = true;
        SetError(Constants.ERROR_REBOOT_REQUIRED);
        SetError(Constants.ERROR_INCOMPLETE);

        // cache the script to the local scrub folder
        string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string destPath = Path.Combine(sScrubDir, Constants.SCRIPTFILE);
        File.Copy(exePath, destPath, true);

        // Create registry keys and set value
        Microsoft.Win32.RegistryKey baseKey = Microsoft.Win32.Registry.LocalMachine;
        Microsoft.Win32.RegistryKey key = baseKey.CreateSubKey(@"SOFTWARE\Microsoft\Office\15.0\CleanC2R", true);
        if (key != null)
        {
            key.SetValue("Rerun", 1, Microsoft.Win32.RegistryValueKind.DWord);
            key.Close();
        }
        baseKey.Close();

        fSetRunOnce = true;
        //    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", true);
        //    if (key != null)
        //    {
        //        key.SetValue("CleanC2R", $"cscript.exe \"{Path.Combine(sScrubDir, Constants.SCRIPTFILE)}\"", Microsoft.Win32.RegistryValueKind.String);
        //        key.Close();
        //    }
    }
    catch (Exception ex)
    {
        Log($"Error in Rerun: {ex.Message}");
    }
}

//-------------------------------------------------------------------------------
//   SetRunOnce
//
//   Create a RunOnce entry to resume setup after a reboot
//-------------------------------------------------------------------------------
public void SetRunOnce()
{
    try
    {
        Microsoft.Win32.RegistryKey baseKey = Microsoft.Win32.Registry.LocalMachine;
        Microsoft.Win32.RegistryKey key = baseKey.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", true);
        if (key != null)
        {
            string sValue = $"cscript.exe \"{Path.Combine(sScrubDir, Constants.SCRIPTFILE)}\" /NoElevate /Relaunched";
            key.SetValue("O15CleanUp", sValue, Microsoft.Win32.RegistryValueKind.String);
            key.Close();
        }
        baseKey.Close();
    }
    catch (Exception ex)
    {
        Log($"Error in SetRunOnce: {ex.Message}");
    }
}
