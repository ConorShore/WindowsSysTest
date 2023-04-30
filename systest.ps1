function TestResultWrapper {
    param (
        $TestResultInput
    )
    echo $TestResultInput
}

function TestFinished {
    Stop-Transcript
    Exit
}

function TestFailed {
    TestResultWrapper("Test failed, see previous messages for cause")
    TestFinished

}



Add-Type -TypeDefinition @'

using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {
  // f(), g(), ... are unused COM method slots. Define these if you care
  int f(); int g(); int h(); int i();
  int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
  int j();
  int GetMasterVolumeLevelScalar(out float pfLevel);
  int k(); int l(); int m(); int n();
  int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
  int GetMute(out bool pbMute);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
  int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
  int f(); // Unused
  int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }
public class Audio {
  static IAudioEndpointVolume Vol() {
    var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
    IMMDevice dev = null;
    Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
    IAudioEndpointVolume epv = null;
    var epvid = typeof(IAudioEndpointVolume).GUID;
    Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
    return epv;
  }
  public static float Volume {
    get {float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v;}
    set {Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty));}
  }
  public static bool Mute {
    get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
    set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
  }
}

'@



$transc_dir = $PSScriptRoot + "\log.txt"

Start-Transcript -Path $transc_dir

echo "Before starting this test, make sure the computer is connected to internet"
$areyouready = 'yes'

$areyouready = Read-Host "Is this computer connected to internet? Please enter yes or no [yes]"

if ($areyouready -eq '') {
    $areyouready = 'yes'
}

if($areyouready -ne 'yes') {
    echo "Test not started"
    TestFailed
}

#Collect system data
TestResultWrapper(Get-WmiObject win32_bios)
TestResultWrapper(Get-WMIObject –class Win32_ComputerSystem)

#test internet
echo "Testing connection to internet"
if (Test-Connection www.google.com -quiet) {
    TestResultWrapper("Internet connection working")
} else {
    TestResultWrapper("Internet connection failed")
    TestFailed
}

#install updates
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
Install-Module -Name PSWindowsUpdate -Force
Get-WindowsUpdate -AcceptAll -Install

#Test audio and video by playing youtube video
echo "Testing audio and video playback"
[audio]::Volume = 1
start-process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" https://www.youtube.com/watch?v=gte3BoXKwP0
sleep 20
Stop-Process -Name "msedge" -Force

#Did audio and video play?
$audiook = Read-Host "Did you hear audio playing? Please enter yes or no [yes]"

if ($audiook -eq '') {
    $audiook = 'yes'
}

if($audiook -ne 'yes') {
    TestResultWrapper("Audio test failed")
    TestFailed
} else {
    TestResultWrapper("Audio test passed")
}

$videook = Read-Host "Did the video play ok? Please enter yes or no [yes]"

if ($videook -eq '') {
    $videook = 'yes'
}

if($videook -ne 'yes') {
    TestResultWrapper("Video test failed")
    TestFailed
} else {
    TestResultWrapper("Audio test passed")
}

#Battery info
echo "Checking battery parameters"
echo "Warning: This is very experimnetal, and as such won't cause a test failure"
try {
[string]$DesignCapstr = Get-WmiObject Win32_Battery | select DesignCapacity
[string]$FullChargeCapstr = Get-WmiObject Win32_Battery | select FullChargeCapacity

[int]$DesignCap = [convert]::ToInt32($DesignCapstr, 10)
[int]$FullChargeCap = [convert]::ToInt32($FullChargeCapstr, 10)


$CurrentBatteryEstimate = ($DesignCap/$FullChargeCap)*100
echo $CurrentBatteryEstimate
} catch {
    echo "Something went wrong with battery test"
}

#Check webcam
Enable-WindowsOptionalFeature -Online -FeatureName "MediaPlayback" -All | SELECT Online
start microsoft.windows.camera:
sleep 10
Stop-Process -Name "WindowsCamera" -Force


$webcok = Read-Host "Did the webcam work ok? Please enter yes or no [yes]"

if ($webcok -eq '') {
    $webcok = 'yes'
}

if($webcok -ne 'yes') {
    TestResultWrapper("Webcam test failed")
    TestFailed
} else {
    TestResultWrapper("Webcam test passed")
}

echo "Test Finished"
Read-Host "Press enter to finish"
TestFinished
