# clsPresenceMon
*A self-contained, threaded class to monitor user presence in several ways*


This is a simple class that monitors for 3 events: The monitor state (off, on, or dimmed), the system's 'user present' message, and, if you're on a laptop, the lid status (open, close). It does this in an entirely self contained class-- the complication with this arises because these events are sent in the form of a `WM_POWERBROADCAST` message, requring a window to receive it. But I didn't want to limit the class to graphical apps and just subclass a form hWnd, although for the sake of completeness the demo app shows how to do it that way too. So the class creates it's own entirely custom hidden window, in a separate thread to avoid the message loop blocking the Form's message loop if present. 

![image](https://github.com/fafalone/clsPresenceMon/assets/7834493/62c830f3-bef1-40e1-a112-fc0d2db5882c)

## Setting things up

The class takes advantage of twinBASIC's parameterized constructors to specify the available arguments right in the new keyword; here that's optional arguments for which events to raise and the hInstance to register under (`App.hInstance` 99% of the time; if you omit it, it will use `GetModuleHandleW()`, which returns the same value as `App.hInstance` without creating a dependency on WinNativeForms). Then we get to one of the great pleasures of tB: Calling `CreateThread` without any elaborate hacks neccessary, as cool as they are in VB6. The ThreadProc itself simply calls the rest of the code.

```
    Sub New(Optional ByVal dwNotifyMask As CPMonEventNotify = CPMEN_ALL, Optional ByVal hInst As LongPtr)
        m_hInst = If(hInst = 0, GetModuleHandleW(), hInst)
        If dwNotifyMask = CPMEN_ERROR Then Exit Sub
        m_Mask = dwNotifyMask
        
        tConfig.hInst = m_hInst
        tConfig.Mask = m_Mask
        
        m_hThread = CreateThread(ByVal 0, 0, AddressOf CPMonProc, tConfig, 0, m_idThread)
            
    End Sub
```

`RegisterClassEx` and `CreateWindowEx`are used in just a normal routine to create a hidden window, with the WndProc being a function within the class thanks to tB's supporting `AddressOf` here. The last important part of setup is that we have to register for the events we want; they're all delivered as `PBT_POWERSETTINGCHANGE` messages. We need to keep a handle for each as a class variable, but registration is straightforward; tbShellLib provides the API and GUIDs:

```
    Private Function RegisterEvents() As Boolean
        m_hEventM = RegisterPowerSettingNotification(m_hWnd, GUID_SESSION_DISPLAY_STATUS, DEVICE_NOTIFY_WINDOW_HANDLE)
        m_hEventP = RegisterPowerSettingNotification(m_hWnd, GUID_SESSION_USER_PRESENCE, DEVICE_NOTIFY_WINDOW_HANDLE)
        m_hEventL = RegisterPowerSettingNotification(m_hWnd, GUID_LIDSWITCH_STATE_CHANGE, DEVICE_NOTIFY_WINDOW_HANDLE)
        If m_hEventM Then Return True
    End Function
```
## Receiving the messages

As already mentioned, they're delivered by uMsg `WM_POWERBROADCAST` with wParam `PBT_POWERSETTINGCHANGE`; we know from MSDN the lParam then points to a `POWERBROADCAST_SETTING` UDT. But we run into a problem:

```
typedef struct {
  GUID  PowerSetting;
  DWORD DataLength;
  UCHAR Data[1];
} POWERBROADCAST_SETTING, *PPOWERBROADCAST_SETTING;
```

That's a C-style array, not a `SAFEARRAY`; the data immediately follows in memory, where variable length arrays in tB (currently) can only be a `SAFEARRAY` with an entirely different structure. The solution tbShellLib uses is a bit clunky;

```
[ Description ("WARNING: You can't use this directly due to the SAFEARRAY. To receive, fill the first 20 bytes, then the data in the array. To send, create a byte buffer excluding the safearray member.") ]
Public Type POWERBROADCAST_SETTING
    PowerSetting As UUID
    DataLength As Long
    Data() As Byte
End Type
```

The idea is to copy the fixed part, then use the length to ReDim the variable part, and do a separate copy into VarPtr(Data(0)). This class takes a shortcut though; since we're only working with a single 4-byte DWORD for the properties we're interested in, it looks like this:

```
            Case WM_POWERBROADCAST
                If wParam = PBT_POWERSETTINGCHANGE Then
                    Dim pSetting As POWERBROADCAST_SETTING
                    CopyMemory pSetting, ByVal lParam, 20
                    If IsEqualGUID(pSetting.PowerSetting, GUID_SESSION_DISPLAY_STATUS) Then
                        Dim pState As MONITOR_DISPLAY_STATE
                        CopyMemory pState, ByVal PointerAdd(lParam, 20), 4
                        Select Case pState
                            Case PowerMonitorOff
                                If (m_Mask And CPMEN_MONITOROFF) Then RaiseEvent MonitorOff()
```

We copy the fixed part so we can check the GUID for which event it is this time, but then just copy 4 bytes starting from the offset of `Data` (a GUID is 16 bytes, plus the 4 byte Long, =20) directly to a variable representing the enum of possible values. Everything besides the last line is provided by tbShellLib, including the generic `PointerAdd`, which safely performs an unsigned addition. Then we just check if the caller wants the event, and raise it. Surprisingly, calling RaiseEvent from out of thread like this worked with no special handling and hasn't crashed yet, and I've left the class running for 6+ hour stretches with many events being raised.

> [!NOTE]
> You'll receive the current status when you first register for events. So it will send 'Monitor on' even though the monitor has not just been turned on. 

## Cleaning up

The class provides a `Destroy()` method to turn off monitoring and get rid of the hidden window. It's required that you call this before trying to set it to `Nothing` if your app plans to keep running (if exiting you can let the system destroy everything). The thread is in it's message loop, so it won't exit util the message loop exits, so to do that, the window needs to be destroyed. You can't call `DestroyWindow` on window in a different thread, so what you do instead is send `WM_CLOSE` with `PostMessage`, and the window destroys itself, and unregisters the events:

```
            Case WM_CLOSE
                DestroyWindow m_hWnd
            
            Case WM_DESTROY
                UnregisterEvents
                PostQuitMessage 0
```

Besides that we give the thread a few seconds to shut down, then unregister our custom window class. After that, everything is cleaned up and the class itself is ready to be destroyed.


```
    Public Sub Destroy()
        If m_hWnd Then PostMessageW(m_hWnd, WM_CLOSE, 0, ByVal 0)
        Dim lRet As WaitForObjOutcomes = WaitForSingleObject(m_hThread, 5000)
        Debug.Print "Wait outcome=" & lRet
        Dim hr As Long = UnregisterClassW(StrPtr(wndClass), m_hInst)
        Debug.Print "Unregister hr=" & hr & ", lastErr=" & Err.LastDllError
    End Sub
```

---
And that's all there is to it! It's all pretty straightforward, but worth writing up since a lot of are new to the whole 'easy multithreading' thing.
