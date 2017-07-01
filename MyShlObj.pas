unit MyShlObj;

interface

uses
  SysUtils, Classes, Controls, Windows, Messages, ExtCtrls, Graphics;

 type
   THUMBBUTTON = record
    dwMask: DWORD;
    iId: UINT;
    iBitmap: UINT;
    hIcon: HICON;
    szTip: packed array[0..259] of WCHAR;
    dwFlags: DWORD;
  end;
  {$EXTERNALSYM THUMBBUTTON}

  tagTHUMBBUTTON = THUMBBUTTON;
  {$EXTERNALSYM tagTHUMBBUTTON}

   TThumbButton = THUMBBUTTON;
   PThumbButton = ^TThumbButton;

const
// SID
  SID_ITaskbarList = '{56FDF342-FD6D-11D0-958A-006097C9A090}';
  SID_ITaskbarList2 = '{602D4995-B13A-429B-A66E-1935E44F4317}';
  SID_ITaskbarList3 = '{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}';
  SID_ITaskbarList4 = '{C43DC798-95D1-4BEA-9030-BB99E2983A1A}';
// GUID
  CLSID_TaskbarList: TGUID  = '{56FDF344-FD6D-11d0-958A-006097C9A090}';
  {$EXTERNALSYM CLSID_TaskbarList}
  IID_ITaskbarList3: TGUID = SID_ITaskbarList3;
  {$EXTERNALSYM IID_ITaskbarList3}

 {$EXTERNALSYM WM_DWMSENDICONICLIVEPREVIEWBITMAP}
  WM_DWMSENDICONICLIVEPREVIEWBITMAP = $0326;

  {$EXTERNALSYM WM_DWMSENDICONICTHUMBNAIL}
  WM_DWMSENDICONICTHUMBNAIL       = $0323;

type
 {$EXTERNALSYM HIMAGELIST}
  HIMAGELIST = THandle;

// interface desribe //

   { interface ITaskbarList }
  ITaskbarList = interface(IUnknown)
    [SID_ITaskbarList]
    function HrInit: HRESULT; stdcall;
    function AddTab(hwnd: HWND): HRESULT; stdcall;
    function DeleteTab(hwnd: HWND): HRESULT; stdcall;
    function ActivateTab(hwnd: HWND): HRESULT; stdcall;
    function SetActiveAlt(hwnd: HWND): HRESULT; stdcall;
  end;

 { interface ITaskbarList2 }
  ITaskbarList2 = interface(ITaskbarList)
    [SID_ITaskbarList2]
    function MarkFullscreenWindow(hwnd: HWND; fFullscreen: BOOL): HRESULT; stdcall;
  end;

  ITaskbarList3 = interface(ITaskbarList2)
    [SID_ITaskbarList3]
    function SetProgressValue(hwnd: HWND; ullCompleted: ULONGLONG;
      ullTotal: ULONGLONG): HRESULT; stdcall;
    function SetProgressState(hwnd: HWND; tbpFlags: Integer): HRESULT; stdcall;
    function RegisterTab(hwndTab: HWND; hwndMDI: HWND): HRESULT; stdcall;
    function UnregisterTab(hwndTab: HWND): HRESULT; stdcall;
    function SetTabOrder(hwndTab: HWND; hwndInsertBefore: HWND): HRESULT; stdcall;
    function SetTabActive(hwndTab: HWND; hwndMDI: HWND;
      tbatFlags: Integer): HRESULT; stdcall;
    function ThumbBarAddButtons(hwnd: HWND; cButtons: UINT;
      pButton: PThumbButton): HRESULT; stdcall;
    function ThumbBarUpdateButtons(hwnd: HWND; cButtons: UINT;
      pButton: PThumbButton): HRESULT; stdcall;
    function ThumbBarSetImageList(hwnd: HWND; himl: HIMAGELIST): HRESULT; stdcall;
    function SetOverlayIcon(hwnd: HWND; hIcon: HICON;
      pszDescription: LPCWSTR): HRESULT; stdcall;
    function SetThumbnailTooltip(hwnd: HWND; pszTip: LPCWSTR): HRESULT; stdcall;
    function SetThumbnailClip(hwnd: HWND; var prcClip: TRect): HRESULT; stdcall;
  end;

// THUMBBUTTON flags
const
  THBF_ENABLED        =  $0000;
  {$EXTERNALSYM THBF_ENABLED}
  THBF_DISABLED       =  $0001;
  {$EXTERNALSYM THBF_DISABLED}
  THBF_DISMISSONCLICK =  $0002;
  {$EXTERNALSYM THBF_DISMISSONCLICK}
  THBF_NOBACKGROUND   =  $0004;
  {$EXTERNALSYM THBF_NOBACKGROUND}
  THBF_HIDDEN         =  $0008;
  {$EXTERNALSYM THBF_HIDDEN}
  THBF_NONINTERACTIVE = $10;
  {$EXTERNALSYM THBF_NONINTERACTIVE}
// THUMBBUTTON mask
  THB_BITMAP          =  $0001;
  {$EXTERNALSYM THB_BITMAP}
  THB_ICON            =  $0002;
  {$EXTERNALSYM THB_ICON}
  THB_TOOLTIP         =  $0004;
  {$EXTERNALSYM THB_TOOLTIP}
  THB_FLAGS           =  $0008;
  {$EXTERNALSYM THB_FLAGS}
  THBN_CLICKED        =  $1800;
  {$EXTERNALSYM THBN_CLICKED}

  TBPF_NOPROGRESS    = 0;
  {$EXTERNALSYM TBPF_NOPROGRESS}
  TBPF_INDETERMINATE = $1;
  {$EXTERNALSYM TBPF_INDETERMINATE}
  TBPF_NORMAL        = $2;
  {$EXTERNALSYM TBPF_NORMAL}
  TBPF_ERROR         = $4;
  {$EXTERNALSYM TBPF_ERROR}
  TBPF_PAUSED        = $8;
  {$EXTERNALSYM TBPF_PAUSED}

implementation

end.
