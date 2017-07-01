unit WinSevenConnect;

           {************************************************}
           {*                                              *}
           {*           WinSevenConnect  v 4.03            *}
           {*            Author: Dukuy Roman               *}
           {*           (C) SPS Team 2009                  *}
           {*              ICQ: 443-731-743                *}
           {*           Mail: free_sps@yahoo.com           *}
           {*                                              *}
           {************************************************}
interface

uses
  SysUtils, Classes, Controls, Windows, Messages, ExtCtrls, Graphics,
	 ShlObj, ComObj, SHLWAPI, Activex,	Forms,MyShlObj;
type
 TNotifyEventPaint = procedure(Sender: TObject; var ThumbImage: TBitmap) of object;
 TNotifyEvent = procedure(Sender: TObject) of object;
 TTaskState = (Noprogres,Identerminate,Normal,Error,Pause);
 TTaskThumbFlag = (tfDisable, tfDismissionClick, tfNoBackground, tfHidden, tfNonInteractive);
 TFlags = set of TTaskThumbFlag;

const
  THBNCount = 7;
  THBN = 40;

  DWMWA_FORCE_ICONIC_REPRESENTATION = 7;
  DWMWA_HAS_ICONIC_BITMAP = 10;
  DWMWA_DISALLOW_PEEK = 11;

  MSG_Add = 1;
  MSG_Sorry = 'Sorry. Component not`s work, because this operation system is to old!';
  MSG_Double = 'This component already exists!';
  MSG_Owner = 'This component must be placed on TCustomForm!';
  TaskState: array[TTaskState] of Longint =
   (TBPF_NOPROGRESS,TBPF_INDETERMINATE,TBPF_NORMAL,TBPF_ERROR,TBPF_PAUSED);

type
 TThumbBtn = class (TPersistent)
  private
    FVisible: Boolean;
    FSHowHint: Boolean;
    FHint: string;
    FImageIndex: integer;
    FOnChange: TNotifyEvent;
    FOnCreate: TNotifyEvent;
    BaseClass: TComponent;
    FFlags: TFlags;
    procedure SetVisible(const Value: Boolean);
    procedure SetShowHint(const Value: Boolean);
    procedure SetHint(const Value: string);
    procedure SetImageIndex(const Value: Integer);
    procedure SetFlags(const Value: TFlags);
  protected
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
    property OnCreate: TNotifyEvent read FOnCreate write FOnCreate;
  public
    constructor Create(AClass: TComponent);
  published
    property Visible: Boolean read FVisible write SetVisible default False;
    property ShowHint: Boolean read FSHowHint write SetShowHint default False;
    property Hint: string read FHint write SetHint;
    property ImageIndex: Integer read FImageIndex write SetImageIndex default -1;
    property Flags: TFlags read FFlags write SetFlags;
  end;

  THButton = class (TPersistent)
  private
   BaseClass: TComponent;
   FBtn: array[1..THBNCount] of TThumbBtn;
   function GetBtn(Index: Integer): TThumbBtn;
   procedure SetInt(Index: Integer; Value: TThumbBtn);
   procedure AChanged(Sender: TObject);
   procedure ACreate(Sender: TObject);
  public
   constructor Create(AClass: TComponent);
  published
    property THBtn1: TThumbBtn index 1; read GetBtn write SetInt;
    property THBtn2: TThumbBtn index 2; read GetBtn write SetInt;
    property THBtn3: TThumbBtn index 3; read GetBtn write SetInt;
    property THBtn4: TThumbBtn index 4; read GetBtn write SetInt;
    property THBtn5: TThumbBtn index 5; read GetBtn write SetInt;
    property THBtn6: TThumbBtn index 6; read GetBtn write SetInt;
    property THBtn7: TThumbBtn index 7; read GetBtn write SetInt;
 end;

 TClipRect = class (TPersistent)
  private
    FInteger: array[1..4] of Integer;
    function GetInt(Index: Integer): Integer;
    procedure SetInt(Index: Integer; Value: Integer);
  published
    property Left: Integer index 1; read GetInt write SetInt default 0;
    property Top: Integer index 2; read GetInt write SetInt default 0;
    property Right: Integer index 3; read GetInt write SetInt;
    property Buttom: Integer index 4; read GetInt write SetInt;
  end;

  TWinSevenConnect = class(TComponent)
  private
    FVisible: Boolean;
    FUseClipRect: Boolean;
    FShowHint: Boolean;
    FShowThumbImage: Boolean;
    BtnMass: array[1..7] of Boolean;
    FOverIconText : string;
    FToolTip: string;
    FAppId: string;
    FProgress: integer;
    FProgressMax: integer;
    FOverIconIndex: Integer;
    FThumbImage: TBitmap;
    FMainWndProc : Pointer;
    FTaskState: TTaskState;
    FImageList: TImageList;
    FClipRect: TClipRect;
    FThumbButton: THButton;
    Timer: Ttimer;
    Handle: THandle;
    Button: array[1..7] of TThumbButton;
    TaskBarList: ITaskBarList;
    TaskBarList2: ITaskBarList2;
    TaskbarList3: ITaskbarList3;
    FOnBtnClick1: TNotifyEvent;
    FOnBtnClick2: TNotifyEvent;
    FOnBtnClick3: TNotifyEvent;
    FOnBtnClick4: TNotifyEvent;
    FOnBtnClick5: TNotifyEvent;
    FOnBtnClick6: TNotifyEvent;
    FOnBtnClick7: TNotifyEvent;
    FOnThumbImagePaint: TNotifyEventPaint;
    FOnMarkFuulScreenWindow: TNotifyEvent;
    FOnTabActive: TNotifyEvent;
    FOnTabOrder: TNotifyEvent;
    FOnRegisterTab: TNotifyEvent;
    FOnUnregisterTab: TNotifyEvent;
    FOnActiveAlt: TNotifyEvent;
    FOnDeleteTab: TNotifyEvent;
    FOnAddTab: TNotifyEvent;
    function GetAppId: string;
    function GetTaskBarEntryHandle: THandle;
    function SetWindowAttribute(dwEnable: Boolean; dwAttr: DWORD): HRESULT;
    procedure MainWndProc(var Message: TMessage);
    procedure SetVisible(const Value: Boolean);
    procedure SetShowHint(const Value: Boolean);
    procedure SetUseClipRect(const Value: Boolean);
    procedure SetShowThumbImage(const Value: Boolean);
    procedure SetOverIconText(const Value: string);
    procedure SetToolTip(const Value: string);
    procedure SetAppId(const Value: string);
    procedure SetProgress(const Value: Integer);
    procedure SetProgressMax(const Value: Integer);
    procedure SetThumbImg(Value: TBitmap);
    procedure SetImageList(const Value: TImageList);
    procedure SetTHButton(const Value: THButton);
    procedure SetClipRect(const Value: TClipRect);
    procedure SetTaskState(const Value: TTaskState);
    procedure SetMessageFilter(msgName: Cardinal; msgParam: Word);
    procedure TimEvent(Sender: TObject);
    procedure NewFlags(i,j: Integer);
  protected
     procedure ThumbImagePaint(var ThumbImage: TBitmap); dynamic;

     procedure SetClick1; dynamic;
     procedure SetClick2; dynamic;
     procedure SetClick3; dynamic;
     procedure SetClick4; dynamic;
     procedure SetClick5; dynamic;
     procedure SetClick6; dynamic;
     procedure SetClick7; dynamic;

     procedure SetMarkFullClick; dynamic;
     procedure SetTabActiveClick; dynamic;
     procedure SetTabOrderClick; dynamic;
     procedure SetRegisterTabClick; dynamic;
     procedure SetUnregisterTabClick; dynamic;
     procedure SetActiveAltClick; dynamic;
     procedure SetDeleteTabClick; dynamic;
     procedure SetAddTabClick; dynamic;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure SetOverIcon(Index: Integer);
    procedure SetThumbImage;
    procedure Invalidate;
    procedure SetLightButton(ATime,ACount,AFlags: Cardinal);
    {Import function}
    function SetTabActive(hwndTab: HWND; hwndMDI: HWND; tbatFlags: integer): HRESULT;
    function MarkFullScreenWindow(wHandle: HWND; fFullscreen: LongBool): HRESULT;
    function SetTabOrder(hwndTab: HWND; hwndInsertBefore: HWND): HRESULT;
    function RegisterTab(hwndTab: HWND; hwndMDI: HWND): HRESULT;
    function UnregisterTab(hwndTab: HWND): HRESULT;
    function SetActiveAlt(hwndAlt: HWND): HRESULT;
    function DeleteTab(hwndTab: HWND): HRESULT;
    function AddTab(hwndTab: HWND): HRESULT;
  published
    property Visible : Boolean read FVisible write SetVisible default True;
    property ShowHint : Boolean read FShowHint write SetShowHint default False;
    property ShowThumbImage : Boolean read FShowThumbImage write SetShowThumbImage
      default False;
    property UseClipRect : Boolean read FUseClipRect write SetUseClipRect
      default False;
    property AppId : string read FAppId write SetAppId;
    property OverIconText : string read FOverIconText write SetOverIconText;
    property Hint : string read FToolTip write SetToolTip;
    property ProgressValue : integer read FProgress write SetProgress default 0;
    property ProgressValueMax : integer read FProgressMax write SetProgressMax
     default 100;
    property ProgressState : TTaskState read FTaskState write SetTaskState
     default Noprogres;
    property ThumbButton : THButton read FThumbButton write SetTHButton;
    property ClipRect : TClipRect read FClipRect write SetClipRect;
    property ImageList : TImageList read FImageList write SetImageList;
    property ThumbImage : TBitmap read FThumbImage write SetThumbImg;

    property onClick1: TNotifyEvent read FOnBtnClick1 write FOnBtnClick1;
    property onClick2: TNotifyEvent read FOnBtnClick2 write FOnBtnClick2;
    property onClick3: TNotifyEvent read FOnBtnClick3 write FOnBtnClick3;
    property onClick4: TNotifyEvent read FOnBtnClick4 write FOnBtnClick4;
    property onClick5: TNotifyEvent read FOnBtnClick5 write FOnBtnClick5;
    property onClick6: TNotifyEvent read FOnBtnClick6 write FOnBtnClick6;
    property onClick7: TNotifyEvent read FOnBtnClick7 write FOnBtnClick7;

    property onTabActive: TNotifyEvent read FOnTabActive write FOnTabActive;
    property onTabOrder: TNotifyEvent read FOnTabOrder write FOnTabOrder;
    property onRegisterTab: TNotifyEvent read FOnRegisterTab write FOnRegisterTab;
    property onUnregisterTab: TNotifyEvent read FOnUnregisterTab write FOnUnregisterTab;
    property onActiveAlt: TNotifyEvent read FOnActiveAlt write FOnActiveAlt;
    property onDeleteTab: TNotifyEvent read FOnDeleteTab write FOnDeleteTab;
    property onAddTab: TNotifyEvent read FOnAddTab write FOnAddTab;
    property onMarkFullScreenWindow: TNotifyEvent
        read FOnMarkFuulScreenWindow
        write FOnMarkFuulScreenWindow;

    property onThumbImagePaint: TNotifyEventPaint
       read FOnThumbImagePaint
       write FOnThumbImagePaint;

  end;


procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('Dukuy', [TWinSevenConnect]);
end;


{ ================================= TWinSevenConnect ========================= }

function TWinSevenConnect.GetTaskBarEntryHandle: THandle;
begin
 if Application.MainFormOnTaskBar then
  Result := TWinControl(Owner).Handle
 else
  Result := Application.Handle;
end;

procedure TWinSevenConnect.NewFlags(i,j: Integer);
begin
      Button[i].dwFlags := 0;
     if tfDisable in FThumbButton.FBtn[i].FFlags then
      Button[i].dwFlags := Button[i].dwFlags + THBF_DISABLED;
     if tfDismissionClick in FThumbButton.FBtn[i].FFlags then
      Button[i].dwFlags := Button[i].dwFlags + THBF_DISMISSONCLICK;
     if tfNoBackground in FThumbButton.FBtn[i].FFlags then
      Button[i].dwFlags := Button[i].dwFlags + THBF_NOBACKGROUND;
     if tfHidden in FThumbButton.FBtn[i].FFlags then
      Button[i].dwFlags := Button[i].dwFlags + THBF_HIDDEN;
     if tfNonInteractive in FThumbButton.FBtn[i].FFlags then
      Button[i].dwFlags := Button[i].dwFlags + THBF_NONINTERACTIVE;
end;

constructor TWinSevenConnect.Create;
var
 i: Byte;
begin
 inherited Create(AOwner);

 for i  := 0 to TWinControl(Owner).ComponentCount -1 do
  if TWinControl(Owner).Components[i] is TWinSevenConnect then
   if TWinControl(Owner).Components[i] <> Self
    then raise Exception.Create(MSG_Double);

 if not (Owner is TCustomForm) then
  raise  Exception.Create(MSG_Owner);

 FVisible := True;
 FUseClipRect := False;
 FShowHint := False;
 FShowThumbImage := False;
 FToolTip := '';
 FAppId := '';
 FProgress := 0;
 FProgressMax := 100;
 FOverIconIndex := -1;
 FClipRect := TClipRect.Create;
 FClipRect.FInteger[1] := 0;
 FClipRect.FInteger[2] := 0;
 FClipRect.FInteger[3] := TWinControl(Owner).ClientWidth;
 FClipRect.FInteger[4] := TWinControl(Owner).ClientHeight;
 FTaskState := Noprogres;

 FThumbImage := TBitmap.Create;

 Handle := GetTaskBarEntryHandle;

 FThumbButton := THButton.Create(Self);

 if not (csDesigning in ComponentState) then begin
  FMainWndProc  := Pointer(GetWindowLong(Handle, GWL_WNDPROC));
  SetWindowLong(Handle, GWL_WNDPROC, Integer(MakeObjectInstance(MainWndProc)));
 end else begin
  if FMainWndProc <> nil then begin
    SetWindowLong(Handle, GWL_WNDPROC, Integer(FMainWndProc));
    FMainWndProc := nil;
  end;
 end;

  // Set Windows Message Filter
  SetMessageFilter(WM_COMMAND,MSG_Add);
  SetMessageFilter(WM_DWMSENDICONICTHUMBNAIL,MSG_Add);

 if CheckWin32Version(6, 1) then
  begin
    CoInitialize(nil); // initializa ActiveX
    TaskBarList := CreateComObject(CLSID_TaskbarList) as ITaskBarList;
    TaskBarList.HrInit;
    Supports(TaskBarList, IID_ITaskbarList3, TaskbarList3);
    if not (csDesigning in ComponentState) then begin
     Timer := TTimer.Create(nil);
     Timer.Interval := 50;
     Timer.OnTimer := TimEvent;
    end;
  end
   else
     if csDesigning in ComponentState then
       MessageBox(Handle,Pchar(MSG_Sorry),'Information',MB_OK);
end;

destructor TWinSevenConnect.Destroy;
begin
 if Assigned(Timer) then begin
   Timer.Enabled := False;
   Timer.Free;
 end;
 // Free

  FThumbImage.Free;
  FClipRect.Free;
  FThumbButton.Free;
  inherited Destroy;
end;

procedure TWinSevenConnect.TimEvent(Sender: TObject);
var
 i,j: Byte;
begin
  SetProgress(FProgress);
  SetToolTip(FToolTip);
  SetTaskState(FTaskState);
  SetProgressMax(FProgressMax);
  SetClipRect(FClipRect);
  SetOverIconText(FOverIconText);
  SetOverIcon(FOverIconIndex);
  if Assigned(FImageList) then begin
   ThumbButton.THBtn1.SetVisible(ThumbButton.THBtn1.FVisible);
  j := 0;


  for I := 1 to THBNCount do
   if BtnMass[i] then begin
     inc(j);
     Button[j].iId := THBN + i;
     Button[j].iBitmap := FThumbButton.FBtn[i].ImageIndex;
     Button[j].dwMask := THB_FLAGS or THB_BITMAP or THB_TOOLTIP;
     NewFlags(i,j);
     if FThumbButton.FBtn[i].FSHowHint then
      StrCopy(Button[i].szTip, PChar(ThumbButton.FBtn[i].FHint))
     else
      StrCopy(Button[i].szTip, PChar(''));
     TaskBarList3.ThumbBarSetImageList(Handle, FImageList.Handle);
   end;
  TaskbarList3.ThumbBarAddButtons(Handle, j, @Button);
  if Assigned(Timer) then Timer.Enabled := false;
 end;
end;

procedure TWinSevenConnect.MainWndProc (var Message : TMessage);
begin
 case Message.Msg of
   WM_COMMAND:
    begin
     if HIWORD(Message.WParam) = THBN_CLICKED then
      case LOWORD(Message.WParam) of
       41: SetClick1;
       42: SetClick2;
       43: SetClick3;
       44: SetClick4;
       45: SetClick5;
       46: SetClick6;
       47: SetClick7;
     end;
    end;
   WM_DWMSENDICONICTHUMBNAIL:
    begin
     if (FThumbImage.Handle <> 0) and FShowThumbImage then
      begin
        ThumbImagePaint(FThumbImage);
        SetThumbImage;
      end;
    end;
   WM_DWMSENDICONICLIVEPREVIEWBITMAP:
    begin

    end;
 end;

 Message.Result := CallWindowProc(FMainWndProc, Handle,
           Message.Msg, Message.wParam, Message.lParam);
end;

procedure TWinSevenConnect.SetAppId(const Value: string);
var
  AppId : function (AppID : PWChar) : HResult; stdcall;
  hndDLLHandle: THandle;
begin
  try
    hndDLLHandle := loadLibrary('Shell32.dll');
    if hndDLLHandle <> 0 then
    begin
      AppId := getProcAddress(hndDLLHandle,'SetCurrentProcessExplicitAppUserModelID');
      if addr(appId) <> nil then
        begin
         AppId(PWChar(WideString(Value)));
         FAppId := Value;
        end;
    end;
  finally
    freeLibrary(hndDLLHandle);
  end;
end;

function TWinSevenConnect.GetAppId: string;
var
  AppId : function (var AppID : PWChar) : HResult; stdcall;
  hndDLLHandle: THandle;
  getid: PWChar;
begin
  try
    Result := '';
    hndDLLHandle := loadLibrary('Shell32.dll');
    if hndDLLHandle <> 0 then
    begin
      AppId := getProcAddress(hndDLLHandle,'GetCurrentProcessExplicitAppUserModelID');
      if addr(appId) <> nil then
        begin
          AppId(getid);
          Result := WideCharToString(getid);
          CoTaskMemFree(getid);
        end;
    end;
  finally
    freeLibrary(hndDLLHandle);
  end;
end;


procedure TWinSevenConnect.SetMessageFilter(msgName: Cardinal; msgParam: Word);
var
  hndDLLHandle: THandle;
  msg: function(msg: Cardinal; action: Word): BOOL; stdcall;
begin
  try
    hndDLLHandle := loadLibrary('user32.dll');
    if hndDLLHandle <> 0 then
    begin
      @msg := getProcAddress(hndDLLHandle, 'ChangeWindowMessageFilter');
      if addr(msg) <> nil then
       msg(msgName, msgParam);
    end;
  finally
    freeLibrary(hndDLLHandle);
  end;
end;

function TWinSevenConnect.SetWindowAttribute(dwEnable: Boolean; dwAttr: DWORD): HRESULT;
var
  Attr: function(hwnd: hwnd; dwAttribute: DWORD; pvAttribute: Pointer;
   cbAttribute: DWORD): HRESULT; stdcall;
 hndDLLHandle : THandle;
 EnableAttribute : DWORD;
begin
  EnableAttribute := DWORD(dwEnable);
  try
    hndDLLHandle := loadLibrary('dwmapi.dll');
    if hndDLLHandle <> 0 then
    begin
      Attr := getProcAddress(hndDLLHandle, 'DwmSetWindowAttribute');
      if addr(Attr) <> nil then
       Result := Attr(Handle, dwAttr, @EnableAttribute, SizeOf(EnableAttribute));
    end;
  finally
    freeLibrary(hndDLLHandle);
  end;
end;

procedure TWinSevenConnect.SetLightButton(ATime,ACount,AFlags: Cardinal);
var
 FWInfo: TFlashWInfo;
begin
  with FWInfo do
    begin
      cbSize := SizeOf(FWInfo);
      hWnd := Handle;
      dwFlags := AFlags;
      uCount := ACount;
      dwTimeOut := ATime;
    end;
  FlashWindowEx(FWInfo);
end;

procedure TWinSevenConnect.SetThumbImage;

procedure p24top32(bmp: TBitmap);
type
  pRGB = ^TRGBQuad;
var
  x, y: Integer;
  Dest: pRGB;
begin
  bmp.PixelFormat := pf32bit;
  for y := 0 to Bmp.Height - 1 do
  begin
    Dest := Bmp.ScanLine[y];
    for x := 0 to Bmp.Width - 1 do
    begin
      with Dest^ do begin
       rgbReserved := 255;
       Inc(Dest);
      end;
    end;
  end;
end;

var
  hndDLLHandle: THandle;
  DrawImage: function (hwnd: HWND; hbmp: HBITMAP; dwSITFlags: DWORD): HResult; stdcall;
begin
  p24top32(FThumbImage);
  try
    hndDLLHandle := loadLibrary('dwmapi.dll');
    if hndDLLHandle <> 0 then
    begin
      DrawImage := getProcAddress(hndDLLHandle, 'DwmSetIconicThumbnail');
      if addr(DrawImage) <> nil then
        DrawImage(Handle,FThumbImage.Handle,0);
    end;
  finally
    freeLibrary(hndDLLHandle);
  end;
end;

procedure TWinSevenConnect.Invalidate;
var
 Refresh:  function (hwnd: HWND): HResult; stdcall;
 hndDLLHandle : THandle;
begin
  try
    hndDLLHandle := loadLibrary('dwmapi.dll');
    if hndDLLHandle <> 0 then
    begin
      Refresh := getProcAddress(hndDLLHandle, 'DwmInvalidateIconicBitmaps');
      if addr(Refresh) <> nil then
        Refresh(Handle);
    end;
  finally
    freeLibrary(hndDLLHandle);
  end;
end;

procedure TWinSevenConnect.SetTaskState(const Value: TTaskState);
begin
     FTaskState := Value;
     if Assigned(TaskBarList3) then begin
      case FTaskState of
       Noprogres, Identerminate, Normal:
         TaskbarList3.SetProgressState(Handle,integer(Value));
       Error:
         TaskbarList3.SetProgressState(Handle,4);
       Pause:
         TaskbarList3.SetProgressState(Handle,8);
       end;
      SetProgress(FProgress);
      end;
end;

procedure TWinSevenConnect.SetShowHint(const Value: Boolean);
begin
 FShowHint := Value;
 if Assigned(TaskbarList3) then
  if FShowHint then
   TaskBarList3.SetThumbnailTooltip(Handle,PWideChar(FToolTip))
  else
   TaskBarList3.SetThumbnailTooltip(Handle,PWideChar(''));
end;

procedure TWinSevenConnect.SetShowThumbImage(const Value: Boolean);
begin
 if Value <> FShowThumbImage then
  begin
   FShowThumbImage := Value;
   SetWindowAttribute(Value,DWMWA_FORCE_ICONIC_REPRESENTATION);
   SetWindowAttribute(Value,DWMWA_HAS_ICONIC_BITMAP);
   SetWindowAttribute(Value,DWMWA_DISALLOW_PEEK);
  end;
end;

procedure TWinSevenConnect.SetUseClipRect(const Value: Boolean);
begin
  if Value <> FUseClipRect then
   begin
    FUseClipRect := Value;
    SetClipRect(FClipRect);
   end;
end;

procedure TWinSevenConnect.SetVisible(const Value: Boolean);
begin
  if Value <> FVisible then
   FVisible := Value;

  if Assigned(TaskBarList) then
   if Value then begin
     if not (csDesigning in ComponentState) then begin
      TaskBarList.AddTab(Handle);
      if Assigned(Timer) then Timer.Enabled := True;
     end;
   end
     else
      TaskBarList.DeleteTab(Handle)
end;

procedure TWinSevenConnect.SetOverIconText(const Value: string);
begin
   FOverIconText := Value;
end;

procedure TWinSevenConnect.SetToolTip(const Value: string);
begin
   FToolTip := Value;
   SetShowHint(FShowHint);
end;

procedure TWinSevenConnect.SetProgressMax(const Value: Integer);
begin
 if Value <> FProgressMax then
  begin
   FProgressMax := Value;
   SetProgress(FProgress);
  end;
end;

procedure TWinSevenConnect.SetProgress(const Value: Integer);
begin
  FProgress := Value;
  if FTaskState = Identerminate then Exit;// - use this line from not lose State
  if Assigned(TaskBarList3) and (FTaskState <> Noprogres) then
   TaskbarList3.SetProgressValue(Handle, FProgress, FProgressMax);
end;

procedure TWinSevenConnect.SetClipRect(const Value: TClipRect);
var
  PRect : ^TRect;
begin
   FClipRect := Value;
   New(PRect);
   if FUseClipRect then begin
    PRect^.Left := FClipRect.FInteger[1];
    PRect^.Top := FClipRect.FInteger[2];
    PRect^.Right := FClipRect.FInteger[3];
    PRect^.Bottom := FClipRect.FInteger[4];
   end else PRect := nil;
   if Assigned(TaskbarList3) then
    TaskbarList3.SetThumbnailClip(Handle,Prect^);
   Dispose(PRect);
end;

procedure TWinSevenConnect.SetOverIcon(Index: Integer);
var
  MyIcon: TIcon;
begin
  if FImageList = nil then Exit;

  if Index <= -1 then
   begin
     if Assigned(TaskBarList3) then
      TaskbarList3.SetOverlayIcon(Handle,0,PChar(FOverIconText));
     Exit;
   end;

   MyIcon := TIcon.Create;
   try
     FImageList.GetIcon(Index, MyIcon);
     if Assigned(TaskBarList3) then
      TaskbarList3.SetOverlayIcon(Handle,MyIcon.Handle, PChar(FOverIconText));
     FOverIconIndex := Index;
   finally
     MyIcon.Free;
   end;
end;

procedure TWinSevenConnect.SetTHButton(const Value: THButton);
begin
 FThumbButton := Value;
end;

procedure TWinSevenConnect.SetThumbImg(Value: TBitmap);
begin
  FThumbImage.Assign(Value);
end;

procedure TWinSevenConnect.SetImageList(const Value: TImageList);
begin
  FImageList := Value;
end;

procedure TWinSevenConnect.ThumbImagePaint(var ThumbImage: TBitmap);
begin
 if Assigned(FOnThumbImagePaint) then FOnThumbImagePaint(Self,FThumbImage);
end;

procedure TWinSevenConnect.SetClick1;
begin
 if Assigned(FOnBtnClick1) then FOnBtnClick1(Self);
end;

procedure TWinSevenConnect.SetClick2;
begin
 if Assigned(FOnBtnClick2) then FOnBtnClick2(Self);
end;

procedure TWinSevenConnect.SetClick3;
begin
 if Assigned(FOnBtnClick3) then FOnBtnClick3(Self);
end;

procedure TWinSevenConnect.SetClick4;
begin
 if Assigned(FOnBtnClick4) then FOnBtnClick4(Self);
end;

procedure TWinSevenConnect.SetClick5;
begin
 if Assigned(FOnBtnClick5) then FOnBtnClick5(Self);
end;

procedure TWinSevenConnect.SetClick6;
begin
 if Assigned(FOnBtnClick6) then FOnBtnClick6(Self);
end;

procedure TWinSevenConnect.SetClick7;
begin
 if Assigned(FOnBtnClick7) then FOnBtnClick7(Self);
end;

procedure TWinSevenConnect.SetMarkFullClick;
begin
 if Assigned(FOnMarkFuulScreenWindow) then FOnMarkFuulScreenWindow(Self);
end;

procedure TWinSevenConnect.SetTabActiveClick;
begin
 if Assigned(FOnTabActive) then FOnTabActive(Self);
end;

procedure TWinSevenConnect.SetTabOrderClick;
begin
 if Assigned(FOnTabOrder) then FOnTabOrder(Self);
end;

procedure TWinSevenConnect.SetRegisterTabClick;
begin
 if Assigned(FOnRegisterTab) then FOnRegisterTab(Self);
end;

procedure TWinSevenConnect.SetUnregisterTabClick;
begin
 if Assigned(FOnUnregisterTab) then FOnUnregisterTab(Self);
end;

procedure TWinSevenConnect.SetActiveAltClick;
begin
 if Assigned(FOnActiveAlt) then FOnActiveAlt(Self);
end;

procedure TWinSevenConnect.SetDeleteTabClick;
begin
 if Assigned(FOnDeleteTab) then FOnDeleteTab(Self);
end;

procedure TWinSevenConnect.SetAddTabClick;
begin
 if Assigned(FOnAddTab) then FOnAddTab(Self);
end;

{ ================================ Import function =========================== }

function TWinSevenConnect.MarkFullscreenWindow(wHandle: HWND; fFullscreen: LongBool): HRESULT;
begin
  if Assigned(TaskbarList3) then
   Result := TaskbarList3.MarkFullscreenWindow(wHandle,fFullscreen);
  SetMarkFullClick;
end;

function TWinSevenConnect.RegisterTab(hwndTab: HWND; hwndMDI: HWND): HRESULT;
begin
  if Assigned(TaskbarList3) then
   Result := TaskbarList3.RegisterTab(hwndTab,hwndMDI);
  SetRegisterTabClick;
end;

function TWinSevenConnect.UnregisterTab(hwndTab: HWND): HRESULT;
begin
  if Assigned(TaskbarList3) then
   Result := TaskbarList3.UnregisterTab(hwndTab);
  SetUnregisterTabClick;
end;

function TWinSevenConnect.AddTab(hwndTab: HWND): HRESULT;
begin
  if Assigned(TaskbarList3) then
   Result := TaskbarList3.AddTab(hwndTab);
  SetAddTabClick;
end;

function TWinSevenConnect.DeleteTab(hwndTab: HWND): HRESULT;
begin
  if Assigned(TaskbarList3) then
   Result := TaskbarList3.DeleteTab(hwndTab);
  SetDeleteTabClick;
end;

function TWinSevenConnect.SetActiveAlt(hwndAlt: HWND): HRESULT;
begin
  if Assigned(TaskbarList3) then
   Result := TaskbarList3.SetActiveAlt(hwndAlt);
  SetActiveAltClick;
end;

function TWinSevenConnect.SetTabOrder(hwndTab: HWND; hwndInsertBefore: HWND): HRESULT;
begin
  if Assigned(TaskbarList3) then
   Result := TaskbarList3.SetTabOrder(hwndTab,hwndInsertBefore);
  SetTabOrderClick;
end;

function TWinSevenConnect.SetTabActive(hwndTab: HWND; hwndMDI: HWND; tbatFlags: integer): HRESULT;
begin
  if Assigned(TaskbarList3) then
   Result := TaskbarList3.SetTabActive(hwndTab,hwndMDI,tbatFlags);
  SetTabActiveClick;
end;

{ =============================== TClipRect ================================== }

function TClipRect.GetInt(Index: Integer): Integer;
begin
 Result := FInteger[ Index ];
end;

procedure TClipRect.SetInt(Index: Integer; Value: Integer);
begin
 if Value <> FInteger[ Index ] then
   FInteger[ Index ] := Value;
end;

{ ============================== TThumbButton ================================ }

constructor TThumbBtn.Create(AClass: TComponent);
begin
 BaseClass := AClass;
 FVisible := False;
 FSHowHint := False;
 FHint := '';
 FImageIndex := -1;
 FFlags := [];
end;

procedure TThumbBtn.SetVisible(const Value: Boolean);
begin
  if Value <> FVisible then
   FVisible := Value;

  if Assigned(FOnCreate) then FOnCreate(Self);
end;


procedure TThumbBtn.SetShowHint(const Value: Boolean);
begin
  if Value <> FSHowHint then
   FSHowHint := Value;

  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TThumbBtn.SetImageIndex(const Value: Integer);
begin
  if Value <> FImageIndex then
   FImageIndex := Value;

  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TThumbBtn.SetHint(const Value: string);
begin
  if Value <> FHint then
   FHint := Value;

 if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TThumbBtn.SetFlags(const Value: TFlags);
begin
  if Value <> FFlags then
   FFlags := Value;

  if Assigned(FOnChange) then FOnChange(Self);
end;

{ ================================= THbtn ==================================== }

constructor THButton.Create(AClass: TComponent);
var
 i: Byte;
begin
 BaseClass := AClass;
 for I := 1 to THBNCount do begin
   FBtn[I] := TThumbBtn.Create(AClass);
   FBtn[I].OnCreate := ACreate;
   FBtn[I].OnChange := AChanged;
 end;
end;

function THButton.GetBtn(Index: Integer): TThumbBtn;
begin
 Result := FBtn[ Index ];
end;

procedure THButton.SetInt(Index: Integer; Value: TThumbBtn);
begin
  if Value <> FBtn[ Index ] then
   FBtn[ Index ] := Value;
end;

procedure THButton.ACreate(Sender: TObject);
var
 i: Byte;
begin
 if not( csDesigning in TWinSevenConnect(BaseClass).ComponentState ) then
  for I := 1 to THBNCount do
   if (Sender = FBtn[i]) and FBtn[i].FVisible then
    TWinSevenConnect(BaseClass).BtnMass[i] := True;
end;

procedure THButton.AChanged(Sender: TObject);
var
 i: Byte;
begin
 if not( csDesigning in TWinSevenConnect(BaseClass).ComponentState ) then
  for I := 1 to THBNCount do
   if (Sender = FBtn[i]) then
    begin
     with TWinSevenConnect(BaseClass) do begin
     if (Button[i].iId <= 47) and (Button[i].iId > 40) then begin
     Button[i].iBitmap := FThumbButton.FBtn[i].FImageIndex;
     Button[i].dwMask := THB_FLAGS or THB_BITMAP or THB_TOOLTIP;
     NewFlags(i,i);
     if FBtn[i].FSHowHint then
      StrCopy(Button[i].szTip, PChar(FBtn[i].FHint))
     else
      StrCopy(Button[i].szTip, PChar(''));
     TaskBarList3.ThumbBarSetImageList(Handle, FImageList.Handle);
     TaskbarList3.ThumbBarUpdateButtons(Handle, 1, @Button[i]);
     end;
    end;
   end;
end;

end.
