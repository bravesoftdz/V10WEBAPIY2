// ****************************************************************************
// * Ajout 01-05-2007 : Sélection a la semaine via option csWeekSelectionOnly
// ****************************************************************************
// * Calendar component for Delphi.
// ****************************************************************************
// * Copyright 2001-2002, Bitvad sz Kft. All Rights Reserved.
// ****************************************************************************
// * This component can be freely used and distributed in commercial and
// * private environments, provied this notice is not modified in any way.
// ****************************************************************************
// * Feel free to contact me if you have any questions, comments or suggestions
// * at support@maxcomponents.net
// ****************************************************************************
// * Web page: www.maxcomponents.net
// ****************************************************************************
// * Date last modified: 18.10.2002
// ****************************************************************************
// * TmxCalendar v1.2
// ****************************************************************************
// * Description:
// *
// * The TmCalendar is a VCL component to create internet link on your forms.
// ****************************************************************************

Unit NewCalendar;

Interface

// *************************************************************************************
// ** List of used units
// *************************************************************************************

Uses
     Windows,
     Messages,
     SysUtils,
     StdCtrls,
     Classes,
     Graphics,
     Controls,
     Extctrls,
     Forms,
     Hent1,
     Buttons;

{$I MAX.INC}
{$DEFINE FRENCH}
Const
  mxCalendarVersion   = $020B;
     NavigateButtonWidth = 23;
     MonthButtonWidth = 100;

Type

  {CustomPanel to resolve the XP Theme problem}
  TmxPanel = class(TCustomControl)
  Private
    FFlat: Boolean;
    FcoordX : integer;
    fCoordY : integer;
    FDate : TDateTime;
    FWeek : Integer;
    Fmonth : integer;
    procedure SetFlat( AValue: boolean );
    procedure WMSize(var Message: TWMSize); message WM_SIZE;
  Protected
    procedure Paint; override;
    procedure CreateParams(var Params: TCreateParams); override;
  Public
    constructor Create(AOwner: TComponent); override;
    Property Flat: Boolean read FFlat Write SetFlat Default False;
    property CoordX : integer read FcoordX write fCoordX;
    property CoordY : integer read FcoordY write fCoordY;
  end;

     TmxCalendarOption = (
          coClearButtonVisible,
          coFlatButtons,
          csFlatDays,
          coFlatHeaders,
          coMonthButtonVisible,
          csWeekSelectionOnly,
          csSelectionEnabled,
          csSetTodayOnStartup,
          coShowDateLabel,
          coShowFooter,
          coShowNextMonth,
          coShowPreviousMonth,
          coShowHeader,
          coShowWeekDays,
          coShowWeeks,
          coTransparentButtons,
          coTodayButtonVisible,
          csUseWeekEndColor,
          csUseWeekEndFont,
          coYearButtonVisible );

     TmxCalendarOptions = Set Of TmxCalendarOption;
     TmxFirstDayOfWeek = ( fdSunday, fdMonday );
     TmxHeaderFormat = ( hfMMMMYYYY, hfYYYYMMMM, hfMMYYYY, hfYYYYMM );

     // ************************************************************************
     // ************************************************************************

     TmxHints = Class( TPersistent )
     Private

          FHints: Array[ 0..6 ] Of ShortString;
          FOnChange: TNotifyEvent;

          Procedure SetHint( Index: Integer; AValue: ShortString );
          Function GetHint( Index: Integer ): ShortString;

     Protected

          Procedure Change; Virtual;

     Public

          Constructor Create;
          Procedure Assign( Source: TPersistent ); Override;

          Property Hints[ Index: Integer ]: ShortString Read GetHint; Default;

     Published

          Property OnChange: TNotifyEvent Read FOnChange Write FOnChange;

          Property NextMonth: ShortString Index 0 Read GetHint Write SetHint;
          Property PreviousMonth: ShortString Index 1 Read GetHint Write SetHint;
          Property NextYear: ShortString Index 2 Read GetHint Write SetHint;
          Property PreviousYear: ShortString Index 3 Read GetHint Write SetHint;
          Property TodayButton: ShortString Index 4 Read GetHint Write SetHint;
          Property Today: ShortString Index 5 Read GetHint Write SetHint;
          Property ClearButton: ShortString Index 6 Read GetHint Write SetHint;

     End;

     // ************************************************************************
     // ************************************************************************

     TmxMessages = Class( TPersistent )
     Private

          FMonthNames: Array[ 1..12 ] Of ShortString;
          FMessages: Array[ 0..7 ] Of ShortString;
          FHints: TmxHints;
          FOnChange: TNotifyEvent;

          Procedure SetMessage( Index: Integer; AValue: ShortString );
          Function GetMessage( Index: Integer ): ShortString;

          Procedure SetMonthName( Index: Integer; AValue: ShortString );
          Function GetMonthName( Index: Integer ): ShortString;

          Procedure SetOnChange( AOnChange: TNotifyEvent );

     Protected

          Procedure Change; Virtual;

     Public

          Constructor Create;
          Destructor Destroy; Override;
          Procedure Assign( Source: TPersistent ); Override;

          Property Messages[ Index: Integer ]: ShortString Read GetMessage; Default;
          Property MonthNames[ Index: Integer ]: ShortString Read GetMonthName;

     Published

          Property OnChange: TNotifyEvent Read FOnChange Write SetOnChange;

          Property Hints: TmxHints Read FHints Write FHints;

          Property Week: ShortString Index 0 Read GetMessage Write SetMessage;
          Property Sunday: ShortString Index 1 Read GetMessage Write SetMessage;
          Property Monday: ShortString Index 2 Read GetMessage Write SetMessage;
          Property Tuesday: ShortString Index 3 Read GetMessage Write SetMessage;
          Property Wednesday: ShortString Index 4 Read GetMessage Write SetMessage;
          Property Thursday: ShortString Index 5 Read GetMessage Write SetMessage;
          Property Friday: ShortString Index 6 Read GetMessage Write SetMessage;
          Property Saturday: ShortString Index 7 Read GetMessage Write SetMessage;

          Property January: ShortString Index 1 Read GetMonthName Write SetMonthName;
          Property February: ShortString Index 2 Read GetMonthName Write SetMonthName;
          Property March: ShortString Index 3 Read GetMonthName Write SetMonthName;
          Property April: ShortString Index 4 Read GetMonthName Write SetMonthName;
          Property May: ShortString Index 5 Read GetMonthName Write SetMonthName;
          Property June: ShortString Index 6 Read GetMonthName Write SetMonthName;
          Property July: ShortString Index 7 Read GetMonthName Write SetMonthName;
          Property August: ShortString Index 8 Read GetMonthName Write SetMonthName;
          Property September: ShortString Index 9 Read GetMonthName Write SetMonthName;
          Property October: ShortString Index 10 Read GetMonthName Write SetMonthName;
          Property November: ShortString Index 11 Read GetMonthName Write SetMonthName;
          Property December: ShortString Index 12 Read GetMonthName Write SetMonthName;
     End;

     // ************************************************************************
     // ************************************************************************

     TmxCalendarButton = Class( TSpeedButton )
     Private
     		Procedure CMDesignHitTest( Var Msg: TCMDesignHitTest ); Message CM_DESIGNHITTEST;
     End;

     // ************************************************************************
     // ************************************************************************

     TmxCustomCalendar = Class( TCustomPanel )
     Private
          FYear: Integer;
          FMonth: Integer;
          FDay: Integer;
          FDateLabel: TLabel;
          FPanel_Header: TmxPanel;
          FPanel_Footer: TmxPanel;
          FPanel_Calendar: TmxPanel;
          FPanel_Days: array[0..6, 0..7] of TmxPanel;
          FYear_Plus: TmxCalendarButton;
          FYear_Minus: TmxCalendarButton;
          FMonth_Plus: TmxCalendarButton;
          FMonth_Minus: TmxCalendarButton;
          FTodayButton: TmxCalendarButton;
          FClearButton: TmxCalendarButton;
          FSelectCurrMonth : TmxCalendarButton;
          FDayColor: TColor;
          FSelectedColor: TColor;
          FDayNameColor: TColor;
          FWeekColor: TColor;
          FTodayColor: TColor;
          FIndirectColor: TColor;
          FWeekEndColor: TColor;
          FDateFormat: String;
          FOptions: TmxCalendarOptions;
          FMessages: TmxMessages;
          FFirstDayOfWeek: TmxFirstDayOfWeek;
          FHeaderFormat: TmxHeaderFormat;
          FVersion: Integer;
          FSelectedFont: TFont;
          FTodayFont: TFont;
          FDayFont: TFont;
          FIndirectFont: TFont;
          FWeekDaysFont: TFont;
          FWeeksFont: TFont;
          FWeekEndFont: TFont;
          FSelectionStart: TDateTime;
          FSelectionEnd: TDateTime;
          FMinDate : TdateTime;
          FMaxDate : TdateTime;
          fWeekSelect : integer;
          Procedure SetVersion( AValue: String );
          Function GetVersion: String;
          Procedure SetSelectedColor( AValue: TColor );
          Procedure SetWeekEndColor( AValue: TColor );
          Procedure SetDayColor( AValue: TColor );
          Procedure SetWeekColor( AValue: TColor );
          Procedure SetTodayColor( AValue: TColor );
          Procedure SetDayNameColor( AValue: TColor );
          Procedure SetHeaderColor( AValue: TColor );
          Procedure SetFooterColor( AValue: TColor );
          Procedure SetIndirectColor( AValue: TColor );
          Function GetHeaderColor: TColor;
          Function GetFooterColor: TColor;
          Procedure SetOptions( AValue: TmxCalendarOptions );
          Procedure SetWeekDaysFont( AValue: TFont );
          Procedure SetIndirectFont( AValue: TFont );
          Procedure SetWeeksFont( AValue: TFont );
          Procedure SetWeekEndFont( AValue: TFont );
          Function GetDateLabelFont: TFont;
          Procedure SetDateLabelFont( AValue: TFont );
          Procedure SetSelectedFont( AValue: TFont );
          Procedure SetTodayFont( AValue: TFont );
          Procedure SetDayFont( AValue: TFont );
          Procedure SetDateFormat( AValue: String );
          Function GetClearButtonGlyph: TBitmap;
          Procedure SetClearButtonGlyph( AValue: TBitmap );
          Function GetTodayButtonGlyph: TBitmap;
          Procedure SetTodayButtonGlyph( AValue: TBitmap );
          Function GetYearMinusGlyph: TBitmap;
          Procedure SetYearMinusGlyph( AValue: TBitmap );
          Function GetYearPlusGlyph: TBitmap;
          Procedure SetYearPlusGlyph( AValue: TBitmap );
          Function GetMonthMinusGlyph: TBitmap;
          Procedure SetMonthMinusGlyph( AValue: TBitmap );
          Function GetMonthPlusGlyph: TBitmap;
          Procedure SetMonthPlusGlyph( AValue: TBitmap );
          Procedure SetFirstDayOfWeek( AValue: TmxFirstDayOfWeek );
          Procedure SetYear( AValue: Integer );
          Procedure SetMonth( AValue: Integer );
          Procedure SetDay( AValue: Integer );
          Procedure _SetDate( AValue: TDateTime );
          Procedure SetDate( AValue: TDateTime );
          Function GetDate: TDateTime;
          Procedure SetHeaderFormat( AValue: TmxHeaderFormat );
{$IFDEF DELPHI4_UP}
          Procedure CMBorderChanged( Var Message: TMessage ); Message CM_BORDERCHANGED;
{$ENDIF}
          Procedure CMCtl3DChanged( Var Message: TMessage ); Message CM_CTL3DCHANGED;
          Procedure CMSysColorChange( Var Message: TMessage ); Message CM_SYSCOLORCHANGE;
          Procedure CMColorChanged( Var Message: TMessage ); Message CM_COLORCHANGED;
          Procedure CMEnabledChanged( Var Message: TMessage ); Message CM_ENABLEDCHANGED;
    			function TheDateInMonth(Decalage: integer): TdateTime;
          function FindWeek (WeekNumber : integer) : TmxPanel;
    			function FindDay(UneDate: TDateTime): TmxPanel;
     Protected
          Procedure SetPanelColor( APanel: TmxPanel; AYear, AMonth: Integer ); Virtual;
          Procedure CMFontChanged( Var Msg: TMessage ); Message CM_FONTCHANGED;
          Procedure CreateParams( Var Params: TCreateParams ); Override;
          Procedure OnChangeMessages( Sender: TObject );
          Procedure OnResizePanels( Sender: TObject );
          Procedure OnCanResizePanels( Sender: TObject; Var NewWidth, NewHeight: Integer; Var Resize: Boolean );
          Procedure SetButtonPositions;
          Procedure SetButtonCaptions;
          Procedure Loaded; Override;
          Procedure RepaintCalendar; Virtual;
          Function SetWeekStart( ADayIndex: Integer ): Integer;
          Procedure YearButtonClick( Sender: TObject );
          Procedure MonthButtonClick( Sender: TObject );
          Procedure CurrMonthButtonClick( Sender: TObject );
          Procedure TodayButtonClick( Sender: TObject );
          Procedure ClearButtonClick( Sender: TObject );
          Procedure DoMouseDown( Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer );
          procedure doWeekSel(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
          Procedure SetSelectionStart( AValue: TDateTime );
          Procedure SetSelectionEnd( AValue: TDateTime );
          Procedure Paint; Override;
          // *** Published ***
          Property Year: Integer Read FYear Write SetYear;
          Property Month: Integer Read FMonth Write SetMonth;
          Property Day: Integer Read FDay Write SetDay;
          Property DateFormat: String Read FDateFormat Write SetDateFormat;
          Property DayColor: TColor Read FDayColor Write SetDayColor;
          Property WeekEndColor: TColor Read FWeekEndColor Write SetWeekEndColor;
          Property SelectedColor: TColor Read FSelectedColor Write SetSelectedColor;
          Property DayNameColor: TColor Read FDayNameColor Write SetDayNameColor;
          Property WeekColor: TColor Read FWeekColor Write SetWeekColor;
          Property TodayColor: TColor Read FTodayColor Write SetTodayColor;
          Property IndirectColor: TColor Read FIndirectColor Write SetIndirectColor;
          Property HeaderColor: TColor Read GetHeaderColor Write SetHeaderColor;
          Property FooterColor: TColor Read GetFooterColor Write SetFooterColor;
          Property Options: TmxCalendarOptions Read FOptions Write SetOptions;
          Property Messages: TmxMessages Read FMessages Write FMessages;
          Property SelectedFont: TFont Read FSelectedFont Write SetSelectedFont;
          Property DayFont: TFont Read FDayFont Write SetDayFont;
          Property TodayFont: TFont Read FTodayFont Write SetTodayFont;
          Property WeeksFont: TFont Read FWeeksFont Write SetWeeksFont;
          Property WeekEndFont: TFont Read FWeekEndFont Write SetWeekEndFont;
          Property DateLabelFont: TFont Read GetDateLabelFont Write SetDateLabelFont;
          Property WeekDaysFont: TFont Read FWeekDaysFont Write SetWeekDaysFont;
          Property IndirectFont: TFont Read FIndirectFont Write SetIndirectFont;
          Property ClearButtonGlyph: TBitmap Read GetClearButtonGlyph Write SetClearButtonGlyph;
          Property TodayButtonGlyph: TBitmap Read GetTodayButtonGlyph Write SetTodayButtonGlyph;
          Property YearMinusGlyph: TBitmap Read GetYearMinusGlyph Write SetYearMinusGlyph;
          Property YearPlusGlyph: TBitmap Read GetYearPlusGlyph Write SetYearPlusGlyph;
          Property MonthMinusGlyph: TBitmap Read GetMonthMinusGlyph Write SetMonthMinusGlyph;
          Property MonthPlusGlyph: TBitmap Read GetMonthPlusGlyph Write SetMonthPlusGlyph;
          Property HeaderFormat: TmxHeaderFormat Read FHeaderFormat Write SetHeaderFormat Default hfMMMMYYYY;
          Property FirstDayOfWeek: TmxFirstDayOfWeek Read FFirstDayOfWeek Write SetFirstDayOfWeek Default fdMonday;
     Public
          Property Date             : TDateTime Read GetDate Write SetDate;
          Property SelectionStart   : TDateTime Read FSelectionStart;
          Property SelectionEnd     : TDateTime Read FSelectionEnd;
          property WeekSelect       : integer read fWeekSelect;
          Constructor Create( AOwner: TComponent ); Override;
          Destructor Destroy; Override;
          Procedure ClearSelection;
          Function DaysInMonth( AYear, AMonth: Integer ): Integer;
          Function WeeksInYear( AYear: Integer ): Integer;
          Function MonthToWeek( AMonth: Integer ): Integer;
          Property MinDate          : TdateTime Read fMinDate write fMinDate;
          Property MaxDate          : TdateTime Read fMaxDate write fMaxDate;
     Published
          Property Version: String Read GetVersion Write SetVersion;
     End;

     TmxCalendar = Class( TmxCustomCalendar )
     Public
          Property DockManager;
     Published
          Property Align;
          Property Anchors;
          Property BevelInner;
          Property BevelOuter;
          Property BevelWidth;
          Property BorderWidth;
          Property BorderStyle;
          Property Color;
          Property Constraints;
          Property Ctl3D;
          Property UseDockManager Default True;
          Property DockSite;
          Property DragCursor;
          Property DragKind;
          Property DragMode;
          Property Enabled;
          Property ParentColor;
          Property ParentCtl3D;
          Property ParentShowHint;
          Property PopupMenu;
          Property ShowHint;
          Property TabOrder;
          Property TabStop;
          Property Visible;
          Property OnCanResize;
          Property OnClick;
          Property OnConstrainedResize;
          Property OnContextPopup;
          Property OnDockDrop;
          Property OnDockOver;
          Property OnDblClick;
          Property OnDragDrop;
          Property OnDragOver;
          Property OnEndDock;
          Property OnEndDrag;
          Property OnEnter;
          Property OnExit;
          Property OnGetSiteInfo;
          Property OnMouseDown;
          Property OnMouseMove;
          Property OnMouseUp;
          Property OnResize;
          Property OnStartDock;
          Property OnStartDrag;
          Property OnUnDock;

          Property Year;
          Property Month;
          Property Day;
          Property DateFormat;
          Property DayColor;
          Property WeekEndColor;
          Property SelectedColor;
          Property DayNameColor;
          Property WeekColor;
          Property TodayColor;
          Property IndirectColor;
          Property HeaderColor;
          Property FooterColor;
          Property Options;
          Property Messages;
          Property SelectedFont;
          Property DayFont;
          Property TodayFont;
          Property WeeksFont;
          Property WeekEndFont;
          Property DateLabelFont;
          Property WeekDaysFont;
          Property IndirectFont;
          Property ClearButtonGlyph;
          Property TodayButtonGlyph;
          Property YearMinusGlyph;
          Property YearPlusGlyph;
          Property MonthMinusGlyph;
          Property MonthPlusGlyph;
          Property HeaderFormat;
          Property FirstDayOfWeek;
     End;

Implementation


{$R *.RES} // a ne surtout pas enlever ni perdre....
constructor TmxPanel.Create(AOwner: TComponent);
begin
  inherited;
  ControlStyle := ControlStyle + [csAcceptsControls];
  FFlat:=False;
end;

procedure TmxPanel.WMSize(var Message: TWMSize);
begin
  inherited;
  Invalidate;
end;

procedure TmxPanel.SetFlat( AValue: boolean );
Begin
  If FFlat <> AValue Then
  Begin
    FFlat:=AValue;
    Invalidate;
  End;
End;

procedure TmxPanel.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  with Params do
    WindowClass.Style := WindowClass.Style and not (CS_HREDRAW or CS_VREDRAW);
end;

procedure TmxPanel.Paint;
var
  Flags: longint;
  X:     integer;
  CalcRect, ARect: TRect;
begin
  inherited;
  ARect := ClientRect;

  Canvas.Pen.Style   := psSolid;
  Canvas.Pen.Mode    := pmCopy;
//  Canvas.Pen.Color   := clWindowFrame;
  Canvas.Pen.Width   := 1;
//  Canvas.Brush.Color := clBtnFace;
  Canvas.Brush.Style := bsClear;

  If Not FFlat Then
    Frame3D(Canvas, ARect, clBtnHighlight, clBtnShadow, 1);

  Canvas.Font.Assign(Font);
  CalcRect := ARect;
  Flags    := DrawTextBiDiModeFlags(DT_EXPANDTABS or DT_CENTER);
{$WARNINGS OFF}
  DrawText(Canvas.Handle, pchar(Text), Length(Text), CalcRect,
    Flags or DT_CALCRECT);
{$WARNINGS ON}

  X := CalcRect.Bottom - CalcRect.Top;
  if (ARect.Bottom - ARect.Top) > X then
  begin
    ARect.Top    := ARect.Top + ((ARect.Bottom - ARect.Top - X) div 2);
    ARect.Bottom := ARect.Top + (CalcRect.Bottom - CalcRect.Top);
{$WARNINGS OFF}
    DrawText(Canvas.Handle, pchar(Text), Length(Text), ARect, Flags);
{$WARNINGS ON}
  end;
end;

Constructor TmxHints.Create;
Begin
     Inherited Create;
{$IFDEF FRENCH}
     FHints[ 0 ] := 'Mois prec';
     FHints[ 1 ] := 'Mois Suiv.';
     FHints[ 2 ] := 'Année prec.';
     FHints[ 3 ] := 'Année Suiv.';
     FHints[ 4 ] := 'Définir à aujourd''hui';
     FHints[ 5 ] := 'Aujourd''hui';
     FHints[ 6 ] := 'Effacer la sélection';
{$ELSE}
     FHints[ 0 ] := 'Go to next month';
     FHints[ 1 ] := 'Go to previous month';
     FHints[ 2 ] := 'Go to next year';
     FHints[ 3 ] := 'Go to previous year';
     FHints[ 4 ] := 'Set date to today';
     FHints[ 5 ] := 'This date is today';
     FHints[ 6 ] := 'Clear date selection';
{$ENDIF}
End;

Procedure TmxHints.Assign( Source: TPersistent );
Var
     I: Integer;
Begin
     If Source Is TmxHints Then
     Begin
          For I := 0 To 6 Do FHints[ I ] := TmxHints( Source ).Hints[ I ];
     End
     Else Inherited Assign( Source );
End;

Function TmxHints.GetHint( Index: Integer ): ShortString;
Begin
     Result := FHints[ Index ];
End;

Procedure TmxHints.SetHint( Index: Integer; AValue: ShortString );
Begin
     If FHints[ Index ] <> AValue Then
     Begin
          FHints[ Index ] := AValue;
          Change;
     End;
End;

Procedure TmxHints.Change;
Begin
     If Assigned( FOnChange ) Then FOnChange( Self );
End;

{TmxMessages}

Constructor TmxMessages.Create;
Var
     I: Integer;
Begin
     Inherited Create;
     FMessages[ 0 ] := 'Sem';
     For I := 1 To 7 Do FMessages[ I ] := ShortDayNames[ I ];
     For I := 1 To 12 Do FMonthNames[ I ] := LongMonthNames[ I ];
     Hints := TmxHints.Create;
End;

Destructor TmxMessages.Destroy;
Begin
     FHints.Free;
     Inherited Destroy;
End;

Procedure TmxMessages.Assign( Source: TPersistent );
Var
     I: Integer;
Begin
     If Source Is TmxMessages Then
     Begin
          For I := 0 To 7 Do FMessages[ I ] := TmxMessages( Source ).Messages[ I ];
          For I := 1 To 12 Do FMonthNames[ I ] := TmxMessages( Source ).MonthNames[ I ];

          FHints.Assign( TmxMessages( Source ).Hints );
     End
     Else Inherited Assign( Source );
End;

Function TmxMessages.GetMessage( Index: Integer ): ShortString;
Begin
     Result := FMessages[ Index ];
End;

Procedure TmxMessages.SetMessage( Index: Integer; AValue: ShortString );
Begin
     If FMessages[ Index ] <> AValue Then
     Begin
          FMessages[ Index ] := AValue;
          Change;
     End;
End;

Function TmxMessages.GetMonthName( Index: Integer ): ShortString;
Begin
     Result := FMonthNames[ Index ];
End;

Procedure TmxMessages.SetOnChange( AOnChange: TNotifyEvent );
Begin
     FOnChange := AOnChange;
     FHints.OnChange := AOnChange;
End;

Procedure TmxMessages.SetMonthName( Index: Integer; AValue: ShortString );
Begin
     If FMonthNames[ Index ] <> AValue Then
     Begin
          FMonthNames[ Index ] := AValue;
          Change;
     End;
End;

Procedure TmxMessages.Change;
Begin
     If Assigned( FOnChange ) Then FOnChange( Self );
End;

{TmxCalendarButton}

Procedure TmxCalendarButton.CMDesignHitTest( Var Msg: TCMDesignHitTest );
Begin
     If Parent.CanFocus And
          PtInRect( Rect( 0, 0, Width, Height ), SmallPointToPoint( Msg.Pos ) ) Then Msg.Result := 1;
End;

{TmxCustomCalendar}

Constructor TmxCustomCalendar.Create( AOwner: TComponent );
Var
     I, X: Byte;
Begin
     Inherited Create( AOwner );
     ControlStyle := ControlStyle - [ csSetCaption ];

     Height := 230;
     Width := 250;
     Caption := '';
     //
     FMinDate := 0;
     FMaxDate := 0;
     //
     BorderWidth := 1;
     BevelInner := bvNone;
     BevelOuter := bvLowered;

     OnResize := OnResizePanels;
     OnCanResize := OnCanResizePanels;

     FIndirectColor := Color;
     FSelectedColor := clNavy;
     FDayColor := clSilver;
     FDayNameColor := clSilver;
     FWeekColor := clSilver;
     FTodayColor := clblack;
     FOptions := [
          coShowDateLabel,
          coShowHeader,
          coShowFooter,
          coShowWeekDays,
          coShowWeeks,
          coFlatButtons,
          coTransparentButtons,
          coYearButtonVisible,
          coMonthButtonVisible,
          coTodayButtonVisible,
          csUseWeekEndFont,
          csUseWeekEndColor,
          coShowNextMonth,
          coShowPreviousMonth
          ];

     Messages := TmxMessages.Create;
     Messages.OnChange := OnChangeMessages;

     FSelectedFont := TFont.Create;
     FSelectedFont.Color := clWhite;
     FSelectedFont.OnChange := OnChangeMessages;

     FTodayFont := TFont.Create;
     FTodayFont.color := clWhite;
     FTodayFont.OnChange := OnChangeMessages;

     FIndirectFont := TFont.Create;
     FIndirectFont.OnChange := OnChangeMessages;

     FDayFont := TFont.Create;
     FDayFont.OnChange := OnChangeMessages;

     FWeekDaysFont := TFont.Create;
     FWeekDaysFont.Color := clblack;
     FWeekDaysFont.OnChange := OnChangeMessages;

     FWeeksFont := TFont.Create;
     FWeeksFont.Color := clblack;
     FWeeksFont.OnChange := OnChangeMessages;

     FWeekEndFont := TFont.Create;
     FWeekEndFont.Color := clBlack;
     FWeekEndFont.OnChange := OnChangeMessages;
     FWeekEndColor := clAqua;

     FFirstDayOfWeek := fdMonday;
     FHeaderFormat := hfMMMMYYYY;

     FDateFormat := 'dd.mm.yyyy';

  FPanel_Header := TmxPanel.Create(Self);
     With FPanel_Header Do
     Begin
          Parent := Self;
          Height := 26;
          Align := alTop;
          Alignment := taCenter;
          BevelInner := bvNone;
          BevelOuter := bvRaised;
          Font.Style := [ fsBold ];
     End;

  FPanel_Footer := TmxPanel.Create(Self);
     With FPanel_Footer Do
     Begin
          Parent := Self;
          Height := 26;
          Align := alBottom;
          Alignment := taCenter;
          BevelInner := bvNone;
          BevelOuter := bvRaised;
     End;

     FDateLabel := TLabel.Create( Self );
     With FDateLabel Do
     Begin
          Parent := FPanel_Footer;
          Left := 3;
          Font.Style := [ fsBold ];
          SetDateLabelFont( Font );
     End;

  FPanel_Calendar := TmxPanel.Create(Self);
     With FPanel_Calendar Do
     Begin
          Parent := Self;
          Top := FPanel_Header.Top + FPanel_Header.Height + 1;
          Align := alClient;
          Alignment := taCenter;
          BevelInner := bvNone;
          BevelOuter := bvRaised;
          ParentColor := True;
     End;

     FYear_Plus := TmxCalendarButton.Create( Self );
     With FYear_Plus Do
     Begin
          Parent := FPanel_Header;
          Width := NavigateButtonWidth;
          Top := 2;
          Caption := '>>';
          Flat := True;
          OnClick := YearButtonClick;
     End;

     FYear_Minus := TmxCalendarButton.Create( Self );
     With FYear_Minus Do
     Begin
          Parent := FPanel_Header;
          Width := NavigateButtonWidth;
          Top := 2;
          Caption := '<<';
          Flat := True;
          OnClick := YearButtonClick;
     End;

     FMonth_Plus := TmxCalendarButton.Create( Self );
     With FMonth_Plus Do
     Begin
          Parent := FPanel_Header;
          Width := NavigateButtonWidth;
          Top := 2;
          Caption := '>';
          Flat := True;
          OnClick := MonthButtonClick;
     End;

     FMonth_Minus := TmxCalendarButton.Create( Self );
     With FMonth_Minus Do
     Begin
          Parent := FPanel_Header;
          Width := NavigateButtonWidth;
          Top := 2;
          Caption := '<';
          Flat := True;
          OnClick := MonthButtonClick;
     End;

//
     FSelectCurrMonth := TmxCalendarButton.Create( Self );
     With FSelectCurrMonth Do
     Begin
          Parent := FPanel_Header;
          Width := MonthButtonWidth;
          Top := 2;
          Caption := '';
          Flat := false;
          Transparent := true;
          OnClick := CurrMonthButtonClick;
     End;
//
     FTodayButton := TmxCalendarButton.Create( Self );
     With FTodayButton Do
     Begin
          Parent := FPanel_Footer;
          Width := NavigateButtonWidth;
          Top := 2;
          glyph.LoadFromResourceName (HInstance,'TODAYBUTTON');
//          Caption := 'T';
          Flat := True;
          OnClick := TodayButtonClick;
     End;

     FClearButton := TmxCalendarButton.Create( Self );
     With FClearButton Do
     Begin
          Parent := FPanel_Footer;
          Width := NavigateButtonWidth;
          Top := 2;
          //Caption := 'C';
          glyph.LoadFromResourceName (HInstance,'CLEARBUTTON');
          Flat := True;
          Visible := False;
          OnClick := ClearButtonClick;
     End;

     For I := 0 To 6 Do
     Begin
          For X := 0 To 7 Do
          Begin
      				 FPanel_Days[I, X] := TmxPanel.Create(Self);
               With FPanel_Days[ I, X ] Do
               Begin
                    Parent := FPanel_Calendar;
                    Alignment := taCenter;
                    CoordX := I;
                    CoordY := X;
                    If I = 0 Then
                    begin
                      Color := FWeekColor;
                    end Else
                    begin
                       If X = 0 Then
                       begin
                        Color := DayNameColor;
                        OnMouseDown := doWeekSel;
                       end Else
                       begin
                        Color := FDayColor;
                       end;
                    end;

                    If ( X <> 0 ) And ( I <> 0 ) Then
                    Begin
                      OnMouseDown := DoMouseDown;
                    End;
               End;
          End;
     End;

     FSelectionStart := SysUtils.Date;
     FSelectionEnd := SysUtils.Date;

     _SetDate( SysUtils.Date );

     FVersion := mxCalendarVersion;
End;

Destructor TmxCustomCalendar.Destroy;
Var
     I, X: Byte;
Begin
     FDateLabel.Free;
     FMessages.Free;
     FSelectedFont.Free;
     FTodayFont.Free;
     FDayFont.Free;
     FWeekDaysFont.Free;
     FWeeksFont.Free;
     FWeekEndFont.Free;
     FIndirectFont.Free;

     For I := 0 To 6 Do For X := 0 To 7 Do FPanel_Days[ I, X ].Free;

     FYear_Plus.Free;
     FYear_Minus.Free;
     FMonth_Plus.Free;
     FMonth_Minus.Free;
     FPanel_Calendar.Free;
     FPanel_Header.Free;
     FPanel_Footer.Free;

     Inherited Destroy;
End;

Procedure TmxCustomCalendar.Loaded;
Begin
     Inherited;

     SetYearPlusGlyph( FYear_Plus.Glyph );
     SetYearMinusGlyph( FYear_Minus.Glyph );
     SetMonthPlusGlyph( FMonth_Plus.Glyph );
     SetMonthMinusGlyph( FMonth_Minus.Glyph );
     SetTodayButtonGlyph( FTodayButton.Glyph );
     SetClearButtonGlyph( FClearButton.Glyph );

     If csSetTodayOnStartup In FOptions Then _SetDate( SysUtils.Date );

     RepaintCalendar;
End;

Procedure TmxCustomCalendar.SetVersion( AValue: String );
Begin
    // *** Does nothing ***
End;

Function TmxCustomCalendar.GetVersion: String;
Begin
{$WARNINGS OFF}
     Result := Format( '%d.%d', [ Hi( FVersion ), Lo( FVersion ) ] );
{$WARNINGS ON}
End;

Procedure TmxCustomCalendar.SetButtonCaptions;
Var
     X, I, Y: Shortint;
Begin
     FPanel_Days[ 0, 0 ].Caption := Messages[ 0 ];
     For I := 0 To 6 Do
     Begin
          FPanel_Days[ I, 0 ].Font.Assign( FWeeksFont );
     End;

     X := Byte( FFirstDayOfWeek );
     Y := 0;

     For I := X To 6 Do
     Begin
          Inc( Y );
          FPanel_Days[ 0, Y ].Caption := Messages[ I + 1 ];
          FPanel_Days[ 0, Y ].Font.Assign( FWeekDaysFont );
     End;

     For I := 0 To X - 1 Do
     Begin
          Inc( Y );
          FPanel_Days[ 0, Y ].Caption := Messages[ I + 1 ];
          FPanel_Days[ 0, Y ].Font.Assign( FWeekDaysFont );
     End;

     FMonth_Minus.Hint := FMessages.Hints[ 0 ];
     FMonth_Plus.Hint := FMessages.Hints[ 1 ];
     FYear_Minus.Hint := FMessages.Hints[ 2 ];
     FYear_Plus.Hint := FMessages.Hints[ 3 ];
     FTodayButton.Hint := FMessages.Hints[ 4 ];
     FClearButton.Hint := FMessages.Hints[ 6 ];
End;

Procedure TmxCustomCalendar.SetButtonPositions;
Var
     I, X: Byte;
     PanelHeight: Integer;
     PanelWidth: Integer;
     DifferenceWidth: Byte;
     DifferenceHeight: Byte;
     NumberOfRows: Byte;
     NumberOfCols: Byte;
     Correction: Byte;
     ButtonPosition: Integer;
Begin
     ButtonPosition := FPanel_Footer.ClientWidth - FTodayButton.ClientWidth - 3;
     If Not FTodayButton.Visible Then ButtonPosition := -100;
     FTodayButton.Left := ButtonPosition;

     ButtonPosition := FTodayButton.Left - FClearButton.Width - 3;
     If Not FClearButton.Visible Then ButtonPosition := -100;
     FClearButton.Left := ButtonPosition;

     ButtonPosition := FPanel_Header.ClientWidth - FYear_Plus.ClientWidth - 3;
     If Not FYear_Plus.Visible Then ButtonPosition := -100;
     FYear_Plus.Left := ButtonPosition;

     ButtonPosition := 3;
     If Not FYear_Minus.Visible Then ButtonPosition := -100;
     FYear_Minus.Left := ButtonPosition;

     If FYear_Plus.Visible Then
          ButtonPosition := FYear_Plus.Left - FMonth_Plus.Width - 2 Else
          ButtonPosition := FPanel_Header.ClientWidth - FMonth_Plus.ClientWidth - 3;

     If Not FMonth_Plus.Visible Then ButtonPosition := -100;
     FMonth_Plus.Left := ButtonPosition;

     If FYear_Minus.Visible Then
          ButtonPosition := FYear_Minus.Left + FYear_Minus.Width + 2 Else
          ButtonPosition := 3;

     If Not FMonth_Minus.Visible Then ButtonPosition := -100;
     FMonth_Minus.Left := ButtonPosition;

//
     if FSelectCurrMonth.visible then
     begin
        ButtonPosition := (FPanel_Header.ClientWidth div 2) - 50 ;
        FSelectCurrMonth.left := ButtonPosition;
     end;
//

     NumberOfRows := 7;
     NumberOfCols := 8;

     If Not ( coShowWeeks In Options ) Then Dec( NumberOfCols );
     If Not ( coShowWeekDays In Options ) Then Dec( NumberOfRows );

     PanelHeight := FPanel_Calendar.ClientHeight Div NumberOfRows;
     PanelWidth := FPanel_Calendar.ClientWidth Div NumberOfCols;

     DifferenceHeight := ( ( FPanel_Calendar.ClientHeight ) Mod NumberOfRows ) Div 2;
     DifferenceWidth := ( ( FPanel_Calendar.ClientWidth ) Mod NumberOfCols ) Div 2;

     For I := 0 To 6 Do
     Begin
          For X := 0 To 7 Do
          Begin
               With FPanel_Days[ I, X ] Do
               Begin
                    Width := PanelWidth;
                    Height := PanelHeight;

                    If ( coShowWeeks In Options ) Then Correction := 0 Else Correction := 1;

                    If ( coShowWeeks In Options ) Or ( Not ( coShowWeeks In Options ) And ( X <> 0 ) ) Then
                         Left := ( ( X - Correction ) * PanelWidth ) + DifferenceWidth Else
                         Left := -100;

                    If ( coShowWeekDays In Options ) Then Correction := 0 Else Correction := 1;

                    If ( coShowWeekDays In Options ) Or ( Not ( coShowWeekDays In Options ) And ( I <> 0 ) ) Then
                         Top := ( ( I - Correction ) * PanelHeight ) + DifferenceHeight Else
                         Top := -100;
               End;
          End;
     End;
End;

Procedure TmxCustomCalendar.OnResizePanels( Sender: TObject );
Begin
     SetButtonPositions;
End;

Procedure TmxCustomCalendar.OnChangeMessages( Sender: TObject );
Begin
     RepaintCalendar;
End;

Procedure TmxCustomCalendar.OnCanResizePanels( Sender: TObject; Var NewWidth, NewHeight: Integer; Var Resize: Boolean );
Begin
     If NewWidth < ( NavigateButtonWidth * 4 ) + 6 Then
     Begin
          NewWidth := ( NavigateButtonWidth * 4 ) + 16;
          Resize := True;
     End;
End;

{$IFDEF DELPHI4_UP}

Procedure TmxCustomCalendar.CMBorderChanged( Var Message: TMessage );
Begin
     Inherited;
     RepaintCalendar;
End;

{$ENDIF}

Procedure TmxCustomCalendar.CMCtl3DChanged( Var Message: TMessage );
Begin
     If NewStyleControls And ( BorderStyle = bsSingle ) Then RecreateWnd;
     Inherited;
End;

Procedure TmxCustomCalendar.CMSysColorChange( Var Message: TMessage );
Begin
     Inherited;
     If Not ( csLoading In ComponentState ) Then
     Begin
          Message.Msg := WM_SYSCOLORCHANGE;
          DefaultHandler( Message );
     End;
End;

Procedure TmxCustomCalendar.CMEnabledChanged( Var Message: TMessage );
Begin
     Inherited;
     RepaintCalendar;
End;

Procedure TmxCustomCalendar.CMColorChanged( Var Message: TMessage );
Begin
     Inherited;
     RecreateWnd;
End;

Procedure TmxCustomCalendar.CMFontChanged( Var Msg: TMessage );
Begin
     Inherited;
     RepaintCalendar;
End;

Procedure TmxCustomCalendar.CreateParams( Var Params: TCreateParams );
Const
     BorderStyles: Array[ TBorderStyle ] Of DWORD = ( 0, WS_BORDER );
Begin
     Inherited CreateParams( Params );
     With Params Do
     Begin
          Style := Style Or BorderStyles[ BorderStyle ];
          If NewStyleControls And Ctl3D And ( BorderStyle = bsSingle ) Then
          Begin
               Style := Style And Not WS_BORDER;
               ExStyle := ExStyle Or WS_EX_CLIENTEDGE;
          End;
          WindowClass.Style := WindowClass.Style And Not ( CS_HREDRAW Or CS_VREDRAW );
     End;
End;

Procedure TmxCustomCalendar.Paint;
Begin
     Inherited;
     RepaintCalendar;
End;

Procedure TmxCustomCalendar.SetSelectedColor( AValue: TColor );
Begin
     If FSelectedColor <> AValue Then
     Begin
          FSelectedColor := AValue;
          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.SetWeekEndColor( AValue: TColor );
Begin
     If FWeekEndColor <> AValue Then
     Begin
          FWeekEndColor := AValue;
          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.SetHeaderColor( AValue: TColor );
Begin
     If FPanel_Header.Color <> AValue Then FPanel_Header.Color := AValue;
End;

Function TmxCustomCalendar.GetHeaderColor: TColor;
Begin
     Result := FPanel_Header.Color;
End;

Procedure TmxCustomCalendar.SetFooterColor( AValue: TColor );
Begin
     If FPanel_Footer.Color <> AValue Then FPanel_Footer.Color := AValue;
End;

Function TmxCustomCalendar.GetFooterColor: TColor;
Begin
     Result := FPanel_Footer.Color;
End;

Procedure TmxCustomCalendar.SetDayNameColor( AValue: TColor );
Var
     I: Byte;
Begin
     If FDayNameColor <> AValue Then
     Begin
          FDayNameColor := AValue;
          For I := 1 To 7 Do FPanel_Days[ 0, I ].Color := FDayNameColor;
     End;
End;

Procedure TmxCustomCalendar.SetWeekColor( AValue: TColor );
Var
     I: Byte;
Begin
     If FWeekColor <> AValue Then
     Begin
          FWeekColor := AValue;
          For I := 0 To 6 Do FPanel_Days[ I, 0 ].Color := FWeekColor;
     End;
End;

Procedure TmxCustomCalendar.SetTodayColor( AValue: TColor );
Begin
     If FTodayColor <> AValue Then
     Begin
          FTodayColor := AValue;
          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.SetIndirectColor( AValue: TColor );
Begin
     If FIndirectColor <> AValue Then
     Begin
          FIndirectColor := AValue;
    RepaintCalendar;
  end;
end;

Procedure TmxCustomCalendar.SetDayColor( AValue: TColor );
Var
     I, X: Byte;
Begin
     If FDayColor <> AValue Then
     Begin
          FDayColor := AValue;

          For I := 1 To 6 Do
          Begin
               For X := 1 To 7 Do
               Begin
                    FPanel_Days[ I, X ].Color := FDayColor;
               End;
          End;

          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.SetOptions( AValue: TmxCalendarOptions );
Var
     X, Y: Integer;
     ThePanel : TmxPanel;
Begin
     If FOptions <> AValue Then
     Begin
          FOptions := AValue;

          if csWeekSelectionOnly in Foptions then
          begin
            FSelectCurrMonth.Enabled := true;
            FSelectCurrMonth.flat := true;
            FSelectCurrMonth.OnClick := nil;
            FTodayButton.enabled := false;
            FClearButton.enabled := false;
            ThePanel := FindWeek(NumSemaine(SysUtils.Date));
            ThePanel.MouseDown (mbLeft,[ssLeft],0,0);
          end else
          begin
            FSelectCurrMonth.Enabled := True;
            FTodayButton.enabled := True;
            FClearButton.enabled := True;
            ThePanel := FindDay(SysUtils.Date);
            if ThePanel <> nil then ThePanel.mouseDown (mbLeft,[],0,0);
          end;

          If Not ( coShowHeader In Options ) Then
          Begin
               FPanel_Header.Align := alNone;
               FPanel_Header.Top := -1000;
          End
          Else FPanel_Header.Align := alTop;

          If Not ( coShowFooter In Options ) Then
          Begin
               FPanel_Footer.Align := alNone;
               FPanel_Footer.Top := -1000;
          End
          Else FPanel_Footer.Align := alBottom;

          FYear_Plus.Flat := coFlatButtons In FOptions;
          FYear_Minus.Flat := coFlatButtons In FOptions;
          FMonth_Plus.Flat := coFlatButtons In FOptions;
          FMonth_Minus.Flat := coFlatButtons In FOptions;
          FTodayButton.Flat := coFlatButtons In FOptions;

          FYear_Plus.Transparent := coTransparentButtons In FOptions;
          FYear_Minus.Transparent := coTransparentButtons In FOptions;
          FMonth_Plus.Transparent := coTransparentButtons In FOptions;
          FMonth_Minus.Transparent := coTransparentButtons In FOptions;

          FYear_Plus.Visible := coYearButtonVisible In FOptions;
          FYear_Minus.Visible := coYearButtonVisible In FOptions;
          FMonth_Plus.Visible := coMonthButtonVisible In FOptions;
          FMonth_Minus.Visible := coMonthButtonVisible In FOptions;
          FTodayButton.Visible := coTodayButtonVisible In FOptions;
          FClearButton.Visible := coClearButtonVisible In FOptions;

          SetDateLabelFont( FDateLabel.Font );

          For X := 1 To 6 Do
               For Y := 1 To 7 Do
        FPanel_Days[X, Y].Flat:=csFlatDays in FOptions;

          For X := 0 To 6 Do
      FPanel_Days[X, 0].Flat:=coFlatHeaders in FOptions;

          For Y := 1 To 7 Do
      FPanel_Days[0,Y].Flat:=coFlatHeaders in FOptions;

          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.SetWeekDaysFont( AValue: TFont );
Begin
     FWeekDaysFont.Assign( AValue );
End;

Procedure TmxCustomCalendar.SetIndirectFont( AValue: TFont );
Begin
     FIndirectFont.Assign( AValue );
End;

Procedure TmxCustomCalendar.SetSelectedFont( AValue: TFont );
Begin
     FSelectedFont.Assign( AValue );
End;

Procedure TmxCustomCalendar.SetTodayFont( AValue: TFont );
Begin
     FTodayFont.Assign( AValue );
End;

Function TmxCustomCalendar.GetDateLabelFont: TFont;
Begin
     Result := FDateLabel.Font;
End;

Procedure TmxCustomCalendar.SetDateLabelFont( AValue: TFont );
Begin
     FDateLabel.Font.Assign( AValue );

     If coShowDateLabel In FOptions Then
          FDateLabel.Top := ( FPanel_Footer.Height - FDateLabel.Height ) Div 2 Else
          FDateLabel.Top := -100;
End;

Procedure TmxCustomCalendar.SetDayFont( AValue: TFont );
Begin
     FDayFont.Assign( AValue );
End;

Procedure TmxCustomCalendar.SetWeeksFont( AValue: TFont );
Begin
     FWeeksFont.Assign( AValue );
End;

Procedure TmxCustomCalendar.SetWeekEndFont( AValue: TFont );
Begin
     FWeekEndFont.Assign( AValue );
End;

Function TmxCustomCalendar.GetYearMinusGlyph: TBitmap;
Begin
     Result := FYear_Minus.Glyph;
End;

Procedure TmxCustomCalendar.SetYearMinusGlyph( AValue: TBitmap );
Begin
     With FYear_Minus Do
     Begin
          Glyph.Assign( AValue );
          If Glyph.Empty Then
               Caption := '<<' Else
               Caption := '';
     End;
End;

Function TmxCustomCalendar.GetTodayButtonGlyph: TBitmap;
Begin
     Result := FTodayButton.Glyph;
End;

Procedure TmxCustomCalendar.SetTodayButtonGlyph( AValue: TBitmap );
Begin
     With FTodayButton Do
     Begin
          Glyph.Assign( AValue );
          If Glyph.Empty Then
               Caption := 'T' Else
               Caption := '';
     End;
End;

Function TmxCustomCalendar.GetClearButtonGlyph: TBitmap;
Begin
     Result := FClearButton.Glyph;
End;

Procedure TmxCustomCalendar.SetClearButtonGlyph( AValue: TBitmap );
Begin
     With FClearButton Do
     Begin
          Glyph.Assign( AValue );
          If Glyph.Empty Then
               Caption := 'C' Else
               Caption := '';
     End;
End;

Function TmxCustomCalendar.GetYearPlusGlyph: TBitmap;
Begin
     Result := FYear_Plus.Glyph;
End;

Procedure TmxCustomCalendar.SetYearPlusGlyph( AValue: TBitmap );
Begin
     With FYear_Plus Do
     Begin
          Glyph.Assign( AValue );
          If Glyph.Empty Then
               Caption := '>>' Else
               Caption := '';
     End;
End;

Function TmxCustomCalendar.GetMonthMinusGlyph: TBitmap;
Begin
     Result := FMonth_Minus.Glyph;
End;

Procedure TmxCustomCalendar.SetMonthMinusGlyph( AValue: TBitmap );
Begin
     With FMonth_Minus Do
     Begin
          Glyph.Assign( AValue );
          If Glyph.Empty Then
               Caption := '<' Else
               Caption := '';
     End;
End;

Function TmxCustomCalendar.GetMonthPlusGlyph: TBitmap;
Begin
     Result := FMonth_Plus.Glyph;
End;

Procedure TmxCustomCalendar.SetMonthPlusGlyph( AValue: TBitmap );
Begin
     With FMonth_Plus Do
     Begin
          Glyph.Assign( AValue );
          If Glyph.Empty Then
               Caption := '>' Else
               Caption := '';
     End;
End;

Procedure TmxCustomCalendar.SetFirstDayOfWeek( AValue: TmxFirstDayOfWeek );
Begin
     If FFirstDayOfWeek <> AValue Then
     Begin
          FFirstDayOfWeek := AValue;
    RecreateWnd;
//          RepaintCalendar;
     End;
End;

Function TmxCustomCalendar.DaysInMonth( AYear, AMonth: Integer ): Integer;
Const
     NumberOfDays: Array[ 1..12 ] Of Integer = ( 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );
Begin
     Result := 0;
     If Not ( AMonth In [ 1..12 ] ) Then Exit;
     Result := NumberOfDays[ AMonth ];
     If ( AMonth = 2 ) And IsLeapYear( AYear ) Then Inc( Result );
End;

Function TmxCustomCalendar.MonthToWeek( AMonth: Integer ): Integer;
Var
     I, _DayOfWeek: Integer;
Begin
     Result := 0;
//     For I := 1 To FMonth - 1 Do Inc( Result, DaysInMonth( FYear, I ) );
     For I := 1 To AMonth - 1 Do
     begin
			Inc( Result, DaysInMonth( FYear, I ) );
     end;

     _DayOfWeek := SetWeekStart( DayOfWeek( EncodeDate( FYear, 1, 1 ) ) );
     If _DayOfWeek > 4 Then Dec( Result, 7 - _DayOfWeek ) Else Inc( Result, _DayOfWeek - 1 );

     Result := ( Result Div 7 ) + 1;
     If ( FFirstDayOfWeek = fdSunday ) And ( DayOfWeek( EncodeDate( FYear, AMonth, 1 ) ) = 7 ) Then Dec( Result );
End;

Function TmxCustomCalendar.SetWeekStart( ADayIndex: Integer ): Integer;
Begin
     Result := ADayIndex;
     If FFirstDayOfWeek = fdMonday Then If Result = 1 Then Result := 7 Else Dec( Result );
End;

Function TmxCustomCalendar.WeeksInYear( AYear: Integer ): Integer;
Var
     _DayOfWeek: Integer;
Begin
     If IsLeapYear( AYear ) Then Result := 366 Else Result := 365;

     _DayOfWeek := SetWeekStart( DayOfWeek( EncodeDate( AYear, 1, 1 ) ) );
     If _DayOfWeek > 4 Then
          Dec( Result, _DayOfWeek ) Else
          Inc( Result, 7 - _DayOfWeek );

     _DayOfWeek := SetWeekStart( DayOfWeek( EncodeDate( AYear, 12, 31 ) ) );
     If _DayOfWeek > 3 Then
          Inc( Result, 7 - _DayOfWeek ) Else
          Dec( Result, _DayOfWeek );

     Result := Result Div 7;
     If Result = 51 Then Result := 52;
End;

Procedure TmxCustomCalendar.SetYear( AValue: Integer );
Begin
     If AValue <> FYear Then
     Begin
          If AValue < 1900 Then
               FYear := 1900 Else
               If AValue > 2100 Then
                    FYear := 2100 Else
                    FYear := AValue;

    if not (csLoading in ComponentState) then
          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.SetMonth( AValue: Integer );
Begin
     If AValue <> FMonth Then
     Begin
          If AValue < 1 Then
               FMonth := 1 Else
               If AValue > 12 Then
                    FMonth := 12 Else
                    FMonth := AValue;

          If DaysInMonth( FYear, FMonth ) < FDay Then FDay := DaysInMonth( FYear, FMonth );

    if not (csLoading in ComponentState) then
          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.SetDay( AValue: Integer );
Begin
     If AValue <> FDay Then
     Begin
          If AValue < 1 Then
               FDay := 1 Else
               If AValue > DaysInMonth( FYear, FMonth ) Then
                    FDay := DaysInMonth( FYear, FMonth ) Else
                    FDay := AValue;

    if not (csLoading in ComponentState) then
          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar._SetDate( AValue: TDateTime );
Var
     Year, Month, Day: Word;
Begin
     DecodeDate( AValue, Year, Month, Day );
     FYear := Year;
     FMonth := Month;
     FDay := Day;
End;

Procedure TmxCustomCalendar.SetDate( AValue: TDateTime );
Begin
     _SetDate( AValue );
     RepaintCalendar;
End;

Function TmxCustomCalendar.GetDate: TDateTime;
Begin
     Result := EncodeDate( FYear, FMonth, FDay );
End;

Procedure TmxCustomCalendar.SetHeaderFormat( AValue: TmxHeaderFormat );
Begin
     If AValue <> FHeaderFormat Then
     Begin
          FHeaderFormat := AValue;
          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.SetPanelColor( APanel: TmxPanel; AYear, AMonth: Integer );
Var
     PanelDate: TDateTime;
     IsSelected,IsInMinMax: Boolean;
     MonthNumber: Byte;
Begin
     APanel.Tag := AMonth;

     PanelDate := EncodeDate( AYear, AMonth, StrToInt( APanel.Caption ) );
     Apanel.FDate := PanelDate;
     Apanel.Fmonth := Amonth;
     IsSelected := ( PanelDate >= Trunc( FSelectionStart ) ) And ( PanelDate <= FSelectionEnd );
     if fmindate > 0 then
     begin
      IsInMinMax := ( PanelDate >= Trunc( fMinDate ) ) And ( PanelDate <= fMaxDate )
     end else IsInMinMax := true;

     if not IsInMinMax then
     BEGIN
      Apanel.caption := '';
      Apanel.OnMouseDown := nil;
     END;

     APanel.Color := FDayColor;
     APanel.Hint := '';

     If IsSelected Then
     Begin
          APanel.Color := FSelectedColor;
          APanel.Font.Assign( FSelectedFont );
     End
     Else
     Begin
          If PanelDate = SysUtils.Date Then
          Begin
             APanel.Color := FTodayColor;
             APanel.Hint := FMessages.Hints[ 5 ];
             APanel.Font.Assign( FTodayFont );
             if not IsinMinMax then APanel.font.Style := APanel.font.Style + [fsStrikeOut];
          End
          Else
          Begin
               If AMonth = FMonth Then
               Begin
                    APanel.Color := FDayColor;
                    APanel.Font.Assign( FDayFont );

                    If DayOfWeek( PanelDate ) In [ 1, 7 ] Then
                    Begin
                         If csUseWeekEndFont In FOptions Then
                              APanel.Font.Assign( FWeekEndFont );

                         If csUseWeekEndColor In FOptions Then
                              APanel.Color := FWeekEndColor;
                    End;
               End
               Else
               Begin
                    APanel.Color := FIndirectColor;
                    APanel.Font.Assign( FIndirectFont );

                    MonthNumber := FMonth + 1;
                    If MonthNumber > 12 Then MonthNumber := 1;

                    If ( Not ( coShowNextMonth In FOptions ) ) And ( AMonth = MonthNumber ) Then
                         APanel.Font.Color := APanel.Color;

                    MonthNumber := FMonth - 1;
                    If MonthNumber < 1 Then MonthNumber := 12;

                    If ( Not ( coShowPreviousMonth In FOptions ) ) And ( AMonth = MonthNumber ) Then
                         APanel.Font.Color := APanel.Color;
               End;
          End;
     End;
  APanel.Invalidate;
End;

function TmxCustomCalendar.TheDateInMonth ( Decalage : integer ): TdateTime;
var Indice : integer;
		Zyear,ZMonth,ZDay : Word;
begin
	for Indice := 1 to 7 do
  begin
//		if FPanel_Days[ Decalage , indice ].Fmonth = Fmonth then
    begin
    	result := FPanel_Days[ Decalage , indice ].FDate;
      DecodeDate (result,Zyear,Zmonth,ZDay);
    end;
  end;
end;

Procedure TmxCustomCalendar.RepaintCalendar;
Var
     FirstDate: TDateTime;
     TotalDays: Integer;
     WeekDay: Integer;
     CurrentDay: Integer;
     WeekNumber: Integer;
     WeeksYear: Integer;
     X, Y: Integer;
     SMonth: ShortString;
     TheDate : TdateTime;
     FisrtWeekofMonth : integer;
Begin
     SetButtonPositions;
     SetButtonCaptions;

     // *** Update View ***
     if fMinDate = 0 then
     begin
        FirstDate := EncodeDate( FYear, FMonth, 1 );
        TotalDays := DaysInMonth( FYear, FMonth );
        CurrentDay := 1;
     end else
     begin
        _setdate (fmindate);
        FirstDate := EncodeDate( FYear, FMonth, 1 );
        TotalDays := DaysInMonth( FYear, FMonth );
        CurrentDay := 1;
     end;
     WeekDay := SetWeekStart( DayOfWeek( FirstDate ) );


     For X := 0 To 5 Do
     Begin
          For Y := 0 To 6 Do
          Begin
               If ( X = 0 ) And ( Y + 1 < WeekDay ) Then
               Begin
                    If FMonth = 1 Then
                    Begin
                         FPanel_Days[ X + 1, Y + 1 ].Caption := IntToStr( DaysInMonth( FYear - 1, 12 ) - WeekDay + Y + 2 );
                         SetPanelColor( FPanel_Days[ X + 1, Y + 1 ], FYear - 1, 12 );
                         FPanel_Days[ X + 1, Y + 1 ].CoordX := X+1;
                         FPanel_Days[ X + 1, Y + 1 ].CoordY := Y+1;
                    End
                    Else
                    Begin
                         FPanel_Days[ X + 1, Y + 1 ].Caption := IntToStr( DaysInMonth( FYear, FMonth - 1 ) - WeekDay + Y + 2 );
                         SetPanelColor( FPanel_Days[ X + 1, Y + 1 ], FYear, FMonth - 1 );
                         FPanel_Days[ X + 1, Y + 1 ].CoordX := X+1;
                         FPanel_Days[ X + 1, Y + 1 ].CoordY := Y+1;
                    End;
               End
               Else
               Begin
                    If CurrentDay > TotalDays Then
                    Begin
                         FPanel_Days[ X + 1, Y + 1 ].Caption := IntToStr( CurrentDay - TotalDays );
                         If FMonth = 12 Then
                              SetPanelColor( FPanel_Days[ X + 1, Y + 1 ], FYear + 1, 1 ) Else
                              SetPanelColor( FPanel_Days[ X + 1, Y + 1 ], FYear, FMonth + 1 );
                         FPanel_Days[ X + 1, Y + 1 ].CoordX := X+1;
                         FPanel_Days[ X + 1, Y + 1 ].CoordY := Y+1;
                    End
                    Else
                    Begin
                         FPanel_Days[ X + 1, Y + 1 ].Caption := IntToStr( CurrentDay );
                         SetPanelColor( FPanel_Days[ X + 1, Y + 1 ], FYear, FMonth );
                         FPanel_Days[ X + 1, Y + 1 ].CoordX := X+1;
                         FPanel_Days[ X + 1, Y + 1 ].CoordY := Y+1;
                    End;

                    Inc( CurrentDay );
               End;
          End;
     End;

     // *** Set Caption ***

     SMonth := IntToStr( FMonth );
     If Length( SMonth ) = 1 Then SMonth := '0' + SMonth;

     Case FHeaderFormat Of
          hfMMMMYYYY: FPanel_Header.Caption := Format( '%s, %d', [ FMessages.FMonthNames[ FMonth ], FYear ] );
          hfYYYYMMMM: FPanel_Header.Caption := Format( '%d, %s', [ FYear, FMessages.FMonthNames[ FMonth ] ] );
          hfMMYYYY: FPanel_Header.Caption := Format( '%s.%d', [ SMonth, FYear ] );
          hfYYYYMM: FPanel_Header.Caption := Format( '%d.%s', [ FYear, SMonth ] );
     End;

     FSelectCurrMonth.Caption := FPanel_Header.Caption;
  FPanel_Header.Invalidate;
  // *** Set Cursor


     // *** Set Week Numbers ***

     WeeksYear := WeeksInYear( FYear );

     For X := 0 To 5 Do
     Begin
     	(*
          WeekNumber := MonthToWeek( FMonth ) + X;
          If WeekNumber > WeeksYear Then WeekNumber := WeekNumber - WeeksYear;
      *)
//      		if X = 0 then
//          begin
            TheDate := TheDateInMonth (X+1);
            WeekNumber := NumSemaine(TheDate);
//          end else
//          begin
//            inc(WeekNumber);
//          end;
          FPanel_Days[ X + 1, 0 ].Caption := IntToStr( WeekNumber );
          FPanel_Days[ X + 1, 0 ].fWeek := WeekNumber;
          FPanel_Days[ X + 1, 0 ].FcoordX  := X+1;
          FPanel_Days[ X + 1, 0 ].FcoordY  := 0;
          FPanel_Days[ X + 1, 0 ].Invalidate;
          for Y := 0 to 6 do
          begin
            FPanel_Days[ X + 1, Y+1 ].fWeek := WeekNumber;
          end;
     End;

     // *** Set Label Caption ***

     If FSelectionStart = 0 Then
     Begin
          FDateLabel.Caption := FormatDateTime( FDateFormat, EncodeDate( FYear, FMonth, FDay ) );
     End
     Else
     Begin
          If FSelectionStart = FSelectionEnd Then
          Begin
               Try
                    FDateLabel.Caption := FormatDateTime( FDateFormat, FSelectionStart );
               Except
                    FDateLabel.Caption := FormatDateTime( 'yyyy.mm.dd', FSelectionStart );
               End;
          End
          Else
          Begin
               Try
                    FDateLabel.Caption := TraduireMemoire('du')+' '+
                         FormatDateTime( FDateFormat, FSelectionStart ) + ' ' +TraduireMemoire('au')+ ' ' +
                         FormatDateTime( FDateFormat, FSelectionEnd );
               Except
                    FDateLabel.Caption := TraduireMemoire('du')+' '+
                         FormatDateTime( 'yyyy.mm.dd', FSelectionStart ) + ' ' +TraduireMemoire('au')+ ' ' +
                         FormatDateTime( 'yyyy.mm.dd', FSelectionEnd );
               End;
          End;
     End;
End;

Procedure TmxCustomCalendar.MonthButtonClick( Sender: TObject );
Begin
     If Sender = FMonth_Plus Then
     Begin
          If FMonth = 12 Then
          Begin
               Inc( FYear );
               FMonth := 1;
          End
          Else Inc( FMonth )
     End
     Else
     Begin
          If FMonth = 1 Then
          Begin
               Dec( FYear );
               FMonth := 12;
          End
          Else Dec( FMonth )
     End;

     RepaintCalendar;
End;

Procedure TmxCustomCalendar.YearButtonClick( Sender: TObject );
Begin
     If Sender = FYear_Plus Then Inc( FYear ) Else Dec( FYear );
     RepaintCalendar;
End;

Procedure TmxCustomCalendar.TodayButtonClick( Sender: TObject );
Begin
     SetDate( SysUtils.Date );
End;

Procedure TmxCustomCalendar.ClearButtonClick( Sender: TObject );
Begin
     FSelectionStart := SysUtils.Date;
     FSelectionEnd := SysUtils.Date;
     RepaintCalendar;
End;

Procedure TmxCustomCalendar.SetSelectionStart( AValue: TDateTime );
Begin
     If AValue <> FSelectionStart Then
     Begin
          FSelectionStart := AValue;
          FSelectionEnd := AValue;
//          If ( FSelectionStart > FSelectionEnd ) Or ( ( Not ( csSelectionEnabled In FOptions ) ) And ( FSelectionStart <> FSelectionEnd ) ) Then
//               SetSelectionEnd( FSelectionStart );
          _SetDate( AValue );
          RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.SetSelectionEnd( AValue: TDateTime );
Begin
     If AValue <> FSelectionEnd Then
     Begin
        FSelectionEnd := AValue;
        If ( FSelectionEnd < FSelectionStart ) Or ( ( Not ( csSelectionEnabled In FOptions ) ) And ( FSelectionStart <> FSelectionEnd ) ) Then
             SetSelectionStart( FSelectionEnd );
        RepaintCalendar;
     End;
End;

Procedure TmxCustomCalendar.doWeekSel ( Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer );
var  APanel,FirstPanel,LastPanel: TmxPanel;
     TheWeekNumber : Integer;
     SavOption : TmxCalendarOptions;
begin
  APanel := ( Sender As TmxPanel );
  TheWeekNumber := Apanel.Fweek;
  fWeekSelect := TheWeekNumber;
  SavOption := Foptions;
  if csWeekSelectionOnly in FOptions then
  begin
    if ssShift in Shift then Shift := Shift - [ssShift];
    Foptions := FOptions - [csWeekSelectionOnly];
  end;
  FirstPanel := FPanel_Days[Apanel.FcoordX,1];
  FirstPanel.MouseDown (button,shift,0,0);
  Apanel := FindWeek (TheWeekNumber);
  if Apanel <> nil then
  begin
    LastPanel := FPanel_Days[Apanel.FcoordX,7];
    LastPanel.MouseDown (button,shift+[ssShift],0,0);
  end;
  Foptions := SavOption;
end;

Procedure TmxCustomCalendar.DoMouseDown( Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer );
Var
     _Year: Word;
     PanelDate: TDateTime;
  APanel:    TmxPanel;
Begin
     if csWeekSelectionOnly in FOptions then
     begin
        // on selectionne la semaine
        doWeekSel(sender,button,shift,X,Y);
        exit;
     end;

	   APanel := (Sender as TmxPanel);

     if Apanel.caption = '' then exit;

     If ( APanel.Tag = 12 ) And ( FMonth = 1 ) Then
          _Year := FYear - 1 Else
          If ( APanel.Tag = 1 ) And ( FMonth = 12 ) Then
               _Year := FYear + 1 Else _Year := FYear;

     PanelDate := Trunc( EncodeDate( _Year, APanel.Tag, StrToInt( APanel.Caption ) ) );

     If (ssLeft In Shift) and not (ssShift in shift)  Then
     begin
        SetSelectionStart( PanelDate + Frac( FSelectionStart ) );
     end;

     if (ssleft in Shift) and (ssShift in Shift) then
     begin
        If ( csSelectionEnabled In FOptions ) then
        Begin
          SetSelectionEnd( PanelDate + Frac( FSelectionEnd ) );
        End;
     end;
End;

Procedure TmxCustomCalendar.ClearSelection;
Begin
     ClearButtonClick( Self );
End;

Procedure TmxCustomCalendar.SetDateFormat( AValue: String );
Begin
     If FDateFormat <> AValue Then
     Begin
          FDateFormat := AValue;
          RepaintCalendar;
     End;
End;


procedure TmxCustomCalendar.CurrMonthButtonClick(Sender: TObject);
begin
//
  FSelectionStart := EncodeDate(Fyear,FMonth,1);
  FSelectionEnd := EncodeDate(Fyear,FMonth,DaysInMonth(Fyear,FMonth));
//
  RepaintCalendar;
end;

function TmxCustomCalendar.FindWeek (WeekNumber : integer) : TmxPanel;
var X: Integer;
begin
  //
  result := nil;
  For X := 0 To 6 Do
  Begin
    if FPanel_Days [X,0].Fweek = WeekNumber Then BEGIN result := FPanel_Days[X,0] ; break; END;
  end;
end;

function TmxCustomCalendar.FindDay (UneDate : TDateTime) : TmxPanel;
var X,Y: Integer;
begin
  //
  result := nil;
  For X := 1 To 6 Do
  Begin
		For Y := 1 To 7 Do
    begin
    	if FPanel_Days [X,Y].FDate = UneDate Then BEGIN result := FPanel_Days[X,Y] ; break; END;
    end;
  end;
end;

End.

