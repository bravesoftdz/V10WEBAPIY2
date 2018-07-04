unit wProgress;

interface

uses
  Controls
  ,Forms
  ;

type
  ///<summary>
  ///Interface fournissant les méthodes disponibles pour <b>surcharger</b> la fenêtre de progression de CBP
  ///</summary>
  IWProgress = interface
  ['{1B032731-FEA4-49CF-BA0D-B0B94F018E74}']
    ///<summary>
    ///Ouvre la fenêtre de progression de CBP
    ///</summary>
    ///<param name="Sender">Form appelant la fenêtre. Si nil, la Form active est prise en compte</param>
    ///<param name="Caption">Titre de la fenêtre de progression</param>
    ///<param name="Title">Titre de la barre de progression</param>
    ///<param name="Count">Nombre d'itérations</param>
    ///<param name="WithCancelButton">Permet d'interrompre le process itéré</param>
    ///<value>
    ///Ouvre la fenêtre de progression de CBP mais s'assure qu'elles ne se 
    ///cumules pas. En cas d'ouvertures imbriquées, Open ne fait rien.
    ///</value>
    procedure Open(Sender: TForm; const Caption, Title: WideString;
      const Count: Integer; WithCancelButton: Boolean);

    ///<summary>
    ///Permet de faire avancer la barre de progression
    ///</summary>
    ///<param name="Title">Titre de la barre de progression</param>
    ///<returns>"False" si le bouton "Cancel" est activé et que l'on a cliqué dessus pour interrompre le process</returns>
    ///<value>
    ///Permet de faire avancer la barre de progression.
    ///En cas d'appels imbriqués de la fenêtres, Next ne fait pas avancer la 
    ///barre de progression mais impacte le titre.
    ///</value>
    function Next(const Title: WideString = ''): Boolean;

    ///<summary>
    ///Fermer la fenêtre de progression
    ///</summary>
    ///<value>
    ///Ferme la fenêtre de progression.
    ///En cas de d'appels imbriqués, Close ne fait rien.
    ///</value>
    procedure Close;

    function GetCount: Integer;
    procedure SetCount(const Value: Integer);
    ///<summary>
    ///Permet de redéfinir à la volée le nombre d'itérations maximum
    ///</summary>
    ///<value>
    ///Permet de redéfinir/consulter à la volée le nombre d'itérations maximum
    ///En cas de d'appels imbriqués, SetCount ne modifie rien
    ///</value>
    property Count: Integer read GetCount write SetCount;
  end;

function wProgressForm: IWProgress;

implementation

uses
  Classes
  ,cbpSingleton
  {$if not (Defined(EAGLSERVER) or Defined(ERADIO))}
    ,Windows
    ,Math
    ,StdCtrls
    ,ExtCtrls
    ,ed_Tools
    ,HEnt1
    ,TradMini
  {$else}
    ,cbpTrace
    ,SysUtils
  {$ifend not(EAGLSERVER or ERADIO)}
  ;

type
  TWProgressSingleton = class(TCbpInterfacedSingleton)
  protected
    function CreateSingletonInstance: IInterface; override;
  public
    function GetProgress: IWProgress;
  end;

  {$if not (Defined(EAGLSERVER) or Defined(ERADIO))}
    TWProgressForm = class(TInterfacedObject, IWProgress)
    private
      FCurrentForm: TForm;
      FCaption, FTitle: WideString;
      FProgressCount, FCount: Integer;
      FWithCancelButton: Boolean;

      procedure SetControlsEnabled(C: TControl; Enable: Boolean);
      procedure DisableControls;
      procedure EnableControls;
      procedure DisableForm;
      procedure EnableForm;
      function GetCount: Integer;
      procedure SetCount(const Value: Integer);
    public
      procedure Open(Sender: TForm; const Caption, Title: WideString;
        const Count: Integer; WithCancelButton: Boolean);
      function Next(const Title: WideString = ''): Boolean;
      procedure Close;

      property Count: Integer read GetCount write SetCount;
    end;
  {$else}
    TWProgressForm = class(TInterfacedObject, IWProgress)
    private
      procedure Trace(const Name, Msg: String);
      function GetCount: Integer;
      procedure SetCount(const Value: Integer);
    public
      procedure Open(Sender: TForm; const Caption, Title: WideString;
        const Count: Integer; WithCancelButton: Boolean);
      function Next(const Title: WideString = ''): Boolean;
      procedure Close;
      property Count: Integer read GetCount write SetCount;
    end;
  {$ifend not(EAGLSERVER or ERADIO)}

var
  _wProgress: TWProgressSingleton;

function wProgressForm: IWProgress;
begin
  Result := _wProgress.GetProgress()
end;

{ TWProgressSingleton }

function TWProgressSingleton.GetProgress: IWProgress;
begin
  Result := Self.GetSingleton as IWProgress
end;

function TWProgressSingleton.CreateSingletonInstance: IInterface;
begin
  Result := TWProgressForm.Create()
end;

{ TWProgress }

{$if not (Defined(EAGLSERVER) or Defined(ERADIO))}

procedure TWProgressForm.Open(Sender: TForm; const Caption,
  Title: WideString; const Count: Integer; WithCancelButton: Boolean);
begin
  if FProgressCount = 0 then
  begin
    FCount := Count;
    FCaption := Caption;
    FTitle := Title;
    FWithCancelButton := WithCancelButton;
    FCurrentForm := Sender;

    if not V_Pgi.SilentMode then
    begin
      if not Assigned(FCurrentForm) then
        FCurrentForm := Screen.ActiveForm;
      if not Assigned(FCurrentForm) or (Assigned(FCurrentForm.Parent) and FCurrentForm.Parent.InheritsFrom(TPanel)) then
        FCurrentForm := Application.MainForm;
      if Assigned(FCurrentForm) then
        DisableForm();
      try
        InitMoveProgressForm(nil, FCaption, FTitle, FCount, FWithCancelButton, True)
      except
        if Assigned(FCurrentForm) then
          EnableForm();
        raise
      end
    end;
  end;

  Inc(FProgressCount);
end;

function TWProgressForm.Next(const Title: WideString): Boolean;
begin
  if not V_Pgi.SilentMode then
  begin
    if FProgressCount = 1 then
      Result := MoveCurProgressForm(Title)
    else
      Result := TextProgressForm(Title)
  end
  else
    Result := True
end;

procedure TWProgressForm.Close;
begin
  FProgressCount := Max(0, FProgressCount - 1);

  if not V_Pgi.SilentMode and (FProgressCount = 0) then
  begin
    if Assigned(FCurrentForm) then
      EnableForm();
    FiniMoveProgressForm();
  end;
end;

function TWProgressForm.GetCount: Integer;
begin
  Result := FCount;
end;

procedure TWProgressForm.SetCount(const Value: Integer);
begin
  if MaxValueProgressForm(Value) then
    FCount := Value;
end;

procedure TWProgressForm.SetControlsEnabled(C: TControl; Enable: Boolean);

  procedure _SetControlEnable(C: TControl);
  var
    C_OnExit: TNotifyEvent;

    procedure __Disable;
    begin
      C.Enabled := False;
      C.Hint := '@' + C.Hint;
    end;

  begin
    if C.Visible and (C.Enabled <> Enable) and not(C is TImage) and (C <> Screen.ActiveForm) then
    begin
      if not Enable then
      begin
        if (C is TWinControl) and Assigned(TEdit(C).OnExit) then
        begin
          C_OnExit := TEdit(C).OnExit;
          TEdit(C).OnExit := nil;
          try
            __Disable()
          finally
            TEdit(C).OnExit := C_OnExit
          end
        end
        else
          __Disable()
      end
      else if Copy(C.Hint, 1, 1) = '@' then
      begin
        hSetHint(C, Copy(hGetHint(C), 2, Length(hGetHint(C)) - 1));
        C.Enabled := True;
      end
    end
  end;

var
  i: Integer;
  WC: TWinControl;
begin
  if C is TWinControl then
  begin
    WC := TWinControl(C);
    for i := 0 to Pred(WC.ControlCount) do
      if WC.Controls[i] is TWinControl then
        SetControlsEnabled(WC.Controls[i], Enable)
  end;
  _SetControlEnable(C)
end;

procedure TWProgressForm.DisableControls;
var
  i: Integer;
begin
  for i := 0 to Pred(FCurrentForm.ControlCount) do
    SetControlsEnabled(FCurrentForm.Controls[i], False);
end;

procedure TWProgressForm.DisableForm;
begin
  if Assigned(FCurrentForm) then
  begin
    FCurrentForm.ControlStyle := FCurrentForm.ControlStyle + [csNoStdEvents];
    DisableControls();
  end;
  EnableMenuItem(GetSystemMenu(FCurrentForm.Handle, False), SC_CLOSE, MF_GRAYED);
end;

procedure TWProgressForm.EnableControls;
var
  i: Integer;
begin
  for i := 0 to Pred(FCurrentForm.ControlCount) do
    SetControlsEnabled(FCurrentForm.Controls[i], True);
end;

procedure TWProgressForm.EnableForm;
begin
  if Assigned(FCurrentForm) then
  begin
    FCurrentForm.ControlStyle := FCurrentForm.ControlStyle - [csNoStdEvents];
    EnableControls();
    EnableMenuItem(GetSystemMenu(FCurrentForm.Handle, False), SC_CLOSE, MF_ENABLED);
  end
end;

{$else} //dummy eAglServer

procedure TWProgressForm.Trace(const Name, Msg: String);
begin
  if cbpTrace.Trace.IsTraceVerboseEnabled() then
    cbpTrace.Trace.TraceVerbose(ClassName, Name + '(' + Msg + ')')
end;

procedure TWProgressForm.Open(Sender: TForm; const Caption,
  Title: WideString; const Count: Integer; WithCancelButton: Boolean);
begin
  Trace('Open', Caption + ' : ' + Title);
end;

function TWProgressForm.Next(const Title: WideString): Boolean;
begin
  Trace('Next', Title);
  Result := True;
end;

procedure TWProgressForm.Close;
begin
  Trace('Close', '')
end;

function TWProgressForm.GetCount: Integer;
begin
  Trace('GetCount', '');
  Result := 0;
end;

procedure TWProgressForm.SetCount(const Value: Integer);
begin
  Trace('SetCount', IntToStr(Value))
end;

{$ifend not(EAGLSERVER or ERADIO)}

initialization
  _wProgress := TWProgressSingleton.Create();
finalization
  if Assigned(_wProgress) then
    _wProgress.Free;
end.

