unit EventDecla;
(*LM20061114

  Petit objet (à étoffer au niveau évènement...) pour mémoriser les évènements de fiche décla

  -Sur le OnArgument :
  .EventDecla.create : empile tous les évènement existant
  .EventDecla.Rebranche pour réaffecter les procedures de la fiche sur une méthode de la tof/tom

  -Sur chaque évènement "rebranché" :
  .EventDecla.Exec : execute de code de la fiche ou le script

  -Sur le OnClose :
  .EventDecla.free

  Pour ajouter un nouvel évènement :
  . implémenter le type : type Txxxxx = procedure ... of object ;
  . implémenter la procédure Rebranche correspondant au type ajouté
  . implémenter la procédure Exec correspondant au type ajouté


*)
interface
uses
  sysutils,
  Classes,
  HCtrls,
  hMsgBox,
  forms,
  controls,
  TypInfo
  ;

{$M+}

function getPropertyVal(obj:TObject ; NomPro : String; var valeur  : Variant) : boolean  ;  //Pour avoir la valeur d'une propriété publique ou publiée LM20070415
function isPropertyExists(obj:TObject ; NomPro : String) : boolean;  //Pour savoir si propriété publique ou publiée existe LM20070415

type TPopupEvent = procedure (x,y:integer) of object ;
type TCloseFrm = procedure (Sender:TObject; var Action:TCloseAction) of object ;//LM20070703
type TQueryCloseFrm = procedure (sender:TObject; var CanClose : boolean ) of object ;


type TItemEvent = class(TObject)
  public
    methode : TMethod;
    oldMethode : TPropInfo ;
  end;

type
  TEventDecla = class(TComponent)
  private
//    tp:tpersistent;
    frm:TForm ;
    fOldTWndMethod:TWndMethod;
    ScriptDecla : TStringList;
    procedure upCase(var st1:string; var st2:string) ;
    procedure AddForm ;
    procedure Add (NomCtrl : string; prop : PPropInfo);
    function  ExecCommun (nomCtrl, nomEvent : string; var m : TMethod ) : boolean;
    function  RebrancheCommun(nomCtrl, nomEvent : string;  newProc : TMethod):boolean;
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    //procedure Dispatch (var Message); override;

  public
    constructor Create (f:TForm); overload;
    destructor Destroy; override ;

    procedure Rebranche(nomCtrl, nomEvent : string;  newProc : TNotifyEvent) ;overload;
    procedure Rebranche(nomCtrl, nomEvent : string;  newProc : TKeyEvent); overload;
    procedure Rebranche(nomCtrl, nomEvent : string;  newProc : TKeyPressEvent); overload;
    procedure Rebranche(nomCtrl, nomEvent : string;  newProc : TPopupEvent); overload;
    procedure Rebranche(nomCtrl, nomEvent : string;  newProc : TQueryCloseFrm); overload;
    procedure Rebranche(nomCtrl, nomEvent : string;  newProc : TCloseFrm); overload;//LM20070703
    procedure Rebranche(nomCtrl, nomEvent : string;  newProc : TChangeRC); overload;
    procedure Rebranche(nomCtrl, nomEvent : string;  newProc : TChangeCell); overload;//LM20071003

    procedure  Exec (nomCtrl, nomEvent : string; sender:TObject); overload;  //TNotifyEvent
    procedure  Exec (nomCtrl, nomEvent : string; sender:TObject; var Key: Word; Shift: TShiftState); overload;  //TKeyEvent
    procedure  Exec (nomCtrl, nomEvent : string; sender:TObject; var Key: Char); overload; //TKeyPressEvent
    procedure  Exec (nomCtrl, nomEvent : string; x,y : integer); overload; //TPopupEvent
    procedure  Exec (nomCtrl, nomEvent : string; sender:TObject; var canClose : boolean); overload;
    procedure  Exec (nomCtrl, nomEvent : string; sender:TObject; var Action:TCloseAction); overload; //LM20070703
    procedure  Exec (nomCtrl, nomEvent : string; Sender: TObject; Ou:Longint; var Cancel: Boolean; Chg:boolean); overload; //TChangeRC
    procedure  Exec (nomCtrl, nomEvent : string; sender:TObject; var ACol, ARow: Longint; var Cancel:Boolean); overload; //LM20071003

  end;
{$M-}

implementation

procedure StringList2File (SL:TStringList; FileName:string);
begin
if (FileExists(FileName)) then DeleteFile(FileName);
SL.SaveToFile(FileName) ;
end;


constructor TEventDecla.Create (f:TForm);
begin
  inherited Create(f);
  ScriptDecla:=TStringList.create ;
  frm:=f;
  // TB plus utile si component
  if frm <> nil then
     frm.FreeNotification(Self);
  fOldTWndMethod:=f.WindowProc;
  AddForm ;
end;

destructor TEventDecla.Destroy;
var i:integer ;
begin
  for i:=ScriptDecla.Count-1 downto 0 do
    ScriptDecla.Objects[i].Free;
  FreeAndNil(ScriptDecla);
  inherited ;
end;

procedure TEventDecla.Notification(AComponent: TComponent; Operation: TOperation);
begin
  // TB 04/07/07
  if (Operation = opRemove) then
  begin
    if (AComponent = frm) then
      frm := nil
  end;
  inherited;
end;

function TEventDecla.ExecCommun (nomCtrl, nomEvent : string; var m : TMethod ) : boolean;
var n:integer;
begin
  result:=false ;
  Application.ProcessMessages;//LM20070213 test sur csFreeNotification?
  if frm = nil then exit; // tb
  if csDestroying in frm.ComponentState then exit ;//LM20070213
  upCase(nomCtrl, nomEvent) ;
  n := ScriptDecla.IndexOf(nomCtrl+ '|' + nomEvent);
  if n > -1 then
  begin
    m := TItemEvent(ScriptDecla.Objects[n]).methode ;
    result:= assigned(m.code)
  end;
end ;


procedure TEventDecla.Exec (nomCtrl, nomEvent : string; sender:TObject; var Key: Word; Shift: TShiftState);
var m:TMethod ;
begin
  if ExecCommun (nomCtrl, nomEvent, m) then
  begin
    //laisser au client key:=0;
    TKeyEvent(m)(sender, key, shift) ;
  end ;
end ;

procedure TEventDecla.Exec (nomCtrl, nomEvent : string; sender:TObject) ;
var m : TMethod ;
begin
  if ExecCommun (nomCtrl, nomEvent, m) then TNotifyEvent(m)(sender) ;
end ;

procedure TEventDecla.Exec (nomCtrl, nomEvent : string; sender:TObject; var Key: char);
var m:TMethod ;
begin
  if ExecCommun (nomCtrl, nomEvent, m) then
  begin
    //laisser au client key:=#0;
    TKeyPressEvent(m)(sender, key) ;
  end ;
end ;

procedure  TEventDecla.Exec (nomCtrl, nomEvent : string; x,y : integer);
var m:TMethod ;
begin
  if ExecCommun (nomCtrl, nomEvent, m) then
    TPopupEvent(m)(x, y) ;
end ;

procedure  TEventDecla.Exec (nomCtrl, nomEvent : string; Sender: TObject; Ou:Longint; var Cancel: Boolean; Chg:boolean);
var m:TMethod ;
begin
  if ExecCommun (nomCtrl, nomEvent, m) then
    TChangeRC(m)(Sender, Ou, Cancel, Chg) ;
end ;

procedure  TEventDecla.Exec (nomCtrl, nomEvent : string; sender:TObject; var canClose : boolean);
var m:TMethod ;
begin
  if ExecCommun (nomCtrl, nomEvent, m) then
    TQueryCloseFrm(m)(Sender, CanClose) ;
end ;

procedure TEventDecla.Exec (nomCtrl, nomEvent : string; sender:TObject; var Action:TCloseAction); //LM20070703
var m:TMethod ;
begin
  if ExecCommun (nomCtrl, nomEvent, m) then
    TCloseFrm(m)(Sender, Action) ;
end ;

procedure TEventDecla.Exec (nomCtrl, nomEvent : string; Sender: TObject; var ACol, ARow: Longint; var Cancel:Boolean);
var m:TMethod ;
begin
  if ExecCommun (nomCtrl, nomEvent, m) then
    TChangeCell(m)(Sender, ACol, ARow, Cancel) ;
end ;


procedure TEventDecla.Add (NomCtrl : string; prop : PPropInfo);
var o:TItemEvent ;
    m:TMethod ;
    nomEvent : string;
begin
  if prop<>nil then
  begin
    nomEvent := prop.Name ;
    upcase(NomCtrl, nomEvent);
    if NomCtrl=frm.name then m := GetMethodProp(frm, prop)
                        else m := GetMethodProp(TControl(frm.Findcomponent(NomCtrl)), prop);
    if (m.code<>nil) then
    begin
      o:=TItemEvent.create;
      o.methode:= m;
      o.oldMethode := TPropInfo(prop^) ;
      ScriptDecla.AddObject((NomCtrl + '|' + nomEvent), Tobject(o)) ;
    end ;
  end ;
end ;

procedure TEventDecla.AddForm ;
var n,c : integer ;
    LstProp : TPropList ;
    s:string ;
    ctrl : TControl ;

      procedure traceAddForm (nom:string);
      begin
        s:= s + nom + ' - ' + TPropInfo(LstProp[n]^).Name + '       ' ;
        if (n mod 8=0) then s:=s+'#10';
      end ;
begin
  //TPropInfo(LstProp[n]^).propType^.Name = TNotifyEvent...
  //forme

  for n := 0 to GetPropList(frm.ClassInfo, tkMethods, @LstProp) - 1 do
    if assigned(LstProp[n]) then
    begin
      Add(frm.Name, LstProp[n]) ;
      //traceAddForm (frm.Name);
    end ;

  //composants
  for c:= 0 to frm.ComponentCount-1 do
  begin
    ctrl:=TControl(frm.Components[c]) ;
    if ctrl.name='' then continue ;//LM20070308
    for n := 0 to GetPropList(ctrl.ClassInfo, tkMethods, @LstProp) - 1 do
      if assigned(LstProp[n]) then
      begin
        Add(Ctrl.Name, LstProp[n]) ;
        //traceAddForm (Ctrl.Name) ;
      end ;
  end ;
  if s<>'' then pgiInfo (s) ;
  //StringList2File(ScriptDecla, 'c:\eventDecla.txt');
end ;

function TEventDecla.RebrancheCommun(nomCtrl, nomEvent : string;  newProc : TMethod):boolean;
var n : integer ;
    pi : TPropInfo ;
    m : TMethod;
    LaForm : boolean ;
    ctrl : TControl;
begin
//  result:=false ;
  upCase(nomCtrl, nomEvent) ;
  n := ScriptDecla.IndexOf(nomCtrl+ '|' + nomEvent);
  LaForm := (NomCtrl=uppercase(frm.name));

  if (n > -1) then  //TObject(newProc.Data).MethodName(newProc.Code)
  begin
    pi := TItemEvent(ScriptDecla.Objects[n]).OldMethode ;
    if LaForm then m := GetMethodProp(frm, @pi)
              else m := GetMethodProp(TControl(frm.FindComponent(nomCtrl)), @pi);
    if (m.code<>nil) then
    begin
      if LaForm then SetMethodProp(frm, @pi, newProc)
                else SetMethodProp(TControl(frm.FindComponent(nomCtrl)), @pi, newProc);
    end ;
  end
  else //Non référencé => on branche direct
  begin
    if LaForm then
      SetMethodProp(frm, nomEvent, newProc)
    else
    begin
      ctrl := TControl(frm.FindComponent(nomCtrl));//+LMO20061129
      if ctrl<>nil then
        SetMethodProp(ctrl, nomEvent, newProc)
      (*else if v_pgi.sav then
        Pgiinfo('EventDecla.RebrancheCommun - ' + nomCtrl + ' introuvable.'
      *)
        ; //-LMO20061129
    end ;

  end ;
  result:=true ;//sert à rien dans le contexte mais je laisse
end ;


procedure TEventDecla.Rebranche(nomCtrl, nomEvent : string; newProc : TNotifyEvent ) ;
begin
  RebrancheCommun (nomCtrl, nomEvent, TMethod(newProc))
end ;

procedure TEventDecla.Rebranche(nomCtrl, nomEvent : string;  newProc : TKeyEvent) ;
begin
  RebrancheCommun (nomCtrl, nomEvent, TMethod(newProc))
end ;

procedure TEventDecla.Rebranche(nomCtrl, nomEvent : string;  newProc : TKeyPressEvent) ;
begin
  RebrancheCommun (nomCtrl, nomEvent, TMethod(newProc))
end ;

procedure TEventDecla.Rebranche(nomCtrl, nomEvent : string;  newProc : TPopupEvent);
begin
  RebrancheCommun (nomCtrl, nomEvent, TMethod(newProc))
end ;

procedure TEventDecla.Rebranche(nomCtrl, nomEvent : string;  newProc : TChangeRC);
begin
  RebrancheCommun (nomCtrl, nomEvent, TMethod(newProc))
end ;

procedure TEventDecla.Rebranche(nomCtrl, nomEvent : string;  newProc : TQueryCloseFrm);
begin
  RebrancheCommun (nomCtrl, nomEvent, TMethod(newProc))
end ;

procedure TEventDecla.Rebranche(nomCtrl, nomEvent : string;  newProc : TCloseFrm); //LM20070703
begin
  RebrancheCommun (nomCtrl, nomEvent, TMethod(newProc))
end ;

procedure TEventDecla.Rebranche(nomCtrl, nomEvent : string;  newProc : TChangeCell);//LM20071003
begin
  RebrancheCommun (nomCtrl, nomEvent, TMethod(newProc))
end ;


procedure TEventDecla.upCase(var st1:string; var st2:string) ;
begin
  st1 := upperCase(st1) ;
  st2 := upperCase(st2) ;
end ;

function isPropertyExists(obj:TObject ; NomPro : String) : boolean;  //LM20070415
Var LProp : PPropList;
    IProp : PPropInfo;
    Ind : Integer;
begin
  result:=false;
  if Assigned(obj) then
  begin
    NomPro := upperCase(NomPro) ;
    new(LProp) ;
    GetPropInfos(obj.ClassInfo,LProp);
    for Ind := 0 to GetTypeData(obj.ClassInfo)^.PropCount-1 do
    begin
      Try
        IProp:=LProp^[Ind];
        if upperCase(IProp^.Name) = NomPro then
        begin
          result:=true ;
          break ;
        end;
      Except
      end ;
    end;
    dispose(LProp);
  end ;
end;


function getPropertyVal(obj:TObject ; NomPro : String; var valeur  : Variant) : boolean  ;  ////LM20070415
Var LProp : PPropList;
    IProp : PPropInfo;
    Ind : Integer;
begin
  result:=false ;
  if Assigned(obj) then
  begin
    valeur:=#0 ;
    NomPro := upperCase(NomPro) ;
    new(LProp) ;
    GetPropInfos(obj.ClassInfo,LProp);
    for Ind := 0 to GetTypeData(obj.ClassInfo)^.PropCount-1 do
    begin
//      Try
        IProp:=LProp^[Ind];
        if upperCase(IProp^.Name) = NomPro then
        begin
          result:=true;
          case IProp^.PropType^.Kind of
            tkInteger :
              if UpperCase(LProp^[Ind]^.PropType^.Name) = 'TCOLOR' then
                valeur :=IntToHex(GetOrdProp(obj, LProp^[Ind]),6)
              else                valeur :=GetOrdProp(obj,IProp);            tkEnumeration : valeur:=GetOrdProp(obj,IProp);
            tkString,tkWString,tkLString : valeur:=GetStrProp(obj,IProp);
            tkFloat : valeur := FormatFloat('0.0000000',GetFloatProp(obj, LProp^[Ind]));
          end;
          break ;
        end;
//      Except
//      end ;
    end;
    dispose(LProp);
  end ;
end;


end.


