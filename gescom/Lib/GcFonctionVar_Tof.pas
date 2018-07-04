{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCFONCTIONVAR ()
Mots clefs ... : TOF;GCFONCTIONVAR
*****************************************************************}
Unit GcFonctionVar_Tof ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}                      
     MaineAGL,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,HTB97,ExtCtrls, Windows,
     UTOF,Graphics,UTOB,AGLInit,Vierge ;

function GCLanceFiche_FonctionVar(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

const LabelWidth = 100 ;
      ControlWidth  = 85 ;
      ControlHeight = 25 ;
      XMargin = 8 ;
      YMargin = 10 ;

Type
  TOF_GCFONCTIONVAR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArg : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    Action   : TActionFiche ;
    BCALCUL,BValider : TToolbarButton97;
    ListeVar : TStrings ;
    ValeurVar :TStrings ;
    ModifVar :TStrings ;
    VisibleVar :TStrings ;
    PosX,PosY,ColDer : Integer ;
    // -- Procédure liées aux champs
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DeterminePosition ;
    function  CreerLabel(Libelle : string ; CoordX,CoordY,ind : integer) : THLabel ;
    function  CreerNumEdit(Nom : string ; CoordX,CoordY,ind : integer) : TWinControl ;
    procedure BCALCULClick(Sender: TObject);
  end ;


Implementation
uses UTilFonctionCalcul;

function GCLanceFiche_FonctionVar(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{========================== Procédure de la TOF GCFONCTIONPARAM ===============================}
{==============================================================================================}
procedure TOF_GCFONCTIONVAR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCFONCTIONVAR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCFONCTIONVAR.OnUpdate ;
var ind,NumChp : integer ;
    Chps : TControl ;
    Resultat : string;
begin
  Inherited ;
NextPrevControl(Ecran);
for ind:=0 to TPanel(GetControl('PCHAMPS')).ControlCount-1 do
   begin
   Chps := TPanel(GetControl('PCHAMPS')).Controls[ind] ;
   if Chps is TLabel then continue ;
   NumChp := TWinControl(Chps).Tag ;
   ValeurVar[NumChp]:= FloatToStr(THNumEdit(Chps).Value) ;
   end;
Resultat:='';
for ind:=0 to ValeurVar.Count-1 do
   begin
   Resultat:=Resultat+ValeurVar[ind]+';';
   end;
TFVierge(Ecran).Retour:=Resultat ;
end ;

procedure TOF_GCFONCTIONVAR.OnLoad ;
Var StLibelle : string ;
    BModif : Boolean;
    ValVar : Double;
    ind : integer ;
    ChpLabel : THLabel ;
    ChpControl,FirstControl : TWinControl ;
begin
  Inherited ;
PosY:=0; PosX:=0; FirstControl:=nil ;
for ind:=0 to ListeVar.Count-1 do
   begin
   if VisibleVar[ind] = '-' then Continue;
   StLibelle:=ListeVar[ind] ;
   BModif:=(ModifVar[ind]='X') ;
   ValVar:=Valeur(ValeurVar[ind]) ;
   DeterminePosition;
   ChpLabel:=CreerLabel(StLibelle,PosX,PosY,ind);
   ChpControl:=CreerNumEdit(StLibelle,PosX+LabelWidth,PosY,ind) ;
   if (FirstControl=Nil) and (BModif) then FirstControl:=ChpControl ;
   if (ChpControl<>nil) and (ChpLabel<>nil) then ChpLabel.FocusControl:=ChpControl ;
   if (ChpControl<>nil) then
     begin
     if Action=taConsult then
       ChpControl.Enabled:=False 
     else
       ChpControl.Enabled:=BModif ;
     if Not BModif then THNumEdit(ChpControl).Color:=ClBtnFace ;
     THNumEdit(ChpControl).Value:=ValVar ;
     end;
   end ;
SetControlVisible('PCHAMPS',True);
SetControlProperty('PCHAMPS','AutoSize',True);
if Action=taConsult then Exit;
if FirstControl<>nil then
  begin
  Ecran.ActiveControl:=FirstControl ;
  FirstControl.SetFocus;
  end;
end ;

procedure TOF_GCFONCTIONVAR.OnArgument (StArg : String ) ;
Var SCB : TScrollBox;
    Ph : TPanel;
    St,StList,StVar,StModif,StVisible : string;
    ind : integer ;
begin
  Inherited ;
ind:=Pos('ACTION=',StArg) ;
if ind>0 then
  begin
  System.Delete(StArg,1,ind+6) ;
  St:=uppercase(ReadTokenSt(StArg)) ;
  if St='CREATION' then begin Action:=taCreat ; end ;
  if St='MODIFICATION' then begin Action:=taModif ; end ;
  if St='CONSULTATION' then begin Action:=taConsult ; end ;
  end ;

ListeVar:=TStringList.Create;
St:=StArg;
if st<>'' then begin Ecran.Caption:=St;UpdateCaption(Ecran); end;
if LaTOB<>Nil then
  begin
  SetControlText('CHPFORMULE',LaTOB.GetValue('AFFICHAGE'));
  StList:=LaTOB.GetValue('LISTEVAR');
  StVar:=LaTOB.GetValue('VALEURVAR');
  StModif:=LaTOB.GetValue('MODIFVAR');
  StVisible:=LaTOB.GetValue('VISIBLEVAR');
  end;
While StList<>'' do ListeVar.Add(ReadTokenSt(StList));
ValeurVar:=TStringList.Create ;
While StVar<>'' do ValeurVar.Add(ReadTokenSt(StVar));
ModifVar:=TStringList.Create ;
While StModif<>'' do ModifVar.Add(ReadTokenSt(StModif));
VisibleVar:=TStringList.Create ;
While StVisible<>'' do VisibleVar.Add(ReadTokenSt(StVisible));
BCALCUL:=TToolbarButton97(GetControl('BCALCUL'));
BCALCUL.OnClick:=BCALCULClick;
BValider := TToolbarButton97(GetControl('BVALIDER'));
// Création de la ScrollBox associée
SCB:=TScrollBox.create(Ecran) ;
SCB.Parent:=TTabSheet(getcontrol('ONGLETFONCTION')) ;
SCB.Name:='SCB';
SCB.ParentFont:=False; SCB.Font.Color:=clWindowText;
SCB.Align:=alClient;
Ph:=TPanel(GetControl('PCHAMPS'));
Ph.Parent:=SCB;
SetControlVisible('PCHAMPS',False);
TFVierge(Ecran).OnKeyDown := FormKeyDown;
end ;

procedure TOF_GCFONCTIONVAR.OnClose ;
begin
  Inherited ;
ListeVar.Free;
ValeurVar.Free;
ModifVar.Free;
end ;

{==============================================================================================}
{=============================== Procédure liées aux champs ===================================}
{==============================================================================================}
procedure TOF_GCFONCTIONVAR.DeterminePosition ;
begin
if (PosX=0) and (PosY=0) then
   begin
   PosX:=XMargin ; PosY:=YMargin ; ColDer:=2;
   end else
   begin
   if ColDer=1 then
      begin
      PosX:=XMargin ;
      PosY:=PosY+ControlHeight+YMargin ; //$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      inc(ColDer);
      end else if ColDer=2 then
      begin
      PosX:=PosX+(LabelWidth+ControlWidth+XMargin) ;
      ColDer:=1;
      end ;
   end ;
end ;

function TOF_GCFONCTIONVAR.CreerLabel(Libelle : string ; CoordX,CoordY,ind : integer) : THLabel ;
var Label2k10 : THLabel ;
begin
Label2k10:=THLabel.Create(Ecran) ;
Label2k10.Name:='T'+Libelle+IntToStr(ind) ;
Label2k10.Parent:=TPanel(GetControl('PCHAMPS')) ;
Label2k10.Autosize:=False ;
Label2k10.Top:=CoordY + 4 ;
Label2k10.Left:=CoordX ;
Label2k10.Height:=13 ;
Label2k10.Width:=LabelWidth-XMargin ;
Label2k10.Caption:=StringReplace(Libelle,'_',' ',[rfReplaceAll]);
Label2k10.Enabled:=True ;
Label2k10.Visible:=True ;
Label2k10.Transparent:=True ;
Label2k10.Alignment:=taLeftJustify ;
Label2k10.Color:=clBtnFace ;
result:=Label2k10 ;
end ;

function TOF_GCFONCTIONVAR.CreerNumEdit(Nom : string ; CoordX,CoordY,ind : integer) : TWinControl ;
var Chps : THNumEdit ;
begin
Chps:=THNumEdit.Create(Ecran) ;
Chps.Name :=Nom+IntToStr(ind) ;
Chps.Parent :=TPanel(GetControl('PCHAMPS')) ;
Chps.Top:=CoordY ;
Chps.Left:=CoordX ;
Chps.Height:=21 ;
Chps.Width:=ControlWidth ;
Chps.Tag:=ind ;
Chps.TabOrder:=ind ;
Chps.Masks.PositiveMask:='##0.000';

Chps.Value:=0 ;
result:=TWinControl(Chps) ;
end ;

procedure TOF_GCFONCTIONVAR.BCALCULClick(Sender: TObject);
Var FoncQte : R_FonctCal ;
    ind,ind2 : integer;
    StNom, Formule : string;
begin
NextPrevControl(Ecran);
Formule:=LaTOB.GetValue('FORMULE');
FoncQte:=RecupVariableFormule(Formule);
FoncQte.Formule:=Formule;
for ind:=0 to ListeVar.Count-1 do
  begin
  for ind2:=0 to High(FoncQte.VarLibelle)-1 do
    begin
    if FoncQte.VarLibelle[ind2]<>ListeVar[ind] then continue;
    if VisibleVar[ind]='X' then
      begin
      StNom:=ListeVar[ind]+IntToStr(ind) ;
      FoncQte.VarValeur[ind2]:=StrToFloat(GetControlText(StNom));
      end
      else
      FoncQte.VarValeur[ind2]:=Valeur(ValeurVar[ind]); 
    end;
  end;
FoncQte:=EvaluationFormule('',FoncQte,True);
SetControlText('RESULTAT',FloatToStr(FoncQte.Resultat));
end ;

{==============================================================================================}
{================================== Touches de contrôle =======================================}
{==============================================================================================}
procedure TOF_GCFONCTIONVAR.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
   VK_F10     : begin
                Key:=0 ;
                if Action<>taConsult then
                  begin
                  bValider.OnClick(nil)  ;
                  TForm(Ecran).ModalResult:=2 ;
                  end;
                end ;
   end;
end;

Initialization
  registerclasses([TOF_GCFONCTIONVAR]) ;
end.
