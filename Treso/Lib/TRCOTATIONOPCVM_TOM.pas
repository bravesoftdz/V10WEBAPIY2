{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 6.0X.xxx.xxx    03/11/04    JP   Création de l'unité : Tom de saisie des cours des OPCVM
--------------------------------------------------------------------------------------}
unit TRCOTATIONOPCVM_TOM;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche,
  {$ELSE}
  FE_Main, Fiche,
  {$ENDIF}
  Forms, Classes, Hctrls, HTB97, UTOM, SysUtils, Controls;

type
  TOM_TRCOTATIONOPCVM = class(TOM)
    procedure OnNewRecord           ; override;
    procedure OnAfterUpdateRecord   ; override;
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnUpdateRecord        ; override;
  private
    CodeParam 	: string;
    DateParam	: TDateTime;
    PeutFermer  : Boolean;
    RetourOk    : Boolean;
    FCanClose   : TCloseQueryEvent;
    FFermeClick : TNotifyEvent;
    FValidClick : TNotifyEvent;
    AppelDirect : Boolean;

    procedure BFermeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ECanClose      (Sender : TObject; var CanClose : Boolean);
    procedure DeviseChange   (Sender : TObject);
    procedure OnAfterShow;
    procedure BFermeClick    (Sender : TObject);
    procedure BValiderClick  (Sender : TObject);
  end;


function TRLanceFiche_TRCoursOPCVM(Dom, Fiche, Range, Lequel, Arguments : string) : string;


implementation

uses
  ExtCtrls, Commun, HEnt1;

var
  FromOpcvm : Boolean;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRCoursOPCVM(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
var
  S : string;
begin
  S := Arguments;
  {C'est que l'on vient de TROPCVM_TOM où l'action n'est pas précisée. Si l'on vient du mul,
   théoriquement Lequel Existe dans la base !!!}
  if ReadTokenSt(S) = '' then begin
  {$IFNDEF EAGLCLIENT}
    {Afin d'éviter le message "Enregistrement inaccessible" : ce n'est pas nécessaire en eAgl
     car il se met automatiquement en création}
    S := Lequel;
    if not ExisteSQL('SELECT TTO_COTATION FROM TRCOTATIONOPCVM WHERE TTO_CODEOPCVM = "' +
                     ReadTokenSt(S) + '" AND TTO_DATE = "' + UsDateTime(StrToDateTime(ReadTokenSt(S))) +  '"') then begin
      Arguments := 'ACTION=CREATION' + Arguments;
    end;
  {$ENDIF}
    FromOpcvm := True;
  end;
  Result := AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{REMARQUE : TTO_DEVISE figure dans la table pour faciliter les jointures sur la table de
            chancellerie notamment pour la fiche de suivi
{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.OnNewRecord ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetField('TTO_CODEOPCVM', CodeParam);
  SetField('TTO_DATE', DateParam);
  SetField('TTO_DEVISE', GetControlText('DEV'));
  PeutFermer := False;
  {Si simili create en série, on n'a que le cours à saisir, donc on lui donne le focus}
  if CodeParam <> '' then
    SetFocusControl('TTO_COTATION');
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetControlEnabled('TTO_CODEOPCVM', True);
//  CodeParam  := GetField('TTO_CODEOPCVM');
  //DateParam  := GetField('TTO_DATE') + 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  CodeParam  := GetField('TTO_CODEOPCVM');
  DateParam  := GetField('TTO_DATE') + 1;
  if CodeParam = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez sélectionner un OPCVM');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if RetourOk then
    {le GetField n'est pas forcément à jour, car les speedbutons ne prennent pas le focus}
    TFFiche(Ecran).Retour := GetControlText('TTO_COTATION')
  else
    TFFiche(Ecran).Retour := '';
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
var
  Str : string;
begin
  inherited;
  SetControlText('DEV', '');
  Ecran.HelpContext := 150;
  TToolBarButton97(GetControl('BFERME')).OnMouseDown := BFermeMouseDown;
  {Action}
  ReadtokenSt(S);
  {Fiche d'appel : appel direct si on ne passe pas par le mul}
  AppelDirect := Trim(ReadtokenSt(S)) = 'DIRECT';
  CodeParam := Trim(ReadtokenSt(S));

  Str := Trim(ReadtokenSt(S));
  if Str = '' then DateParam := Date
              else DateParam := StrToDate(Str);

  {Pour simuler une création en série : on ne peut fermer la fiche par le bouton Valider que
   si on n'est pas passer par le mode création}
  FCanClose := TFFiche(Ecran).OnCloseQuery;
  TFFiche(Ecran).OnCloseQuery := ECanClose;
  PeutFermer := True;
  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('TTO_CODEOPCVM')).OnChange := DeviseChange;
  {$ELSE}
  THDBValComboBox(GetControl('TTO_CODEOPCVM')).OnChange := DeviseChange;
  {$ENDIF}
  TFFiche(Ecran).OnAfterFormShow := OnAfterShow;
  FFermeClick := TFFiche(Ecran).BFerme.OnClick;
  TFFiche(Ecran).BFerme.OnClick := BFermeClick;
  FValidClick := TFFiche(Ecran).BValider.OnClick;
  TFFiche(Ecran).BValider.OnClick := BValiderClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.ECanClose(Sender : TObject; var CanClose : Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := PeutFermer or AppelDirect;
  if CanClose then FCanClose(Sender, Canclose);
  PeutFermer := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.BFermeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{---------------------------------------------------------------------------------------}
begin
  PeutFermer := True;
  RetourOk := False;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.DeviseChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Devise : string;
begin
  Devise := GetControlText('TTO_CODEOPCVM');
  if Devise <> '' then begin
    Devise := RetDeviseOPCVM(Devise);
    {Mise à jour des contrôles}
    AssignDrapeau(TImage(GetControl('IDEV')), Devise);
    SetControlCaption('DEV', Devise);
    {Éventuelle mise à jour du Champ}
    if VarToStr(GetField('TTO_DEVISE')) = '' then
      SetField('TTO_DEVISE', Devise);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.OnAfterShow;
{---------------------------------------------------------------------------------------}
begin
  {$IFNDEF EAGLCLIENT}
  {Au chargement le OnChange du combo est exécuté en eAgl mais pas en 2/3, d'où la ligne ci-dessous}
  DeviseChange(GetControl('TTO_CODEOPCVM'));
  {$ENDIF}
  if ((VarToStr(GetField('TTO_CODEOPCVM')) = '') or
      (Length(VarToStr(GetField('TTO_CODEOPCVM'))) > 3)) and (CodeParam <> '') then
    SetField('TTO_CODEOPCVM', CodeParam);
  if CodeParam <> '' then
    SetFocusControl('TTO_COTATION');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.BFermeClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  FFermeClick(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRCOTATIONOPCVM.BValiderClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  RetourOk := True;
  FValidClick(Sender);
  {$IFNDEF EAGLCLIENT}
  Ecran.Close;
  {$ENDIF}
end;

initialization
  RegisterClasses([TOM_TRCOTATIONOPCVM]);

end.

