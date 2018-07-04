{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PLANAUXI ()
Mots clefs ... : TOF;PLANAUXI
*****************************************************************}
Unit PLANAUXI_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
  eMul, uTob, MaineAGL, eQRS1,
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  mul, FE_Main, QRS1,
{$ENDIF}
  ParamSoc,		//GetParamSocSecur YMO
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF;

procedure CPLanceFiche_PLANAUXI(const Compte : String);

Type
  TOF_PLANAUXI = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure AuxiElipsisClick         (Sender: TObject);
  private
    FTri, TYPEEDITION : THRadioGroup;
    procedure TriOnClick(Sender : TObject);
    procedure TypeOnClick(Sender : TObject);

    procedure SetItemIndex(Combo : String);
  end ;

Implementation

uses UTofMulParamGen;

procedure CPLanceFiche_PLANAUXI(const Compte : String);
begin
  AGLLanceFiche('CP','CPPLANAUXI','','',Compte);
end;

procedure TOF_PLANAUXI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PLANAUXI.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PLANAUXI.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PLANAUXI.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PLANAUXI.OnArgument(S : String);
var
  i : Integer;
begin
  Inherited;

  SetControlText('T_AUXILIAIRE', S);
  SetControlText('T_AUXILIAIRE_', S);

  SetItemIndex('T_NATUREAUXI');
  SetItemIndex('T_DEVISE');
  SetItemIndex('T_PAYS');
  SetItemIndex('T_REGIMETVA');

  FTri := THRadioGroup(GetControl('FTRI', True));
  FTri.OnClick := TriOnClick;

  TYPEEDITION := THRadioGroup(GetControl('TYPEEDITION', True));
  TYPEEDITION.OnClick := TypeOnClick;
  TypeOnClick(nil);

  // Libellé des tables libres (Evite des requêtes à chaque fiche)
  for i := 0 to 9 do
    SetControlText('T'+IntToStr(i), RechDom('TTTABLESLIBRESAUX', 'T0'+IntToStr(i), False));

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('T_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('T_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;
end ;

procedure TOF_PLANAUXI.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PLANAUXI.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PLANAUXI.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_PLANAUXI.AuxiElipsisClick(Sender: TObject);
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


// Choix du tri de l'édition : Par compte auxiliaire ou Libellé
procedure TOF_PLANAUXI.TriOnClick(Sender: TObject);
begin
  if (FTri.ItemIndex = 0) then SetControlText('XX_ORDERBY' , 'T_AUXILIAIRE')
                          else SetControlText('XX_ORDERBY' , 'T_LIBELLE')
end;

procedure TOF_PLANAUXI.SetItemIndex(Combo: String);
begin
  TComboBox(GetControl(Combo, True)).ItemIndex := 0;
end;

// Choix du type d'édition : Liste ou Fiche
procedure TOF_PLANAUXI.TypeOnClick(Sender: TObject);
begin
  if (TYPEEDITION.ItemIndex = 0) then TFQRS1(Ecran).CodeEtat := 'AU2'
                                 else TFQRS1(Ecran).CodeEtat := 'AUX';
  {JP 02/11/05 : On n'autorise l'export qu'en mode liste}
  if (TYPEEDITION.ItemIndex > 0) then SetControlChecked('FLISTE', False);
  SetControlEnabled('FLISTE', (TYPEEDITION.ItemIndex = 0));
end;

Initialization
  registerclasses ( [ TOF_PLANAUXI ] ) ;
end.
