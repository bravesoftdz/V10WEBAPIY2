{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULZLNIVEAU ()
Mots clefs ... : TOF;PGMULZLNIVEAU
*****************************************************************
PT1   25/06/2008 FC V850 FQ 15067 Pb positionnement ligne en cwas du coup on avait l'impression que
                         l'élément dynamique était saisi sur tous les établissements alors qu'il ne
                         l'était que sur le 1er établissement
}
unit UTofPGMulZLNiveau;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db, HDB, Fe_Main,
{$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
  mul,
{$ELSE}
  eMul,
  uTob,
  MainEAGL,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  hqry,
  Pgpopuloutils,
  UTOF;

type
  TOF_PGMULZLNIVEAU = class(TOF)
    procedure OnArgument(S: string); override;
    procedure OnLoad; override;
  private
    TypeSaisie, pred: string;

{$IFNDEF EAGLCLIENT}
    Grille: THDBGrid;
{$ELSE}
    Grille: THGrid;
{$ENDIF}


    procedure GrilleDblClick(Sender: TObject);
    procedure OnClickpopactive(Sender: TObject);
  end;

implementation

procedure TOF_PGMULZLNIVEAU.Onload;
begin
  inherited;
  if TCheckBox(GetControl('POPACTIVE')) <> nil then
    if (GetControlText('POPACTIVE') = 'X') then setcontroltext('XX_WHERE', 'PPO_TYPEPOP LIKE "%PAI%" and PPC_PREDEFINI = "' + pred + '"')
    else setcontroltext('XX_WHERE', 'PPO_TYPEPOP LIKE "%PAI%"');
end;

procedure TOF_PGMULZLNIVEAU.OnArgument(S: string);
var
  popactive: TCheckBox;
  st: string;
begin
  inherited;
  TypeSaisie := ReadTokenPipe(S, ';');
  If TypeSaisie = 'ETB' then TFMul(Ecran).Caption := 'Eléments dynamiques établissement'
  else TFMul(Ecran).Caption := 'Eléments dynamiques population';
  Updatecaption(TFMul(Ecran));
{$IFNDEF EAGLCLIENT}
  Grille := THDBGrid(GetControl('Fliste'));
{$ELSE}
  Grille := THGrid(GetControl('Fliste'));
{$ENDIF}
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;

  // initialise le plus de la tablette PGPOPULATIONSAL  et la liste
  pred := GetPredefiniPopulation('PAI');

  popactive := TCheckBox(GetControl('POPACTIVE'));
  if popactive <> nil then
  begin
    if popactive.checked = true then st := ' and PPC_PREDEFINI = "' + pred + '"' else st := '';
    setcontrolproperty('PPC_POPULATION', 'plus', st);
//  popactive:=TCheckBox(GetControl('POPACTIVE'));
    popactive.OnClick := OnClickpopactive;
  end;
end;

procedure TOF_PGMULZLNIVEAU.GrilleDblClick(Sender: TObject);
var Q_Mul: THQuery;
  St: string;
begin
{$IFDEF EAGLCLIENT}
//  THQuery(Ecran).Q.TQ.Seek(THQuery(Ecran).FListe.Row - 1); //PT1 Nécessaire pour se positionner sur la bonne ligne en cwas
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1); //PT1 Nécessaire pour se positionner sur la bonne ligne en cwas
{$ENDIF}
  St := '';
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if TypeSaisie = 'ETB' then St := Q_MUL.FindField('ET_ETABLISSEMENT').AsString
  else St := Q_MUL.FindField('PPC_CODEPOP').AsString + ';' + Q_MUL.FindField('PPC_POPULATION').AsString;
  if St <> '' then AGLLanceFiche('PAY', 'SAISIEZLNIVEAU', '', '', TypeSaisie + ';' + St);
end;

procedure TOF_PGMULZLNIVEAU.Onclickpopactive(sender: Tobject);
var
  popactive: Tcheckbox;
  st: string;
begin
  popactive := TCheckBox(GetControl('POPACTIVE'));
  if popactive.checked = true then st := ' and PPC_PREDEFINI = "' + pred + '"' else st := '';
  setcontrolproperty('PPC_POPULATION', 'plus', st);
  TFMul(Ecran).BCherche.Click;
end;

initialization
  registerclasses([TOF_PGMULZLNIVEAU]);
end.

