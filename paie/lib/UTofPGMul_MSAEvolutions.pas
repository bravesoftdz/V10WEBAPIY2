{***********UNITE*************************************************
Auteur  ...... :  JL
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMSAMULEVOLUTIONS ()
Mots clefs ... : TOF;PGMSAMULEVOLUTIONS
*****************************************************************
PT1 19/05/2005 V_60 JL FQ 12296 Modif des caption ecran
PT2 15/11/2006 V_70 JL FQ 13605 Ajout message si pas d'enregistrement sur double click
PT3 01/08/2007 V_80 JL FQ 14479 Ajout période
}
unit UTofPGMul_MSAEvolutions;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} mul, FE_Main, DBGrids, HDB,
  {$ELSE}
  eMul, uTob, MainEAgl,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, HTB97, HQry, EntPaie, PGOutils,PGOutils2;

type
  TOF_PGMSAMULEVOLUTIONS = class(TOF)
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
  private
    {$IFNDEF EAGLCLIENT}
    Grille: THDBGrid;
    {$ELSE}
    Grille: THGrid;
    {$ENDIF}
    procedure CreerEvolution(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
  end;

implementation

procedure TOF_PGMSAMULEVOLUTIONS.OnLoad;
begin
  inherited;
end;

procedure TOF_PGMSAMULEVOLUTIONS.OnArgument(S: string);
var
  BNew: TToolBarButton97;
  Edit: THEdit;
  DateDebut, DateFin: TDateTime;
  aa, mm, jj, Mois: word;
begin
  inherited;
  {$IFNDEF EAGLCLIENT}
  Grille := THDBGrid(GetControl('Fliste'));
  {$ELSE}
  Grille := THGrid(GetControl('Fliste'));
  {$ENDIF}
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;
  BNew := TToolbarButton97(getcontrol('BInsert'));
  if BNew <> nil then BNew.OnClick := CreerEvolution;
  Edit := THEdit(GetControl('PE2_SALARIE'));
  if Edit <> nil then Edit.OnExit := ExitEdit;
  Ecran.Caption := 'Saisie des évolutions (segments PE21 à PE24)'; //PT1
  UpdateCaption(Ecran);
  //DEBUT PT3
  DecodeDate(V_PGI.DateEntree, aa, mm, jj);
  Mois := 1;
  if (mm < 7) and (mm > 3) then Mois := 4;
  if (mm < 10) and (mm > 6) then Mois := 7;
  if mm > 9 then Mois := 10;
  DateDebut := EncodeDate(aa, Mois, 1);
  DateDebut := PlusMois(DateDebut,-3); //PT1
  DateFin := PlusMois(DateDebut, 2);
  DateFin := FinDeMois(DateFin);
  SetControlText('PE2_DATEEFFET', DateToStr(Datedebut));
  SetControltext('PE2_DATEEFFET_', dateToStr(DateFin));
  //FIN PT3
end;

procedure TOF_PGMSAMULEVOLUTIONS.CreerEvolution(Sender: TObject);
begin
  AglLanceFiche('PAY', 'MSAEVOLUTIONS', '', '', 'ACTION=CREATION');
end;

procedure TOF_PGMSAMULEVOLUTIONS.GrilleDblClick(Sender: TObject);
var
  Q_Mul: THQuery;
  St: string;
begin
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  {$IFDEF EAGLCLIENT}
  TFmul(Ecran).Q.TQ.Seek(Grille.Row - 1);
  {$ENDIF}
  If Q_Mul.RecordCount = 0 then //PT2
  begin
       PGIBox('Aucun enregistrement sélectionné',Ecran.Caption);
       Exit;
  end;
  st := Q_MUL.FindField('PE2_SALARIE').AsString + ';' + DateToStr(Q_MUL.FindField('PE2_DATEEFFET').AsDateTime) + ';' +
    Q_MUL.FindField('PE2_TYPEEVOLMSA').AsString;
  AglLanceFiche('PAY', 'MSAEVOLUTIONS', '', St, 'ACTION=MODIFICATION');
end;

procedure TOF_PGMSAMULEVOLUTIONS.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

initialization
  registerclasses([TOF_PGMSAMULEVOLUTIONS]);
end.

