{***********UNITE*************************************************
Auteur  ...... :  JL
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMSAMULPERIODES ()
Mots clefs ... : TOF;PGMSAMULPERIODES
*****************************************************************
PT1 21/04/2005 V_60 JL FQ 12192 Trimestre précédent par défaut pour préparation fichier
PT2 19/05/2005 V_60 JL FQ 12296 Modif des caption ecran
PT3 17/05/2006 V_65 JL Gestion filtre sur établissement pour affichage grille dans MSA_ENTREPRISE
PT4 25/07/2007 V_80 JL FQ 14570 Ajout exclusion salariés sortis
}
unit UTOFPGMul_MSAPeriodes;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} mul, FE_Main, DBGrids, HDB,
  {$ELSE}
  eMul, uTob, MainEAgl,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, HTB97, HQry, ParamDat, EntPaie, PGOutils,PGOutils2;

type
  TOF_PGMSAMULPERIODES = class(TOF)
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
  private
    Arg: string;
    {$IFNDEF EAGLCLIENT}
    Grille: THDBGrid;
    {$ELSE}
    Grille: THGrid;
    {$ENDIF}
    procedure CreerPeriode(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);
    procedure DateElipsisClick(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure ClickSortie(Sender : TObject);
    procedure VerifDate(Sender : TObject);
  end;

implementation

procedure TOF_PGMSAMULPERIODES.OnLoad;
begin
  inherited;
  //DEBUT PT4
  If TCheckBox(GetControl('CKSORTIE')) <> Nil then
  begin
    If GetCheckBoxState('CKSORTIE') = CbChecked then SetControlText('XX_WHERE',' AND PE3_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+
    '(PSA_DATESORTIE>="'+UsDateTime(StrtoDate(GetControlText('DATEARRET')))+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL))')
    else SetControlText('XX_WHERE','');
  end;
  //FIN PT4
end;

procedure TOF_PGMSAMULPERIODES.OnArgument(S: string);
var
  BNew: TToolBarButton97;
  DateDebut, DateFin: TDateTime;
  THDate, Edit: THEdit;
  aa, mm, jj, Mois: word;
  check : TCheckBox;
begin
  inherited;
  Arg := ReadTokenPipe(S, ';');
  {$IFNDEF EAGLCLIENT}
  Grille := THDBGrid(GetControl('Fliste'));
  {$ELSE}
  Grille := THGrid(GetControl('Fliste'));
  {$ENDIF}
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;
  BNew := TToolbarButton97(getcontrol('BInsert'));
  if BNew <> nil then BNew.OnClick := CreerPeriode;
  DecodeDate(V_PGI.DateEntree, aa, mm, jj);
  Mois := 1;
  if (mm < 7) and (mm > 3) then Mois := 4;
  if (mm < 10) and (mm > 6) then Mois := 7;
  if mm > 9 then Mois := 10;
  DateDebut := EncodeDate(aa, Mois, 1);
  DateDebut := PlusMois(DateDebut,-3); //PT1
  DateFin := PlusMois(DateDebut, 2);
  DateFin := FinDeMois(DateFin);
  SetControlText('PE3_DATEDEBUT', DateToStr(Datedebut));
  SetControltext('PE3_DATEFIN', dateToStr(DateFin));
  THDate := THEdit(GetControl('PE3_DATEDEBUT'));
  if THDate <> nil then THDate.OnElipsisClick := DateElipsisClick;
  THDate := THEdit(GetControl('PE3_DATEFIN'));
  if THDate <> nil then
  begin
    THDate.OnElipsisClick := DateElipsisClick;
    THDate.OnExit := VerifDate;
  end;
  if Arg = 'GENERATION' then
  begin
    TFMul(Ecran).BInsert.Visible := False;
    TFMul(Ecran).Caption := 'Préparation du fichier MSA (Segment PE41)';
  end
  else TFMul(Ecran).Caption := 'Saisie des périodes (segment PE31)'; //PT2
  UpdateCaption(TFMul(Ecran));
  Edit := THEdit(GetControl('PE3_SALARIE'));
  if Edit <> nil then Edit.OnExit := ExitEdit;
  //DEBUT PT4
  Check := TCheckBox(GetControl('CKSORTIE'));
  If Check <> Nil then Check.OnClick := ClickSortie;
  //FIN PT4
end;

procedure TOF_PGMSAMULPERIODES.CreerPeriode(Sender: TObject);
begin
  AglLanceFiche('PAY', 'MSAPERIODESPE31', '', '', 'ACTION=CREATION');
end;

procedure TOF_PGMSAMULPERIODES.GrilleDblClick(Sender: TObject);
var
  Q_Mul: THQuery;
  St,Etab : string;
begin
  if Arg <> 'GENERATION' then
  begin
    Q_Mul := THQuery(Ecran.FindComponent('Q'));
    {$IFDEF EAGLCLIENT}
    TFmul(Ecran).Q.TQ.Seek(Grille.Row - 1);
    {$ENDIF}
    st := Q_MUL.FindField('PE3_SALARIE').AsString + ';' + Q_MUL.FindField('PE3_DATEDEBUT').AsString + ';' +
      Q_MUL.FindField('PE3_DATEFIN').AsString + ';' + IntToStr(Q_MUL.FindField('PE3_ORDRE').AsInteger);
    AglLanceFiche('PAY', 'MSAPERIODESPE31', '', St, 'ACTION=MODIFICATION');
  end
  else
  begin
    St := GetControlText('PE3_DATEDEBUT') + ';' + GetControlText('PE3_DATEFIN');
//    If GetCheckBoxState('CALCULTCP') = cbchecked then St := St + ';OUI'
//    else St := St + ';NON';
    Etab := GetControlText('PE3_ETABLISSEMENT'); //PT3
    St := St+';'+Etab;
    AglLanceFiche('PAY', 'MSA_ENTREPRISE', '', '', St);
  end;
end;

procedure TOF_PGMSAMULPERIODES.DateElipsisClick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGMSAMULPERIODES.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGMSAMULPERIODES.ClickSortie(Sender : TObject);    //PT4
begin
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_PGMSAMULPERIODES.VerifDate(Sender : TObject);
var DateD,DateF : TDateTime;
begin
  If THEdit(Sender) = Nil then Exit;
  DateD := StrToDate(GetControlText('PE3_DATEDEBUT'));
  DateF := StrToDate(GetControlText('PE3_DATEFIN'));
  If DateF < DateD then
  begin
    PGIBox('La date de fin ne peut pas être inférieure à la date de début',Ecran.Caption);
    SetFocusControl(THEdit(Sender).Name);
  end;
end;

initialization
  registerclasses([TOF_PGMSAMULPERIODES]);
end.

