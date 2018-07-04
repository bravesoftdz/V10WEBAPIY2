{***********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 18/02/2000
Modifié le ... : 18/02/2000
Description .. : Multi critère de lancement, de sélection des salariés
Suite ........ : en vue de faire unle pré_calcul de la masse salariale
Suite ........ : et du temps de ptésence pour le calcul de la participation/intéressement
Suite ........ : Confectionne une clause XX_WHERE
Mots clefs ... : PG_MulPrepCalParticp
*****************************************************************}
{
}
unit UTofPG_MulPrepCalParticp;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97, Hqry,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Mul, Fe_Main,
  {$ELSE}
  MaineAgl, eMul,
  {$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOF, UTOB, UTOM, Vierge,
  AGLInit, PgOutils;
type
  TOF_PGMulPrepCalParticp = class(TOF)
  private
    DateAnc: TDateTime;
    DD, DF: TDateTime;
    SuspP, SalS: string;
    procedure ActiveWhere(Sender: TObject);
    procedure LanceFichePrep(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

procedure TOF_PGMulPrepCalParticp.ActiveWhere(Sender: TObject);
var
  WW: THEdit;
  TheSal: string;
  St: string;
  i: Integer;
  Q: TQuery;
begin
  WW := THEdit(GetControl('XX_WHERE'));
  if WW = nil then exit;
  st := GetControlText('SALARIE');
  i := 0;
  while st <> '' do
  begin
    TheSal := readtokenst(st);
    if i > 0 then WW.Text := WW.Text + 'OR';
    WW.Text := WW.Text + ' (PSA_SALARIE <> "' + TheSal + '") ';
    i := i + 1;
  end;
  if GetControlText('PARPARTICIP') = '' then
  begin
    PgiBox ('Veuillez saisir un accord participation/intéressement', Ecran.Caption);
    exit;
  end;
  st := 'SELECT * FROM PARPARTICIPATION WHERE PPA_NUMORDRE=' + GetControlText('PARPARTICIP');
  Q := OpenSql(st, TRUE);
  if not Q.EOF then
  begin
    DateAnc := Q.FindField('PPA_JOURANC').AsDateTime;
    DateAnc := DateAnc - Q.FindField('PPA_NBJANCIENNETE').AsInteger;
    DD := Q.FindField('PPA_DEBUTPAIE').AsDateTime;
    DF := Q.FindField('PPA_FINPAIE').AsDateTime;
    if WW.text <> '' then WW.text := WW.text + 'AND ';
    WW.text := WW.text +
      ' ((PSA_DATEANCIENNETE > "' + USDateTime(IDate1900) + '") AND (PSA_DATEANCIENNETE <= "' + USDateTime(DateAnc) + '"))' +
      ' OR ((PSA_DATEENTREE > "' +  USDateTime(IDate1900) + '") AND (PSA_DATEENTREE <= "' + USDateTime(DateAnc) + '") AND PSA_DATEANCIENNETE IS NULL)';
    SuspP := Q.FindField('PPA_EXCLUSIONSUSP').AsString;
    SalS := Q.FindField('PPA_EXCLUSIONPER').AsString;
    if SuspP = 'X' then
    begin
      if WW.text <> '' then WW.text := WW.text + 'AND ';
      WW.text := WW.text + ' PSA_SUSPENSIONPAIE = "-"';
    end;
    if SalS = 'X' then
    begin
      if WW.text <> '' then WW.text := WW.text + 'AND ';
      WW.text := WW.text + ' ((PSA_DATESORTIE IS NOT NULL) AND (PSA_DATESORTIE > "'+USDateTime(DF) + '"))';
    end;

  end;

  Ferme(Q);

end;

procedure TOF_PGMulPrepCalParticp.LanceFichePrep(Sender: TObject);
var
  st: string;
  LadateD, LadateF, DateDeb, DateFin: TDateTime;
  rep: Integer;
  Q: TQuery;
  OkOk: Boolean;
begin
  if GetControlText('PARPARTICIP') = '' then
  begin
    PgiError('Vous n''avez pas sélectionné de paticipation/intéressement', Ecran.Caption);
    exit;
  end;

  {  rep := PGIAsk ('Attention, la période de paie est supérieure à 1 mois ou à cheval sur 2 mois','Préparation automatique');
    if rep = mrYes then
    begin
    st := 'Select PPU_SALARIE FROM PAIEENCOURS WHERE PPU_DATEDEBUT="'+USDateTime(StrToDate(DD.Text))+'" AND PPU_DATEFIN="'+USDateTime(StrToDate(DF.Text))+'"';
    Q:=OpenSql(st, TRUE);
    if NOT Q.Eof then
       rep := PGIAsk ('Attention, il existe déjà des paies sur la même période. Voulez-vous continuer ?','Préparation automatique');//PT3
    Ferme (Q);
    end;
    St := St + ' ORDER BY ETB_ETABLISSEMENT';
    Q:=OpenSql(st, TRUE);
    While NOT Q.Eof do
       begin

       Q.next;
       end;
    Ferme (Q);
  }
end;

procedure TOF_PGMulPrepCalParticp.OnArgument(Arguments: string);
var
  BtnValidMul: TToolbarButton97;
begin
  inherited;
  BtnValidMul := TToolbarButton97(GetControl('BTNLANCE'));
  if BtnValidMul <> nil then BtnValidMul.OnClick := LanceFichePrep;
  if Arguments = 'PRE' then Ecran.caption := 'Calcul de la masse salariale et de la présence'
  else Ecran.caption := 'Calcul de la participation/intéressement';
  UpdateCaption(TFMul(Ecran));
end;

procedure TOF_PGMulPrepCalParticp.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;
initialization
  registerclasses([TOF_PGMulPrepCalParticp]);
end.

