{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 22/05/2002
Modifié le ... :   /  /
Description .. : Integration de la saisie des éléments variables
Suite ........ :
Mots clefs ... : PAIE;PGINTEGPRIM
*****************************************************************}  
{
PT1   : 22/03/2004 PH V_500 Prise en compte des donnees participation stockées dans histosaisprim
---- PH 10/08/2005 Suppression directive de compil $IFDEF AGL550B ----

}
unit UTofPG_IntegPrim;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DBGrids,
  {$ENDIF}
  Grids, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Vierge, ed_tools, AGLInit, EntPaie;

type
  TOF_PGIntegPrim = class(TOF)
  private
    DateDeb, DateFin: THEDit;
    procedure LanceRecopie(Sender: TObject);
    procedure ExecuteRecopie;
    procedure DateDebExit(Sender: TObject);
    procedure DateFinExit(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation

uses PGVisuObjet, PgOutils2;
{ Fonction qui va lancer la recopie des éléments variables saisis dans
la fourchette de dates
}
procedure TOF_PGIntegPrim.ExecuteRecopie;
var
   St: string;
  q: TQuery;
  TPrim, TR: TOB;
  T1, T2: TOB;
  DD, DF: TDateTime;
  NomP, NomR: string;
  i, k, j, FP, FR: integer;
begin
  DD := StrToDate(DateDeb.Text);
  DF := StrToDate(DateFin.Text);
  if DD > DF then
  begin
    PGIBox('La date de début est supérieure à la date de fin ?', 'Abandon du traitement');
    exit;
  end;
  try
    // PT1   : 22/03/2004 PH V_500 Prise en compte des donnees participation stockées dans histosaisprim
    st := 'SELECT * FROM HISTOSAISPRIM WHERE PSP_DATEDEBUT >="' + UsDateTime(DD) + '"' +
      ' AND PSP_DATEFIN <="' + UsDateTime(DF) + '" AND PSP_ORDRE < 100';
    // FIN PT1
    TPrim := TOB.Create('Les primes du mois', nil, -1);
    Q := OpenSql(st, TRUE);
    TPrim.LoadDetailDB('HISTOSAISPRIM', '', '', Q, FALSE);
    ferme(Q);
    InitMoveProgressForm(nil, 'Chargement des données de la saisie', 'Veuillez patienter SVP ...', TPrim.detail.count + 3, TRUE, TRUE);
    TR := TOB.Create('Les Historiques', nil, -1);
    FP := PrefixeToNum('PSP');
    FR := PrefixeToNum('PSD');
    for k := 0 to TPrim.detail.Count - 1 do
    begin
      T1 := Tprim.detail[k];
      T2 := TOB.Create('HISTOSAISRUB', TR, -1);
      for i := 1 to T1.NbChamps do
      begin
        NomP := T1.GetNomChamp(i);
        if (Copy(NomP, 5, length(NomP) - 4) = 'LIBELLE') then
           continue ; // On exclut le traitement du champ libellé

        NomR := 'PSD_' + Copy(NomP, 5, length(NomP) - 4);
        j := T2.GetNumChamp(NomR);
        if j > 0 then
        begin
          if V_PGI.DECHAMPS[FP, i].Tipe = V_PGI.DECHAMPS[FR, j].Tipe
            then T2.PutValeur(j, T1.GetValeur(i));
        end;
      end;
      MoveCurProgressForm('Ligne en cours : ' + IntToStr(k));
    end;
    MoveCurProgressForm('Réaffectation des numéros en cours');
    for i := 0 to TR.Detail.count - 1 do
    begin
      T2 := TR.Detail[i];
      T2.PutValue('PSD_ORDRE', T2.GetValue('PSD_ORDRE') + 100);
      T2.PutValue('PSD_DATECOMPT', IDate1900);
    end;
    MoveCurProgressForm('Intégration en cours');
    TR.SetAllModifie (TRUE) ;
    if TR.Detail.count - 1 >= 1 then TR.InsertOrUpdateDb(TRUE);
    MoveCurProgressForm('Mise à jour de la saisie des éléménts variables en cours ');
    for i := 0 to TPrim.Detail.count - 1 do
    begin
      T2 := TPrim.Detail[i];
      T2.PutValue('PSP_DATEINTEGRAT', Date);
      // PT1   : 22/03/2004 PH V_500 Prise en compte des donnees participation stockées dans histosaisprim
      T2.PutValue('PSP_AREPORTER', 'NON');
      // FIN PT1
    end;
    if TPrim.Detail.count - 1 >= 1 then TPrim.InsertOrUpdateDb(FALSE);
    TPrim.free;
//  Tprim := nil;
    TR.Free;
//  TR := nil;
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la Recopie', Ecran.caption);
  end;

  FiniMoveProgressForm();
end;

procedure TOF_PGIntegPrim.LanceRecopie(Sender: TObject);
begin

  if PGIAsk('Confirmez vous la recopie des éléments variables', Ecran.Caption) = mrYes then ExecuteRecopie;
end;

procedure TOF_PGIntegPrim.DateDebExit(Sender: TObject);
begin
  if not IsValidDate(DateDeb.Text) then
  begin
    PGIBox('La date de début n''est pas valide', 'Saisie par Rubrique');
    DateDeb.SetFocus;
  end;
end;

procedure TOF_PGIntegPrim.DateFinExit(Sender: TObject);
var
  Date1, Date2: TDateTime;
begin
  if not IsValidDate(DateFin.Text) then
  begin
    PGIBox('La date de fin n''est pas valide', Ecran.caption);
    DateFin.SetFocus;
    exit;
  end;
  Date2 := StrToDate(DateFin.Text);
  Date1 := StrToDate(DateDeb.Text);
  if Date1 > Date2 then
  begin
    PGIBox('La date de début est supérieure à la date de fin', Ecran.caption);
    DateDeb.SetFocus;
    DateFin.Text := '';
    exit;
  end;
  if (FINDEMOIS(Date1) <> FINDEMOIS(Date2)) then
  begin
    PGIBox('Les dates de début et de fin doivent être comprise dans le même mois', Ecran.caption);
    DateFin.SetFocus;
    exit;
  end;
end;


procedure TOF_PGIntegPrim.OnArgument(Arguments: string);
var
  BtnVal: TToolbarButton97;
  DebPer, FinPer, ExerPerEncours: string;
begin
  inherited;
  DateDeb := ThEdit(getcontrol('XX_VARIABLEDEB'));
  if DateDeb <> nil then DateDeb.OnExit := DateDebExit;
  DateFin := ThEdit(getcontrol('XX_VARIABLEFIN'));
  if DateFin <> nil then DateFin.OnExit := DateFinExit;
  if RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer) = True then
  begin
    if DateDeb <> nil then DateDeb.text := DebPer;
    if DateFin <> nil then DateFin.text := FinPer;
  end;
  BtnVal := TToolbarButton97(GetControl('BValider'));
  if BtnVal <> nil then BtnVal.OnClick := LanceRecopie;
end;

initialization
  registerclasses([TOF_PGIntegPrim]);
end.

