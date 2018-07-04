{***********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 29/08/2003
Modifié le ... : 29/08/2003
Description .. : Multi de reaffectation des codes PCS 1982 à 2003
Suite ........ :
Suite ........ :
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   15/10/2003 SB V_42 Maj de date modification de la table des salaries pour introduction PAIE
PT2   15/10/2003 PH      Prise en compte des 2 cas affectables et non affectables dans la clause XX_WHERE
                         au lieu de passer par des vues ==> listes utilisent des tables avec jointure
}

unit UTofPG_ReaffCodePcs;

interface
uses StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HTB97,
  HStatus,
{$IFNDEF EAGLCLIENT}
  HDB, DBGrids, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, mul,
{$ELSE}
  MaineAGL, eMul,
{$ENDIF}
  HCtrls,
  HEnt1,
  vierge,
  HMsgBox,
  Hqry,
  UTOF,
  UTOB,
  ParamSoc,
  EntPaie,
  AGLInit;
type
  TOF_PGReaffCodePcs = class(TOF)
  private
    Trait: string;
    Q_Mul: THQuery;
    BtnReaff: TToolbarButton97;
//       procedure GrilleDblClick (Sender: TObject);
    procedure LanceTrait(Sender: TObject);
    procedure ActiveWhere;
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;

  end;

implementation
uses PgOutils;

procedure TOF_PGReaffCodePcs.ActiveWhere;
var WW: THEdit;
begin
  WW := THEdit(GetControl('XX_WHERE'));
  if WW <> nil then
  begin
// PT2   15/10/2003 PH      Prise en compte des 2 cas affectables et non affectables dans la clause XX_WHERE
    if Trait <> 'N' then WW.Text := ' (PSA_CODEEMPLOI <> "" OR PSA_CODEEMPLOI IS NULL)' +
      ' AND EXISTS (SELECT PCE_CODEEMPLOI FROM PCSESE WHERE PCE_CODEEMPLOI = SALARIES.PSA_CODEEMPLOI)' +
        ' AND PCE_CODEEMPLOI=SALARIES.PSA_CODEEMPLOI AND PCE_ATRAITER="X"'
    else WW.Text := ' (PSA_CODEEMPLOI <> "" OR PSA_CODEEMPLOI IS NULL)' +
      ' AND NOT EXISTS (SELECT PCE_CODEEMPLOI FROM PCSESE WHERE PCE_CODEEMPLOI = SALARIES.PSA_CODEEMPLOI)';
  end;
end;
{
procedure TOF_PGReaffCodePcs.GrilleDblClick(Sender: TObject);
var Sal :String ;
begin
  if Not Assigned (Q_Mul) then exit ;
  Sal :=TFMUL(Ecran).Q.findfield('PSA_SALARIE').Asstring ;
  if Sal <> '' then AGLLanceFiche ('PAY', 'SALARIE', '',Sal, '') ;
end;
}

procedure TOF_PGReaffCodePcs.LanceTrait(Sender: TObject);
var Q: TQuery;
  st: string;
  ii, Rep: Integer;
begin

  if Assigned(Q_MUL) and not Q_MUL.EOF then
  begin
    st := 'SELECT COUNT (*) NBRE FROM PCSESE C1 WHERE EXISTS (SELECT C2.PCE_CODEEMPLOI FROM PCSESE C2 ' +
      'WHERE C1.PCE_CODEEMPLOI = C2.PCE_CODEEMPLOI AND C1.PCE_NUMORDRE <> C2.PCE_NUMORDRE ' +
      ' AND C2.PCE_ATRAITER="X") AND C1.PCE_ATRAITER="X"';
    Q := OpenSql(st, true);
    if not Q.eof then ii := Q.findfield('NBRE').AsInteger
    else ii := 0;
    Ferme(Q);
    if ii = 0 then
    begin
      Rep := PGIAsk('Voulez vous réaffecter automatiquement vos codes PCS', Ecran.Caption);
      if Rep = mrYes then
      begin
        ExecuteSQL('UPDATE SALARIES SET PSA_CODEPCS82=PSA_CODEEMPLOI'); // Sauvegarde ancien code PCS
        st := 'UPDATE SALARIES SET PSA_CODEEMPLOI = (SELECT PCE_CODEPCSESE FROM PCSESE WHERE PCE_CODEEMPLOI = PSA_CODEEMPLOI AND PCE_ATRAITER="X") ' +
          ',PSA_DATEMODIF="' + UsTime(Now) + '" ' + //PT1
          ' WHERE (PSA_CODEEMPLOI <> "" OR PSA_CODEEMPLOI IS NULL)';
        ExecuteSQL(st);
        PgiBox('Réaffectation terminée', Ecran.Caption);
        SetParamSoc('SO_PGPCS2003', 'X');
        VH_Paie.PGPCS2003 := TRUE;
        ChangePCS82A03();
      end;
    end
    else PgiBox('Traitement impossible, vous avez des doublons dans votre table de correspondance', Ecran.Caption);
  end
  else PgiBox('Traitement impossible car votre liste est vide', Ecran.Caption);
end;

procedure TOF_PGReaffCodePcs.OnArgument(Arguments: string);
begin
  inherited;
  Trait := Trim(ReadTokenSt(Arguments));
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  BtnReaff := TToolbarButton97(GetControl('REAFFECTATION'));
  if BtnReaff <> nil then BtnReaff.onClick := LanceTrait;
  if Trait = 'N' then
  begin
    SetControlVisible('REAFFECTATION', FALSE);
    if Q_MUL <> nil then TFMul(Ecran).SetDBListe('PGERRPCS');
    Ecran.Caption := 'Salariés dont le code PCS n''est pas réaffectable';
    UpdateCaption(TFmul(Ecran));
  end;

end;

procedure TOF_PGReaffCodePcs.OnLoad;
begin
  inherited;
  ActiveWhere();
end;

initialization
  registerclasses([TOF_PGReaffCodePcs]);
end.

