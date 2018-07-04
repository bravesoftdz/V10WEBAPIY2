{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/02/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BUDGETPAIE (BUDGETPAIE)
Mots clefs ... : TOM;BUDGETPAIE
*****************************************************************}
unit BUDGETPAIE_TOM;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche,
  {$ELSE}
  eFiche,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  EntPaie;

type
  TOM_BUDGETPAIE = class(TOM)
  private
    TypBudg: string;
  public
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  end;

implementation

procedure TOM_BUDGETPAIE.OnNewRecord;
var
  Q: TQuery;
begin
  inherited;
  SetField('PBG_TYPEBUDG', TypBudg);
  SetField('PBG_DEBUTEMPLOI', Date);
  SetField('PBG_FINEMPLOI', Idate2099);
  Q := OpenSql('SELECT MAX (PBG_NUMORDRE) NB FROM BUDGETPAIE WHERE PBG_TYPEBUDG="' + TypBudg + '"', True);
  if not Q.EOF then
    SetField('PBG_NUMORDRE', Q.FindField('NB').AsInteger + 1)
  else SetField('PBG_NUMORDRE', 0);
  Ferme(Q);
end;

procedure TOM_BUDGETPAIE.OnDeleteRecord;
begin
  inherited;
end;

procedure TOM_BUDGETPAIE.OnUpdateRecord;
begin
  inherited;
end;

procedure TOM_BUDGETPAIE.OnAfterUpdateRecord;
begin
  inherited;
end;

procedure TOM_BUDGETPAIE.OnLoadRecord;
begin
  inherited;
end;

procedure TOM_BUDGETPAIE.OnChangeField(F: TField);
begin
  inherited;
end;

procedure TOM_BUDGETPAIE.OnArgument(S: string);
var
  Numero, Libelle: string;
  Num: Integer;
begin
  inherited;
  TypBudg := ReadTokenst(S); // Memorisation du type de budget (Offre emploi)
  // Visibilité des codes Organisations
  for num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PGNbreStatOrg then break;
    libelle := '';
    if Num = 1 then libelle := VH_Paie.PGLibelleOrgStat1;
    if Num = 2 then libelle := VH_Paie.PGLibelleOrgStat2;
    if Num = 3 then libelle := VH_Paie.PGLibelleOrgStat3;
    if Num = 4 then libelle := VH_Paie.PGLibelleOrgStat4;
    if (Libelle <> '') then
    begin
      SetControlVisible('PBG_TRAVAILN' + Numero, TRUE);
      SetControlVisible('TPBG_TRAVAILN' + Numero, TRUE);
      SetControlText('TPBG_TRAVAILN' + Numero, Libelle);
    end;
  end;
  //Visibilité du code statistique
  if VH_Paie.PGLibCodeStat <> '' then
  begin
    SetControlVisible('TPBG_CODESTAT', TRUE);
    SetControlText('TPBG_CODESTAT', VH_Paie.PGLibCodeStat);
    SetControlVisible('PBG_CODESTAT', TRUE);
  end;

  // Visibilité des codes combos libres
  for num := 1 to VH_Paie.PgNbCombo do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PgNbCombo then break;
    libelle := '';
    if Num = 1 then libelle := VH_Paie.PgLibCombo1;
    if Num = 2 then libelle := VH_Paie.PgLibCombo2;
    if Num = 3 then libelle := VH_Paie.PgLibCombo3;
    if Num = 4 then libelle := VH_Paie.PgLibCombo4;
    if (Libelle <> '') then
    begin
      SetControlVisible('PBG_LIBREPCMB' + Numero, TRUE);
      SetControlVisible('TPBG_LIBREPCMB' + Numero, TRUE);
      SetControlText('TPBG_LIBREPCMB' + Numero, Libelle);
    end;
  end;

  for num := 1 to VH_Paie.PgNbSalLib do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PgNbSalLib then break;
    SetControlVisible('PBG_SALAIREMOIS' + Numero, TRUE);
    SetControlVisible('PBG_SALAIRANN' + Numero, TRUE);
    SetControlVisible('SALLIB' + Numero, TRUE);
    if Num = 1 then Libelle := VH_Paie.PgSalLib1
    else
      if Num = 2 then Libelle := VH_Paie.PgSalLib2 else
      if Num = 3 then Libelle := VH_Paie.PgSalLib3 else
      if Num = 4 then Libelle := VH_Paie.PgSalLib4 else
      if Num = 5 then Libelle := VH_Paie.PgSalLib5;
    SetControlText('SALLIB' + Numero, Libelle);
    if num = 1 then
    begin
      SetControlVisible('LBLMENSUEL', TRUE);
      SetControlVisible('LBLANNUEL', TRUE);
    end;
  end;

end;

procedure TOM_BUDGETPAIE.OnClose;
begin
  inherited;
end;

procedure TOM_BUDGETPAIE.OnCancelRecord;
begin
  inherited;
end;

initialization
  registerclasses([TOM_BUDGETPAIE]);
end.

