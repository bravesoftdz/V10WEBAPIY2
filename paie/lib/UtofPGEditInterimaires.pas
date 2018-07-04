{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITINTERIM ()
Mots clefs ... : TOF;PGEDITINTERIM
*****************************************************************
PT1 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
---- JL 20/03/2006 modification clé annuaire ----
}
unit UtofPGEditInterimaires;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ELSE}
  eQRS1,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, EntPaie,
  PGEditOutils,PgEditOutils2, ParamDat, HQry, P5Def, ParamSoc;

type
  TOF_PGEDITINTERIM = class(TOF)
    procedure OnUpdate; override;
    procedure OnArgument(S: string); override;
  private
    TriCN: string;
    procedure DateElipsisclick(Sender: TObject);
    procedure ChangeLieuTrav(Sender: TObject);
  end;

implementation

procedure TOF_PGEDITINTERIM.OnUpdate;
var
  SQL, Where, TriEtab, Tri: string;
  Pages: TPageControl;
  Int, Sta: Boolean;
begin
  TriEtab := '';
  Tri := '';
  Pages := TPageControl(GetControl('Pages'));
  Where := RecupWhereCritere(Pages);
  if Where <> '' then Where := Where + ' AND '
  else Where := 'Where ';
  if GetControlText('CINT') = 'X' then Int := True
  else Int := False;
  if GetControlText('CSTA') = 'X' then Sta := True
  else Sta := False;
  if (Int = True) and (Sta = True) then Where := Where + ' PSI_TYPEINTERIM IN ("INT","STA") AND ';
  if (Int = True) and (Sta = False) then Where := Where + ' PSI_TYPEINTERIM="INT" AND ';
  if (Int = False) and (Sta = True) then Where := Where + ' PSI_TYPEINTERIM="STA" AND ';
  if TriCN = '' then TriCN := ' ORDER BY ';
  if GetControlText('CETAB') = 'X' then
  begin
    TriEtab := TriCN + 'PEI_ETABLISSEMENT,';
    SetControlText('XX_RUPTURE2', 'PEI_ETABLISSEMENT');
  end
  else
  begin
    TriEtab := TriCN;
    SetControlText('XX_RUPTURE2', '');
  end;
  if GetControlText('ALPHA') = 'X' then Tri := TriEtab + 'PSI_LIBELLE,'
  else Tri := TriEtab;

  SQL := 'SELECT PEI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM,PSI_ADRESSE1,PSI_ADRESSE2' +
    ',PSI_ADRESSE3,PSI_CODEPOSTAL,PSI_VILLE,PSI_TELEPHONE,PSI_PORTABLE,' +
    'PSI_EMAIL,PSI_DATENAISSANCE,PEI_ORDRE,PEI_ETABLISSEMENT,PEI_TRAVAILN1,' +
    'PEI_TRAVAILN2,PEI_TRAVAILN3,PEI_TRAVAILN4,PEI_CODESTAT,PEI_COEFFICIENT,' +
    'PEI_QUALIFICATION,PEI_CODEEMPLOI,PEI_LIBELLEEMPLOI,PSI_SEXE,PSI_NUMEROSS,' +
    'PEI_DEBUTEMPLOI,PEI_FINEMPLOI,PEI_AGENCEINTGU,PEI_CENTREFORMGU,PSI_TYPEINTERIM,PEI_TUTEURSAL' +
    ',PEI_SALREMPL,PEI_MOTIFCTINT,PEI_NBHEUREHEB,PEI_TAUXHORAIRE' +
    ' FROM EMPLOIINTERIM LEFT JOIN INTERIMAIRES ON PEI_INTERIMAIRE=PSI_INTERIMAIRE ' +
    Where + 'PEI_DEBUTEMPLOI<="' + UsDateTime(StrToDate(GetControlText('DATEFIN'))) + '"' +
    ' AND PEI_FINEMPLOI>="' + UsDateTime(StrToDate(GetControlText('DATEDEBUT'))) + '"' +
    ' AND PEI_ETABLISSEMENT>="' + GetControlText('ETAB1') + '" AND PEI_ETABLISSEMENT<="' + GetControlText('ETAB2') + '"' +
    Tri + 'PEI_INTERIMAIRE,PEI_ORDRE';
  TFQRS1(Ecran).WhereSQL := SQL;
end;

procedure TOF_PGEDITINTERIM.OnArgument(S: string);
var
  Defaut: THEdit;
  Min, Max: string;
  TC: TCheckBox;
  Num: Integer;
begin
  inherited;
  Ecran.Caption := 'Edition';
  UpdateCaption(Ecran);
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    {$IFNDEF EAGLCLIENT}
//    Defaut.text := V_PGI_env.LibDossier; //PT1 Mise en commentaire
  {$ENDIF}
  Defaut.text := GetParamSoc('SO_LIBELLE');
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  SetControlText('ETAB1', Min);
  SetControlText('ETAB2', Max);
  Defaut := THEdit(GetControl('DATEDEBUT'));
  if Defaut <> nil then Defaut.OnElipsisClick := DateElipsisClick;
  Defaut := THEdit(GetControl('DATEFIN'));
  if Defaut <> nil then Defaut.OnElipsisClick := DateElipsisClick;
  // Gestion affichage des champs de l'onglet compléments :
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PEI_TRAVAILN' + IntToStr(Num)), GetControl('TPEI_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PEI_CODESTAT'), GetControl('TPEI_CODESTAT'));
  Defaut := THEdit(GetControl('PEI_TRAVAILN1'));
  if Defaut <> nil then
  begin
    SetControlVisible('PEI_TRAVAILN1_', Defaut.Visible);
    SetControlVisible('CN1', Defaut.Visible);
  end;
  Defaut := THEdit(GetControl('PEI_TRAVAILN2'));
  if Defaut <> nil then
  begin
    SetControlVisible('PEI_TRAVAILN2_', Defaut.Visible);
    SetControlVisible('CN2', Defaut.Visible);
  end;
  Defaut := THEdit(GetControl('PEI_TRAVAILN3'));
  if Defaut <> nil then
  begin
    SetControlVisible('PEI_TRAVAILN3_', Defaut.Visible);
    SetControlVisible('CN3', Defaut.Visible);
  end;
  Defaut := THEdit(GetControl('PEI_TRAVAILN4'));
  if Defaut <> nil then
  begin
    SetControlVisible('PEI_TRAVAILN4_', Defaut.Visible);
    SetControlVisible('CN4', Defaut.Visible);
  end;
  Defaut := THEdit(GetControl('PEI_CODESTAT'));
  if Defaut <> nil then
  begin
    SetControlVisible('PEI_CODESTAT_', Defaut.Visible);
    SetControlVisible('CN5', Defaut.Visible);
  end;
  // Fin affichage

  TC := TCheckBox(GetControl('CN1'));
  if TC <> nil then TC.OnClick := ChangeLieuTrav;
  TC := TCheckBox(GetControl('CN2'));
  if TC <> nil then TC.OnClick := ChangeLieuTrav;
  TC := TCheckBox(GetControl('CN3'));
  if TC <> nil then TC.OnClick := ChangeLieuTrav;
  TC := TCheckBox(GetControl('CN4'));
  if TC <> nil then TC.OnClick := ChangeLieuTrav;
  TC := TCheckBox(GetControl('CN5'));
end;

procedure TOF_PGEDITINTERIM.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGEDITINTERIM.ChangeLieuTrav(Sender: TObject);
var
  CN1, CN2, CN3, CN4, CN5: TCheckBox;
  LN1, LN2, LN3, LN4, LN5: THLabel;
begin
  TriCN := '';
  SetControlText('XX_RUPTURE1', '');
  SetControlText('XX_VARIABLE1', '');
  CN1 := TCheckBox(GetControl('CN1'));
  CN2 := TCheckBox(GetControl('CN2'));
  CN3 := TCheckBox(GetControl('CN3'));
  CN4 := TCheckBox(GetControl('CN4'));
  CN5 := TCheckBox(GetControl('CN5'));
  LN1 := THLabel(GetControl('TPEI_TRAVAILN1'));
  LN2 := THLabel(GetControl('TPEI_TRAVAILN2'));
  LN3 := THLabel(GetControl('TPEI_TRAVAILN3'));
  LN4 := THLabel(GetControl('TPEI_TRAVAILN4'));
  LN5 := THLabel(GetControl('TPEI_CODESTAT'));
  if (CN1 <> nil) and (CN2 <> nil) and (CN3 <> nil) and (CN4 <> nil) and (CN5 <> nil) then
  begin
    CN1.Enabled := True;
    CN2.Enabled := True;
    CN3.Enabled := True;
    CN4.Enabled := True;
    CN5.Enabled := True;
    if CN1.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PEI_TRAVAILN1');
      SetControltext('XX_VARIABLE1', LN1.Caption);
      TriCN := ' ORDER BY PEI_TRAVAILN1,';
      CN2.Enabled := False;
      CN3.Enabled := False;
      CN4.Enabled := False;
      CN5.Enabled := False;
    end;
    if CN2.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PEI_TRAVAILN2');
      SetControltext('XX_VARIABLE1', LN2.Caption);
      TriCN := ' ORDER BY PEI_TRAVAILN2,';
      CN1.Enabled := False;
      CN3.Enabled := False;
      CN4.Enabled := False;
      CN5.Enabled := False;
    end;
    if CN3.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PEI_TRAVAILN3');
      SetControltext('XX_VARIABLE1', LN3.Caption);
      TriCN := ' ORDER BY PEI_TRAVAILN3,';
      CN1.Enabled := False;
      CN2.Enabled := False;
      CN4.Enabled := False;
      CN5.Enabled := False;
    end;
    if CN4.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PEI_TRAVAILN4');
      SetControltext('XX_VARIABLE1', LN4.Caption);
      TriCN := ' ORDER BY PEI_TRAVAILN4,';
      CN1.Enabled := False;
      CN2.Enabled := False;
      CN3.Enabled := False;
      CN5.Enabled := False;
    end;
    if CN5.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PEI_CODESTAT');
      SetControltext('XX_VARIABLE1', LN5.Caption);
      TriCN := ' ORDER BY PEI_CODESTAT,';
      CN1.Enabled := False;
      CN2.Enabled := False;
      CN3.Enabled := False;
      CN4.Enabled := False;
    end;
  end;
end;

initialization
  registerclasses([TOF_PGEDITINTERIM]);
end.


