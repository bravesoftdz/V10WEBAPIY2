{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 29/11/2006
Modifié le ... :   /  /
Description .. : Gestion : Fiche libre utilisable dans le cas d'utilitaire
Mots clefs ... : PAIE;STANDARD
*****************************************************************}
{
PT1 MF 22/01/2007  V_720  Titres restaurant (ACCOR): On utilise la
                                   fiche LIBRE pour la saisie du CC_LIBRE
                                   de la tablette des codes distribution
PT2 PH 14/05/2007 V_720 Mise en place moulinette multidossier plan de paie partagé ou non
}
unit UTofPG_LIBRE;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
  Fe_Main,
{$ELSE}
  MainEagl,
{$ENDIF}
  HDebug, HCtrls, HEnt1, HMsgBox, UTOF, Ed_Tools, UTOB;
type
  TOF_PG_LIBRE = class(TOF)
  private
    CodeDistribution, CCcode, CClibre: string; // PT1
    procedure LANCEREAFF(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override; // PT1
  end;

implementation
uses pgOutils;

procedure TOF_PG_LIBRE.LANCEREAFF(Sender: TObject);
// DEB PT2
var RDBTN, RDBTN2: TRadiobutton;
begin
  RDBTN := TRadiobutton(GetControl('RDB1'));
  if RDBTN <> nil then
  begin
    if RDBTN.Checked then ReaffectationNoDossier()
    else
    begin
      RDBTN2 := TRadiobutton(GetControl('RDB2'));
      if RDBTN2 <> nil then
      begin
        if RDBTN2.Checked then ReaffNoDossierMulti()
        else ReparationNoDossier();
      end;
    end;
// FIN PT2
  end;
end;

procedure TOF_PG_LIBRE.OnArgument(Arguments: string);
var
  BtnVal: TToolbarButton97;
begin
  inherited;
  BtnVal := TToolbarButton97(GetControl('BValider'));
  if BtnVal = nil then exit;
  if (Trim(Arguments)) = 'REAFFNODOSSIER' then
  begin
    Ecran.Caption := 'Utilitaire de contrôle et de réaffectation de numéro dossier';
    BtnVal.OnClick := LANCEREAFF;
    SetControlText('LBL1', 'Cet utilitaire permet d''attribuer un numéro dossier pour');
    SetControlText('LBL2', 'l''ensemble du plan de paie dans le cadre de l''utilisation');
    SetControlText('LBL3', 'du plan de paie partagé pour plusieurs sociétés.');
    SetControlVisible('LBL2', TRUE);
    SetControlVisible('LBL3', TRUE);
// DEB PT2
    SetControlVisible('RDB1', TRUE);
    SetControlVisible('RDB2', TRUE);
// FIN PT2
    UpdateCaption(Ecran);
//d PT1
  end
  else
  // Cas de traitement : titres restaurant ACCOR
  begin
    CodeDistribution := ReadTokenSt(Arguments);
    if (CodeDistribution) = 'CODISTRIBUTION' then
    begin
      CCcode := ReadTokenSt(Arguments);
      CClibre := ReadTokenSt(Arguments);
      Ecran.Caption := 'Point de distribution ACCOR';
      SetControlVisible('LBL1', FALSE);
      SetControlVisible('LACCORDISTRIB', TRUE);
      SetControlVisible('ACCORDISTRIB', TRUE);
      SetControlEnabled('ACCORDISTRIB', TRUE);
      SetControlText('ACCORDISTRIB', CClibre);
    end;
  end;
// f PT1

end;

// d PT1
// maj tablette des codes distribution

procedure TOF_PG_LIBRE.OnUpdate;
var
  CodDis: string;
begin
  inherited;
  if (CodeDistribution = 'CODISTRIBUTION') then
  begin
    CodDis := GetControlText('ACCORDISTRIB');
    if ((CodDis <> '') and (CodDis <> CClibre)) then
    begin
      if ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="PCD" AND CC_CODE="' + CCcode + '"') then
        ExecuteSQL('UPDATE CHOIXCOD SET CC_LIBRE="' + CodDis + '" WHERE CC_CODE="' + CCcode + '" AND CC_TYPE="PCD"');
    end;
  end;

end;
// f PT1

initialization
  registerclasses([TOF_PG_LIBRE]);
end.

