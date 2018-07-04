{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 02/01/2002
Modifié le ... : 02/01/2002
Description .. : Unit qui lance l'affectation par défaut des données
Suite ........ : Prud'Homales en fonction des valeurs etablissement
Suite ........ :
Mots clefs ... : PAIE;PRUDH
*****************************************************************}
{

}
unit UTOFPG_PRUDH_DEF;

interface

uses windows, StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Vierge, HPanel, HTB97,
  UTofPG_EnvoiPrudh;

type
  TOF_PG_PRUDH_DEF = class(TOF)
    procedure OnArgument(S: string); override;
  private
    procedure RendLeWhere(var St: string);
    procedure LanceAffect(Sender: TObject);
    procedure ImprimPrudh(Sender: TObject);
  end;

implementation

procedure TOF_PG_PRUDH_DEF.ImprimPrudh(Sender: TObject);
var
  FileN: string;
begin
  FileN := GetControlText('NOMFIC');
  if not FileExists(FileN) then
  begin
    PgiBox('Le fichier ' + GetControlText('NOMFIC') + ' est inconnu ?', Ecran.caption);
    exit;
  end;
  EditBorderAcc(FileN);
end;

procedure TOF_PG_PRUDH_DEF.LanceAffect(Sender: TObject);
var
  St: string;
  Okok: Integer;
begin
  st := 'Vous allez lancer l''affectation automatique de vos données prud''homales des salariés !' +
    '#13#10 Voulez vous continuer ?';
  OkOk := PGIAsk(st, Ecran.Caption);
  if OkOk = mrYes then
  begin
    RendLeWhere(st);
    try
      BeginTrans;
      ExecuteSQL('UPDATE SALARIES SET PSA_PRUDHCOLL= (SELECT ETB_PRUDHCOLL FROM ETABCOMPL WHERE ETB_ETABLISSEMENT=PSA_ETABLISSEMENT)' + st);
      ExecuteSQL('UPDATE SALARIES SET PSA_PRUDHSECT= (SELECT ETB_PRUDHSECT FROM ETABCOMPL WHERE ETB_ETABLISSEMENT=PSA_ETABLISSEMENT)' + st);
      ExecuteSQL('UPDATE SALARIES SET PSA_PRUDHVOTE= (SELECT ETB_PRUDHVOTE FROM ETABCOMPL WHERE ETB_ETABLISSEMENT=PSA_ETABLISSEMENT)' + st);
      CommitTrans;
      PGIBox('Fin de traitement', Ecran.Caption);
    except
      Rollback;
      PGIBox('Une erreur est survenue lors de la mise à jour des salariés', Ecran.caption);
    end;
  end;
end;

procedure TOF_PG_PRUDH_DEF.OnArgument(S: string);
var
  BtnVal, BtnImp: TToolbarButton97;
begin
  inherited;
  BtnVal := TToolbarButton97(GetControl('BValider'));
  BtnImp := TToolbarButton97(GetControl('BImprimer'));

  if S = 'DEF' then
  begin
    if (BtnVal <> nil) then BtnVal.OnClick := LanceAffect;
    if (BtnImp <> nil) then BtnImp.Visible := FALSE;
  end
  else
  begin
    if (BtnVal <> nil) then BtnVal.Visible := FALSE;
    if BtnImp <> nil then BtnImp.OnClick := ImprimPrudh;
  end;
end;

procedure TOF_PG_PRUDH_DEF.RendLeWhere(var St: string);
var
  Etab, Etab1: string;
begin
  st := '';
  Etab := GetControlText('CBXETAB');
  Etab1 := GetControlText('CBXETAB1');
  if (Etab <> '') and (Etab1 <> '') then St := ' WHERE PSA_ETABLISSEMENT >= "' + Etab + '" AND PSA_ETABLISSEMENT <= "' + Etab1 + '"';
  if (Etab <> '') and (Etab1 = '') then St := ' WHERE PSA_ETABLISSEMENT = "' + Etab + '"';
  if (Etab = '') and (Etab1 <> '') then St := ' WHERE PSA_ETABLISSEMENT = "' + Etab1 + '"';
end;

initialization
  registerclasses([TOF_PG_PRUDH_DEF]);
end.

