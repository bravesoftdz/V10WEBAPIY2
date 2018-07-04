unit UTOFPG_ORIGINERUB;
{
PT1  13/09/2004 PH V_50 FQ 11581 Controle origine de la rubrique en cas de commentaire
PT2  23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT3  27/03/2006 PH V_65 Prise en compte des champs motifs absences dédoublés (Heures,Jours)
PT4  06/12/2006 SB V_70 FQ 13734 refonte gestion origine rubrique
PT5  11/01/2007 PH V_70 Correction PT4 et FQ 13808
}

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Mul, Fe_Main, DBGrids,
{$ELSE}
  MaineAgl, eMul,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, HTB97, Vierge,
  HPanel, Dialogs, Windows, PGoutils, HQry;

type
  TOF_PG_ORIGINERUB = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad; override;
  private
    CodeOrig: string;
    procedure BValClik(Sender: TObject);
  end;

implementation
uses P5Util;

// Affectation du code retour correspondant à la saisie du type de traitemant à faire

procedure TOF_PG_ORIGINERUB.BValClik(Sender: TObject);
var
  CBXO: THvalComboBox;
begin
  CBXO := THvalComboBox(GetControl('CBXORIGINE'));
  if CBXO = nil then exit;
  if CodeOrig <> 'PRO' then TFVierge(Ecran).Retour := CBXO.Value
  else TFVierge(Ecran).Retour := '';
end;

procedure TOF_PG_ORIGINERUB.OnArgument(stArgument: string);
var
  St, Rub, Rubriq, Nat: string;
  L1, L2, L3, L4, LBLO: THLabel;
  CBXO: THvalComboBox;
  BtnVal: TToolbarButton97;
  Q: TQuery;
  OkOk: Boolean;
begin
  inherited;
  // SELECT DISTINCT(PMA_RUBRIQUE) FROM MOTIFABSENCE
  st := Trim(StArgument);
  BtnVal := TToolbarButton97(GetControl('BValider'));
  if BtnVal <> nil then BtnVal.OnClick := BValClik;
  CodeOrig := ReadTokenSt(st); // Recup Origine
// DEB PT1
  Rub := ReadTokenSt(st); // Recup Rubrique
  Nat := ReadTokenSt(st); // Recup Nature
  if Nat = 'REM' then
  begin
    if RechCommentaire(Rub) = TRUE then rubriq := RendCodeRub (rub)
    else Rubriq := Rub;
    st := 'SELECT PMA_RUBRIQUE FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## (PMA_RUBRIQUE="' + Rubriq + '")'+
          ' OR (PMA_RUBRIQUEJ="' + Rubriq + '")'; // PT3
    Q := OpenSQl(st , TRUE); { PT2 }
    if (NOT Q.EOF) and (RechCommentaire(Rub) = TRUE) then
    begin
      SetControlEnabled('CBXORIGINE', FALSE);
      SetControlText('CBXORIGINE', 'BUL');
      OkOk := FALSE;
    end;
    Ferme(Q);
  end
  else okok := TRUE;
// FIN PT1
  CBXO := THvalComboBox(GetControl('CBXORIGINE'));
  if CBXO = nil then exit;
  if CodeOrig = '' then
  begin
    PGIBox('Pas de profil ?', 'Origine des Rubriques');
    exit;
  end;
  L1 := THLabel(GetControl('LBL1'));
  L2 := THLabel(GetControl('LBL2'));
  L3 := THLabel(GetControl('LBL3'));
  L4 := THLabel(GetControl('LBL4'));
  LBLO := THLabel(GetControl('LBLORIGINE')); { PT4 }
  if (L1 = nil) or (L2 = nil) or (L3 = nil) or (L4 = nil) or (LBLO = nil) then  { PT4 }
  begin
    PGIBox('Pas de contôle sur les messages', 'Origine des Rubriques');
    exit;
  end;
  L1.Caption := '';
  L2.Caption := '';
  L3.Caption := '';
  L4.Caption := '';
  if CodeOrig = 'PRO' then // Cas d'une rubrique provenant d'une profil
  begin
    L1.Caption := TraduireMemoire('La rubrique provient d''un profil salarié.');  { PT4 }
    L2.Caption := TraduireMemoire('Vous ne pouvez pas modifier l''origine de cette rubrique.');   { PT4 }
    CBXO.Enabled := FALSE;
    CBXO.Visible := FALSE;
    LBLO.Visible := FALSE;
    L3.Visible := FALSE;
    L4.Visible := FALSE;
  end
  else
  begin  { DEB PT4 }
  if (Length(Rub) > 4)  then   //cas d'une ligne de commentaire ou d'une régul
    begin
      L1.Caption := TraduireMemoire('Vous ne pouvez pas modifier l''origine de cette rubrique.');
      L2.Caption := TraduireMemoire('Cette rubrique ne sera présente que sur le bulletin en cours');
      L3.Caption := TraduireMemoire('et ne peut pas être reportée sur le bulletin suivant.');
      L4.Visible := FALSE;
// DEB PT5
      if CodeOrig = 'SAL' then CBXO.Enabled := TRUE
      else CBXO.Enabled := FALSE;
      if (RechCommentaire(Rub) = TRUE) AND (CodeOrig = 'SAL') then
      begin
        L1.Caption := TraduireMemoire('Vous pouvez modifier l''origine de cette rubrique.');
        L2.Caption := TraduireMemoire('Si vous indiquez bulletin comme origine alors la rubrique');
        L3.Caption := TraduireMemoire('ne sera pas reportée sur le bulletin suivant.');
      end;
//      CBXO.Visible := FALSE;
//      LBLO.Visible := FALSE;
    end
// FIN PT5
  else   { FIN PT4 }
  // DEB PT1
    if ((Nat <> 'REM') or (Okok)) OR (CodeOrig='BUL')  then  { PT4 }
    begin
      CBXO.Enabled := TRUE;
      L1.Caption := TraduireMemoire('Indiquez Salarié dans la zone Origine de la rubrique');
      L2.Caption := TraduireMemoire('si vous voulez que la rubrique soit conservée sur le');
      L3.Caption := TraduireMemoire('bulletin suivant, ');
      L4.Caption := TraduireMemoire('Sinon elle ne sera présente que sur ce bulletin');
    end
    else
  if CodeOrig = 'SAL' then   { PT4 }
    begin
      L1.Caption := TraduireMemoire('Indiquez Bulletin dans la zone Origine de la rubrique');   { PT4 }
      L2.Caption := TraduireMemoire('si vous voulez que la rubrique ne soit pas conservée');  { PT4 }
      L3.Caption := TraduireMemoire('sur le bulletin suivant. ');  { PT4 }
      L4.Visible := FALSE;   { PT4 }
    end;
    // FIN PT1
    CBXO.Value := CodeOrig ; //'BUL'; // PT5
  end;
end;

procedure TOF_PG_ORIGINERUB.OnLoad;
begin

end;

initialization
  registerclasses([TOF_PG_ORIGINERUB]);
end.

