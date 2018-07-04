{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 10/02/2003
Modifié le ... :   /  /
Description .. : Gestion du multicritères de saisie groupée des CP
Mots clefs ... : PAIE;CP
*****************************************************************}
{
PT1 10/02/2003 SB V595 FQ 10485 Collezerodevant zone salarié
PT2 14/05/2004 SB V_50 FQ 11167 Mise en place Multisélection
PT3 10/08/2004 SB V_50 FQ 11385 Remplacemnt THValCombo par ThMultiValcombo
PT4 20/10/2004 PH V_50 FQ 11710 Multiple sélection avec rafraichissement de la liste
PT5 24/01/2004 SB V_65 FQ 11624 Intégration saisie groupée des absences
PT6 05/10/2005 SB V_65 FQ 12624 Condition gestion Cp que si saisie groupée CP
PT7 22/11/2005 SB V_65 FQ 12341 Ajout critère date sortie arrêté le
PT8 13/04/2006 SB V_65 FQ 12924 Reprise PT6
PT9 11/05/2006 SB V_65 FQ 13125 Affichage message de confirmation de traitement
PT10 09/01/2007 FCO V_80 FQ 13548 Autoriser la saisie des CP même pour les établissements sans CP
PT11 09/10/2008 JPP FQ 15627 Interdire la saisie de congès groupés si aucun salarié n'est selectionné
PT12 09/10/2008 JPP FQ 15627 Améliorations, débuggage dans la sélection des salariés
}
unit UTOFCongesGrpMul;

interface
uses Classes, Sysutils, Controls, AglInit,StdCtrls,
{$IFNDEF EAGLCLIENT}
  mul, FE_Main,
{$IFNDEF DBXPRESS}
     dbTables,
{$ELSE}
     uDbxDataSet,
{$ENDIF}
{$ELSE}
  emul, MaineAgl,
{$ENDIF}
  Utof, hctrls, HTB97, HMsgBox, HQry, Hent1,Utob;

type
  TOF_CongesGrpMul = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad ;                        override; { PT7 }
  private
    StMode,StCaption : String; { PT5 }
    procedure ExitEdit(Sender: TObject); //PT1
    procedure ClickBOuvrir(Sender: TObject); { PT2 }
    procedure OnClickSalarieSortie(Sender: TObject);  { PT7 }
  end;

var
  CritEtab : String;     //PT10


implementation
uses Entpaie, P5Def,  PGOutils2, DB;


procedure TOF_CongesGrpMul.OnArgument(stArgument: string);
var
  ChampLieu: THEdit;
  Num: integer;
  Btn: TToolBarButton97;
  Check : TCheckBox;
begin
  inherited;
  StMode := stArgument;   { DEB PT5 }
  if StMode = 'CP' then
    Begin
    StCaption := 'Saisie groupée des congés payés';
    SetControlText('XX_WHEREZZ', 'AND PSA_CONGESPAYES = "X" AND PSA_ETABLISSEMENT IN (SELECT ETB_ETABLISSEMENT FROM ETABCOMPL ' +
      'WHERE ETB_CONGESPAYES = "X")');
    End
  else
    Begin
    StCaption := 'Saisie groupée des absences';
    SetControlProperty('PSA_ETABLISSEMENT','DataType','TTETABLISSEMENT'); { PT6 }
    SetControlProperty('PSA_SALARIE','DataType','PGSALPREPAUTO'); {PT6 07/11/2005}
    End;
  Ecran.Caption := StCaption;
  UpdateCaption(Ecran);   { FIN PT5 }

  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));
  { DEB PT1 }
  ChampLieu := THEdit(GetControl('PSA_SALARIE'));
  if ChampLieu <> nil then ChampLieu.OnExit := ExitEdit;
  ChampLieu := THEdit(GetControl('PSA_SALARIE_'));
  if ChampLieu <> nil then ChampLieu.OnExit := ExitEdit;
  { FIN PT1 }

  { DEB PT2 }
  Btn := TToolBarButton97(GetControl('BOuvrir'));
  if Btn <> nil then Btn.Onclick := ClickBOuvrir;
  { FIN PT2 }

  { DEB PT7 }
   Check:=TCheckBox(GetControl('CKSORTIE'));
 if Check=nil then
   Begin
    SetControlVisible('DATEARRET',False);
    SetControlVisible('TDATEARRET',False);
    End
  else
    Check.OnClick:=OnClickSalarieSortie;
  { FIN  PT7 }

end;

{ DEB PT1 }

procedure TOF_CongesGrpMul.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;
{ FIN PT1 }


{ DEB PT2 }

procedure TOF_CongesGrpMul.ClickBOuvrir(Sender: TObject);
var
  Salarie, st : string;
  i: integer;
  Q_Mul: THQuery;
  LeWhere,Etab : String; //PT10
  Q : TQuery; //PT10
  T : TOB; //PT10
  bTous : Boolean; //PT10
begin
  CritEtab := GetControlText('PSA_ETABLISSEMENT'); //PT10
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then exit;
  bTous := False; //PT10
  if TFMul(Ecran).FListe.AllSelected then bTous := True; //PT10
  { Gestion de la sélection de salarié }
  if (TFMul(Ecran).FListe.nbSelected > 0) and (not TFMul(Ecran).FListe.AllSelected) then
    Begin
    if PgiAsk('Vous avez selectionné des salariés. Voulez-vous générer des absences que pour ces salariés?', Ecran.caption) = mrYes then { PT5 }
    begin
      St := '';
      { Suppression CP pour salarié sélectionné }
      { Composition du clause WHERE pour limiter le mul à ces salariés }
      for i := 0 to TFMul(Ecran).FListe.NbSelected - 1 do
      begin
{$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
        TFMul(Ecran).FListe.GotoLeBookmark(i);
        Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
        St := St + ' PSA_SALARIE="' + Salarie + '" OR';
      end;
      TFMul(Ecran).FListe.ClearSelected;
      if St <> '' then
        SetControlText('XX_WHEREZZ', GetControlText('XX_WHEREZZ') + ' AND (' + Copy(St, 1, Length(st) - 2) + ')'); // PT4 { PT6 } PT12

      if not (TFMul(Ecran).FListe.AllSelected) then //PT10
        TFMul(Ecran).BCherche.Click;
    end
    else
      exit; //PT10
   end
    else if (TFMul(Ecran).FListe.nbSelected = 0) and (not TFMul(Ecran).FListe.AllSelected) then // PT12
      // DEB PT11
      begin
        PgiInfo('Vous devez préalablement sélectionner les salariés', Ecran.caption);
        exit;
      end;
      // FIN PT11

  // AVANT PT11
  if PgiAsk('Voulez-vous générer des absences pour tous les salariés affichés?', Ecran.caption) = mrno then exit;

  { Récupération de la Query pour traitement dans la fiche vierge }
{$IFDEF EAGLCLIENT}
  if TFMul(Ecran).Fetchlestous then TheMulQ := TOB(Ecran.FindComponent('Q'));
{$ELSE}
  TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}

  //DEB PT10
  if bTous then
  begin
    LeWhere := RecupWhereCritere(TFMul(Ecran).Pages);
    St := 'SELECT DISTINCT PSA_ETABLISSEMENT FROM SALARIES ' + LeWhere;
    T := Tob.Create('lesetab',nil,-1);
    T.LoadDetailDBFromSQL('SALARIES',St);
    CritEtab := '';
    For i := 0 to T.detail.Count - 1 do
    begin
      if CritEtab = '' then
        CritEtab := T.detail[i].GetValue('PSA_ETABLISSEMENT')
      else
        CritEtab := CritEtab + ';' + T.detail[i].GetValue('PSA_ETABLISSEMENT');
    end;
    FreeAndNil(T);
  end
  else if (CritEtab = '') or (CritEtab = '<<Tous>>') then
  begin
    LeWhere := RecupWhereCritere(TFMul(Ecran).Pages);
    Q := Opensql('SELECT DISTINCT PSA_ETABLISSEMENT FROM SALARIES ' + LeWhere,True);
    CritEtab := '';
    While Not Q.Eof do
    begin
      if CritEtab = '' then
        CritEtab := Q.Findfield('PSA_ETABLISSEMENT').AsString
      else
        CritEtab := CritEtab + ';' + Q.Findfield('PSA_ETABLISSEMENT').AsString;
      Q.Next;
    end;
    Ferme(Q);
  end;
  //FIN PT10

  { Ouverture de la fiche }
  AglLanceFiche('PAY', 'CONGESGRP', '', '', StMode);   { PT5 }
  TheMulQ := nil;
  if StMode = 'CP' then              { DEB PT6 }
    SetControlText('XX_WHEREZZ', 'AND PSA_CONGESPAYES = "X" AND PSA_ETABLISSEMENT IN (SELECT ETB_ETABLISSEMENT FROM ETABCOMPL ' +
      'WHERE ETB_CONGESPAYES = "X")')
 else
   SetControlText('XX_WHEREZZ', '');  { FIN PT6 }
 TFMul(Ecran).BCherche.Click; //PT10
end;
{ FIN PT2 }

{ DEB PT7 }
procedure TOF_CongesGrpMul.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;
{ FIN PT7 }

{ DEB PT7 }
procedure TOF_CongesGrpMul.OnLoad;
var
  DateArret : TDateTime;
  St        : String;
begin
{ PT8 Mise en commentaire
if StMode = 'CP' then
    SetControlText('XX_WHEREZZ', 'AND PSA_CONGESPAYES = "X" AND PSA_ETABLISSEMENT IN (SELECT ETB_ETABLISSEMENT FROM ETABCOMPL ' +
      'WHERE ETB_CONGESPAYES = "X")');
else
    SetControlText('XX_WHEREZZ', '');}

if  TCheckBox(GetControl('CKSORTIE'))<>nil then
  Begin
  if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then   //DEB PT2
     Begin
     DateArret:=StrtoDate(GetControlText('DATEARRET'));
     st:=' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'"'+
         ' OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'"'+
         ' OR PSA_DATESORTIE IS NULL)'+
         ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'" ';
     SetControlText('XX_WHEREZZ',GetControlText('XX_WHEREZZ') + st);
     End;
  End;
end;
{ FIN PT7 }

initialization
  registerclasses([TOF_CongesGrpMul]);
end.

