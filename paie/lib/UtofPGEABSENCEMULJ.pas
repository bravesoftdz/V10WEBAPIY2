{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 22/06/2005
Modifié le ... :   /  /
Description .. : Multi critère listes des absences du jour avec la hiérarachie
Suite ........ :
Suite ........ :
Mots clefs ... : PAIE;ABSENCES;PGDEPORTEE
*****************************************************************
}
unit UtofPgEAbsenceMulJ;

interface
uses Controls, Classes, sysutils, HTB97,
{$IFDEF EAGLCLIENT}
  UtileAGL, eMul, MaineAgl, HStatus, Utob,
{$ELSE}
  db, dbTables, HDB, FE_Main, Mul, HStatus,
{$ENDIF}
  HCtrls, HQry, HEnt1, HMsgBox, UTOF;

type
  tof_PgEAbsenceMulJ = class(TOF)
  private
    Ficprec, Stperiode, StSal: string;
    WW: THEdit;
    QMul: TQUERY; // Query recuperee du mul
    vcbxMois, vcbxAnnee: THValComboBox;
    LaDated, Ladatef: TDateTime;
    GblNbUpdate: integer;
    procedure ActiveWhere(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure btValidclick(sender: tobject);
    procedure btRecapclick(sender: tobject);
    procedure MajAbsenceExport (Mode : String);
    procedure traiteliste(mode: string);
    procedure enableChampOngletModeAgl;
    procedure btPlanningclick(sender: tobject);
    procedure GrilleDblClick(Sender: TObject);
    procedure OnExitSalarie(Sender: TObject);
    procedure ElipsisClickSal(Sender: TObject);
  end;

implementation
uses EntPaie, P5Util, P5Def, PGoutils2, PgOutilseAgl, PgPlanning, PgPlanningOutils;

{ La confection de la clause where depend du parametre et donc aussi de la fiche
  utilisée. Les listes sont différentes
  3 cas : Salarié        ==> absences du salariés
          Responsable    ==> absences des salariés qui ont comme responsable
          Administrateur ==> accès à toutes les absences
}

procedure tof_PgEAbsenceMulJ.ActiveWhere(Sender: TObject);
var Annee, Mois, JJ: WORD;
begin
// initilalisation du code salarié en fonction de la connection
  if WW = nil then exit;
  WW.Text := '';
  LaDated := idate1900; LaDateF := idate1900;
  if Ficprec = 'SAL' then
  begin // Cas absence du salarie
    if LeSalarie <> '' then WW.text := ' PCN_SALARIE = "' + LeSalarie + '" '
    else // Anomalie donc ne rien afficher dans le Mul donc requete impossible
      WW.text := ' PCN_SALARIE = "0000000000" AND PCN_SALARIE <> "0000000000" ';
  end
  else
  begin
    if Ficprec = 'RESP' then
    begin // Cas des Absences à valider pour le responsable
      if LeSalarie <> '' then
      begin
      WW.text := ' (PSE_RESPONSABS = "' + LeSalarie + '"'+
       'OR PSE_RESPONSABS IN (SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
       'WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSABS="'+LeSalarie+'")))';
      end
      else WW.text := ' PSE_RESPONSABS = "0000000000" AND PSE_RESPONSABS <> "0000000000"';
    end
    else // Cas administrateur doit tout voir sinon erreur ne voit rien
      if Ficprec <> 'ADM' then WW.text := ' PCN_SALARIE = "0000000000" AND PCN_SALARIE <> "0000000000" ';
  end;
  StPeriode := '';
  DecodeDate(Idate1900, Annee, Mois, JJ);
  if vcbxMois.Value <> '' then Mois := Trunc(StrToInt(vcbxMois.Value));
  if vcbxAnnee.Value <> '' then Annee := Trunc(StrToInt(vcbxAnnee.Text));
  if ((vcbxMois.Value <> '') and ((VcbxMois.value <> '') and (vcbxAnnee.value <> ''))) then // cas où mois renseigné alors année obligatoire
  begin
    LaDated := EncodeDate(annee, mois, 1);
    LaDatef := EncodeDate(annee, mois, 1);
  end
  else if vcbxAnnee.Value <> '' then
  begin
    LaDatef := EncodeDate(annee, 12, 1);
    LaDated := EncodeDate(annee, 1, 1);
  end;
  if ((vcbxMois <> nil) and (WW <> nil) and (vcbxAnnee <> nil)) then
    if Ladated > idate1900 then
    begin
      StPeriode := ' AND (PCN_DATEFINABS <="' + UsDateTime(Findemois(Ladatef)) + '"' +
        ' AND PCN_DATEFINABS >= "' + UsDateTime(Debutdemois(Ladated)) + '")'; //PT-3 modif champ PCN_DATEDEBUTABS
      WW.Text := WW.Text + StPeriode;
    end;
  if ((Ficprec = 'ADM') and (GetControlText('CKSORTIE') = '-')) or (Ficprec <> 'ADM') then
  begin
    StSal := ' AND (PSA_DATESORTIE<="' + UsDateTime(idate1900) + '" ' +
      'OR PSA_DATESORTIE is null ' +
      'OR PSA_DATESORTIE>="' + UsDatetime(Vh_Paie.PGECabDateIntegration) + '")';
    WW.text := WW.text + StSal;
  end
  else
    StSal := '';
end;

procedure tof_PgEAbsenceMulJ.OnArgument(Arguments: string);
var
  Edit, FP: thedit;
  Bt: ttoolbarbutton97;
  Zone: Tcontrol;
  num: integer;
  Grille: THGrid;
begin
  inherited;
  Ficprec := Arguments;
  WW := THEdit(GetControl('XX_WHERE'));
  bt := ttoolbarbutton97(getcontrol('TVALID'));
  if bt <> nil then bt.onclick := btValidclick;
  bt := ttoolbarbutton97(getcontrol('TPLANNING'));
  if bt <> nil then
  begin
    if Ficprec <> 'SAL' then Bt.Visible := TRue;
    bt.onclick := btPlanningclick;
  end;
  FP := THedit(getcontrol('FICPREC'));
  if FP <> nil then FP.text := Ficprec;
  Grille := THGrid(getcontrol('FLISTE'));
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;
  if ((FicPrec = 'ADM') or (FicPrec = 'RESP')) and (GetControl('PSA_TRAVAILN1') <> nil) then
  begin
    SetControlProperty('Pavance', 'TabVisible', True);
    SetControlProperty('PComplement', 'TabVisible', VH_PAIE.PGEcabHierarchie);
    SetControlVisible('TPSE_CODESERVICE', VH_PAIE.PGEcabHierarchie);
    SetControlVisible('PSE_CODESERVICE', VH_PAIE.PGEcabHierarchie);
    for Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
      SetControlProperty('PComplement', 'TabVisible', True);
      if Num > 4 then Break;
      VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
    end;
    for Num := 1 to VH_Paie.PgNbCombo do
    begin
      SetControlProperty('TBChamplibre', 'TabVisible', True);
      if Num > 4 then Break;
      VisibiliteChampLibreSal(IntToStr(Num), GetControl('PSA_LIBREPCMB' + IntToStr(Num)), GetControl('TPSA_LIBREPCMB' + IntToStr(Num)));
    end;
  end;

  enableChampOngletModeAgl;
  Edit := THedit(getcontrol('PCN_SALARIE'));
  if Edit <> nil then
  begin Edit.OnExit := OnExitSalarie;
    if FicPrec <> 'ADM' then Edit.OnElipsisClick := ElipsisClickSal;
  end;
  SetControlvisible('CKSORTIE', (FicPrec = 'ADM'));
end;

procedure tof_PgEAbsenceMulJ.btValidclick(sender: tobject);
begin
  traiteListe('VAL');
end;

procedure tof_PgEAbsenceMulJ.btPlanningclick(sender: tobject);
var
  i: integer;
  ListeSal, St, StFiche, StParam, StListe: string;
  DD, FD: TDateTime;
  NiveauRupt: TNiveauRupture;
begin
  ListeSal := ''; StFiche := ''; StParam := ''; StListe := '';
  if (LeSalarie <> '') then
  begin
    if Ficprec = 'SAL' then
      StFiche := 'AND PSE_SALARIE = "' + LeSalarie + '" '
    else
      if Ficprec = 'RESP' then
        StFiche := 'AND PSE_RESPONSABS = "' + LeSalarie + '" ';
  end;

  if GetControlText('PCN_SALARIE') <> '' then StParam := 'OR PSE_SALARIE="' + GetControlText('PCN_SALARIE') + '" ';
  if GetControlText('PSA_LIBELLE') <> '' then StParam := StParam + 'OR PSA_LIBELLE LIKE "' + GetControlText('PSA_LIBELLE') + '%" ';
  if StParam <> '' then StParam := 'AND (' + Copy(StParam, 3, Length(StParam)) + ')';
  St := StParam;
  for i := 0 to TFmul(Ecran).Fliste.NbSelected - 1 do
  begin
    TFmul(Ecran).Fliste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
    TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
    if Pos(TFmul(Ecran).Q.FindField('PCN_SALARIE').asstring, StListe) = 0 then
      StListe := StListe + '"' + TFmul(Ecran).Q.FindField('PCN_SALARIE').asstring + '",';
  end;
  if StListe <> '' then
    St := 'AND PSE_SALARIE IN (' + Copy(StListe, 1, Length(StListe) - 1) + ')';
  St := StFiche + St;
  ListeSal := 'WHERE (PSA_DATESORTIE<="' + UsDateTime(idate1900) + '" ' +
    'OR PSA_DATESORTIE is null ' +
    'OR PSA_DATESORTIE>="' + UsDatetime(Vh_Paie.PGECabDateIntegration) + '")';
  if (St <> '') then ListeSal := ListeSal + ' AND ' + Copy(St, 4, Length(St));
  if (LaDated <= IDate1900) and (VH_Paie.PGECabDateIntegration <> idate1900) then //DEB PT-5
  begin
    DD := DebutDeMois(VH_Paie.PGECabDateIntegration);
    FD := FinDeMois(VH_Paie.PGECabDateIntegration);
  end
  else
  begin
    DD := LaDated;
    FD := FindeMois(DD);
  end;
  NiveauRupt.ChampsRupt[1] := ''; NiveauRupt.ChampsRupt[2] := '';
  NiveauRupt.ChampsRupt[3] := ''; NiveauRupt.ChampsRupt[4] := '';
  NiveauRupt.CondRupt := ''; NiveauRupt.NiveauRupt := 0;
  PGPlanningAbsence(DD, FD, FicPrec, '', '', ListeSal, LeSalarie, NiveauRupt);
end;

procedure tof_PgEAbsenceMulJ.btRecapclick(sender: tobject);
begin
  AgllanceFiche('PAY', 'MULRECAPSAL', '', '', Ficprec);
end;
// fonction de validation en bloc des absences
// On utilise la multi sélection, soit pour tout valider en bloc
// Vu avec JLD pour optimisation pour le pb de validation par lot

procedure tof_PgEAbsenceMulJ.traiteliste(mode: string);
var I: integer;
begin
  if TFMul(Ecran).FListe.NbSelected < 1 then
  begin
    PGIBox('Vous devez séléctionner les lignes d''absences en attente à valider.', Ecran.caption);
    exit;
  end;
  GblNbUpdate := 0;
  if PgiAsk('Confirmez-vous la validation des absences ?', Ecran.Caption) <> mrYes then exit; { PT-14 }
  with TFMul(Ecran) do
  begin
    if TFMul(Ecran).FListe.AllSelected then
    begin
      InitMove(Q.RecordCount, '');
      Q.First;
      while not Q.EOF do
      begin
        MoveCur(False);
        MajAbsenceExport(Mode);
        Q.NEXT;
      end;
      TFMul(Ecran).FListe.AllSelected := False;
    end else
    begin
      InitMove(TFMul(Ecran).FListe.NbSelected, '');
      for i := 0 to TFMul(Ecran).FListe.NbSelected - 1 do
      begin
        TFMul(Ecran).FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
        MajAbsenceExport(Mode);
        MoveCur(False);
      end;
      TFMul(Ecran).FListe.ClearSelected;
    end;
    FiniMove;
    if GblNbUpdate > 0 then PgiInfo(IntToStr(GblNbUpdate) + ' mouvement(s) d''absence validé(s).Aucune anomalie n''a été détectée.', Ecran.Caption); { PT-15-2 }
  end;
  TFMUL(Ecran).BChercheClick(nil);
  if (Ficprec = 'RESP') then
    ExecuteSql('UPDATE DEPORTSAL SET PSE_EMAILENVOYE="-" WHERE PSE_SALARIE="' + LeSalarie + '"');
end;
// Fonction qui rend invisible les boutons en focntion du type de traitement

procedure tof_PgEAbsenceMulJ.enableChampOngletModeAgl;
begin
  SetcontrolEnabled('TREFUS', (Ficprec <> 'ADM'));
  SetControlEnabled('BOUVRIR', (Ficprec <> 'ADM'));
end;

procedure tof_PgEAbsenceMulJ.OnLoad;
var Q: TQuery;
  LeNom, st: string;
  Nbre: Integer;
begin
  inherited;
  ActiveWhere(nil);
  LeNom := RechDom('PGSALARIE', LeSalarie, False);
  if FicPrec = 'SAL' then Ecran.Caption := 'Gestion de mes absences : ' + LeNom {+' '+LePrenom}
  else if FicPrec = 'RESP' then Ecran.Caption := LeNom {+' '+LePrenom } + ' : gestion des absences de mes collaborateurs'
  else if FicPrec = 'ADM' then Ecran.Caption := 'Supervision des absences déportées par l''administrateur';
// Test en fonction du type de traitement de l'affichage des controls
  setcontrolVisible('LSALARIE', (Ficprec = 'RESP') or (Ficprec = 'ADM'));
  setcontrolVisible('PCN_SALARIE', (Ficprec = 'RESP') or (Ficprec = 'ADM'));
  setcontrolVisible('LVALIDRESP', (Ficprec = 'RESP') or (Ficprec = 'ADM'));
  setcontrolVisible('TREFUS', (Ficprec = 'RESP'));
  if Ficprec = 'RESP' then
  begin
    st := 'SELECT COUNT (*) NOMBRE FROM ABSENCESALARIE LEFT JOIN DEPORTSAL ON PSE_SALARIE=PCN_SALARIE ' +
      'LEFT JOIN SALARIES ON PCN_SALARIE=PSA_SALARIE ' +
      'WHERE ((PCN_TYPEMVT="ABS" AND PCN_SENSABS="-")OR PCN_TYPECONGE="PRI") AND PCN_VALIDRESP = "ATT"  ' +
      'AND PSE_RESPONSABS="' + LeSalarie + '" ' + StPeriode + StSal;
    q := Opensql(St, true);
    if not Q.EOF then
    begin
      Nbre := Q.FindField('NOMBRE').AsInteger;
      setcontrolvisible('AVERT', True);
      setcontroltext('AVERT', 'Vous avez ' + inttostr(Nbre) + ' mouvement(s) à valider');
    end;
    Ferme(Q);
  end;
  UpdateCaption(TFmul(Ecran));
end;

// gestion du dblclick attention dans le cas de l'administrateur, ON PROPOSE une
// intervention identique au responsable

procedure tof_PgEAbsenceMulJ.GrilleDblClick(Sender: TObject);
var Ordre, sal, etab, typemvt, lib, StAction, ValidResp, exportok: string;
begin
{$IFDEF EAGLCLIENT}
  QMUL := TOB(TFMul(Ecran).Q.TQ);
{$ELSE}
  QMUL := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}
  if QMUL.EOF then exit // Si grille vide alors pas de création possible
  else
  begin
{$IFDEF EAGLCLIENT}
    TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
    if QMUL.Eof = true then begin exit; end;
    Ordre := inttostr(QMUL.findfield('PCN_ORDRE').Asinteger);
    Typemvt := QMUL.findfield('PCN_TYPEMVT').Asstring;
    Sal := QMUL.findfield('PCN_SALARIE').Asstring;
    Lib := QMUL.findfield('PSA_LIBELLE').Asstring;
    Etab := QMUL.findfield('PCN_ETABLISSEMENT').Asstring;
    exportok := QMUL.findfield('PCN_EXPORTOK').Asstring;
    ValidResp := QMUL.findfield('PCN_VALIDRESP').Asstring;
    if Ficprec = 'SAL' then
    begin
      if ValidResp = 'ATT' then StAction := 'ACTION=MODIFICATION' else StAction := 'ACTION=CONSULTATION';
      AglLanceFiche('PAY', 'EABSENCE', 'PCN_SALARIE=' + sal, TYPEMVT + ';' + sal + ';' + Ordre, sal + ';E;' + etab + ';' + StAction + ';;' + Ficprec)
    end
    else
    begin
      if ExportOk = '-' then StAction := 'ACTION=MODIFICATION' else StAction := 'ACTION=CONSULTATION';
      if IfMotifabsenceSaisissable(QMUL.findfield('PCN_TYPECONGE').Asstring, Ficprec) then StAction := 'ACTION=MODIFICATION'; //PT12
      if (Ficprec = 'RESP') or (Ficprec = 'ADM') then
        AglLanceFiche('PAY', 'EABSENCE', 'PCN_SALARIE=' + sal, TYPEMVT + ';' + sal + ';' + Ordre, sal + '!' + lib + ';E;' + etab + ';' + StAction + ';;' + Ficprec);
    end;
  end;
  TFMUL(Ecran).BChercheClick(Sender);
end;
// gestion dE INSERT attention dans le cas de l'administrateur, il n'a pas le droit de créer
// des enregistrements, il le fera dans la paie mais dans la base de production de la paie

procedure tof_PgEAbsenceMulJ.OnExitSalarie(Sender: TObject);
var edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure tof_PgEAbsenceMulJ.ElipsisClickSal(Sender: TObject);
begin
  AfficheTabSalResp(Sender, LeSalarie);
end;

procedure tof_PgEAbsenceMulJ.MajAbsenceExport (Mode : String) ;
var st,Sal,TypMvt    : String;
    ordre            : Integer;
begin
if TFmul(Ecran).Q.findfield('PCN_VALIDRESP').AsString <>'ATT' then exit; { PT-15-2 }
TypMvt := TFmul(Ecran).Q.findfield('PCN_TYPEMVT').asstring;
Sal    := TFmul(Ecran).Q.findfield('PCN_SALARIE').asstring;
Ordre  :=TFmul(Ecran).Q.findfield('PCN_ORDRE').asinteger;
st := 'UPDATE ABSENCESALARIE SET PCN_VALIDRESP = "'+Mode+'",PCN_VALIDABSENCE="'+LeSalarie+'" '+
      'WHERE PCN_TYPEMVT = "'+TypMvt+'" AND PCN_SALARIE ="'+Sal+'" '+
      'AND PCN_ORDRE ='+IntToStr(Ordre);
GblNbUpdate := GblNbUpdate + ExecuteSql (st);   { PT-15-2 }
//PT-10 Recalcul des compteurs en cours lors de la validation
CalculRecapAbsEnCours(Sal);
end;


initialization
  registerclasses([tof_PgEAbsenceMulJ]);
end.

