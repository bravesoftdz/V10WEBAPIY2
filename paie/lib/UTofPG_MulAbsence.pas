{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 30/08/2001
Modifié le ... :   /  /
Description .. : TOF de controle du multi critere gestion des absences
Mots clefs ... : PAIE;PGABSENCES
*****************************************************************}
{
 PT1 31/08/2001 V547 PH Recherche de l'année en cours pour eviter de charger
     le liste complète de toutes les absences à chaque appel du multi critere ????
 PT2 26/09/2001 V562 SB Reconception du module des absences
 PT3 06/12/2001 V562 SB Fiche de bug n°393 Les mvts à cheval sur 2 mois ne sont pas visibles
 PT4 29/08/2005 V_65  SB FQ 11902 ajout liste de choix permettant de basculer le multicritère
                        sur liste des salariés ou des absences
 PT5 21/04/2006 V_65 SB FQ 12556 : Filtrage des absences via saisie bulletin
 PT6 09/05/2006 V_65 SB FQ 13122 Rplcmt des zones préfixées pour gestion double liste
 PT7 15/11/2006 V_70 PH Mise en place filtrage des habilitations/poupulations
 PT8 12/01/2006 V_80 FCO Rajout de la case à cocher Exclure les salariés sortis
 PT9 25/01/2007 V_80 GGU Gestion de la fiche PRESENCE_MUL
 PT10 14/02/2007 V_80 FC Mise en commentaire du spécifique Habilitations/populations
                         Création d'une vue avec jointure au lieu d'utiliser directement
                         la table ABSENCESALARIE car il y avait un BUG (AGL?)
 PT11 23/03/2007 V_70 FC FQ 13666 : différenciation d'appel entre le bulletin et la saisie par rubrique
 PT12 20/06/2007 V_72 FC FQ 14196 Accès à la saisie des absences pour le salarié confidentiel
                         si on saisit son matricule dans les critères
 PT13 21/12/2007 V_81 FC FQ 14996 Concept accessibilité fiche salarié
 }
unit UTofPG_MulAbsence;

interface
uses Controls, Classes, sysutils,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}mul,Fe_Main,
{$ELSE}
  UTOB, emul,MaineAgl,
{$ENDIF}
  HCtrls, HEnt1, UTOF, Vierge, HMsgBox,
  HQry,
  StdCtrls; //PT8

type
  TOF_PGMULABSENCE = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnClose; override;
  private
    salarie: string;
    Origine: string;
    Q_Mul: THQuery;
    GblListe, OriMenu: string;
    SsCp, Creat: Boolean; { PT5 }
    Action:String;
    procedure ActiveWhere(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure ChangeListe(Sender: TObject);
    procedure OnClickSalarieSortie(Sender: TObject);   //PT8
    procedure GrilleDblClick(Sender:TObject);//PT13
  end;

implementation

uses EntPaie, P5Def, PgOutils2;

procedure TOF_PGMULABSENCE.ActiveWhere(Sender: TObject);
var Annee, Mois, JJ: WORD;
  LaDated, Ladatef: TDateTime;
  StWhere, Pref, LePrefixe: string; // PT7
  vcbxMois, vcbxAnnee: THValComboBox;
  Where2: THEdit;
  DateArret : TdateTime;  //PT8
  StDateArret : string;   //PT8
  St : String;
begin { PT6 Refonte de la proc. }
  SetControlText('XX_WHERE', '');
  StWhere := '';
  if ((GetControl('CBLISTE')<> nil) and (GetControlText('CBLISTE') = 'ABS')) or (OriMenu = 'PRE')then //PT9
  begin
    Ladated := 0; Ladatef := 0; StWhere := '';
    DecodeDate(idate1900, Annee, Mois, JJ);
    vcbxMois := THValComboBox(GetControl('CBXMOIS'));
    vcbxAnnee := THValComboBox(GetControl('CBXANNEE'));
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

    //if ((vcbxMois <> NIL) AND (vcbxAnnee <> NIL)) then
    if Ladated > 0 then
      StWhere := '(PCN_DATEFINABS <="' + UsDateTime(Findemois(Ladatef)) + '"' +
        ' AND PCN_DATEFINABS >= "' + UsDateTime(Debutdemois(Ladated)) + '") AND '; // PT3 PCN_DATEDEBUTABS
    if  (OriMenu <> 'PRE') then  //PT9
    begin
      { DEB PT5 }
      if SSCp then
        StWhere := StWhere + '(PCN_TYPEMVT = "ABS")'
      else
        StWhere := StWhere + '(PCN_TYPEMVT = "ABS" OR PCN_TYPECONGE ="PRI")';
      { FIN PT5 }
    end;
  end;
{ DEB PT6 }
  if (GetControl('CBLISTE')<> nil) and (GetControlText('CBLISTE') = 'SAL') then Pref := ' AND PSA' else Pref := ' AND PCN'; //PT9
  if GetControlText('SALARIE') <> '' then
    StWhere := StWhere + Pref + '_SALARIE="' + GetControlText('SALARIE') + '" ';
  if GetControlText('ETABLISSEMENT') <> '' then
    StWhere := StWhere + Pref + '_ETABLISSEMENT="' + GetControlText('ETABLISSEMENT') + '" ';
  if GetControlText('TRAVAILN1') <> '' then
    StWhere := StWhere + Pref + '_TRAVAILN1="' + GetControlText('TRAVAILN1') + '" ';
  if GetControlText('TRAVAILN2') <> '' then
    StWhere := StWhere + Pref + '_TRAVAILN2="' + GetControlText('TRAVAILN2') + '" ';
  if GetControlText('TRAVAILN3') <> '' then
    StWhere := StWhere + Pref + '_TRAVAILN3="' + GetControlText('TRAVAILN3') + '" ';
  if GetControlText('TRAVAILN4') <> '' then
    StWhere := StWhere + Pref + '_TRAVAILN4="' + GetControlText('TRAVAILN4') + '" ';
{ FIN PT6 }

  //DEB PT8
  StDateArret:='';
  if  TCheckBox(GetControl('CKSORTIE')) <> nil then
  Begin
    if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET'))) then
    Begin
      DateArret   := StrtoDate(GetControlText('DATEARRET'));
      StDateArret := StDateArret + ' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
      StDateArret := StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
    End
  end;

  StWhere := StWhere + StDateArret;
  //DEB PT12
  St := SQLConf('SALARIES');
  if St <> '' then
  begin
    if StWhere <> '' then
      StWhere := StWhere + ' AND ' + St
    else
      StWhere := St;
  end;
  //FIN PT12

  //FIN PT8
  SetControlText('XX_WHERE', StWhere);

  ChangeListe(Sender);

  //DEB PT10 Mise en commentaire
  {
  // DEB PT7
  if (OriMenu = 'PAI') OR (OriMenu = 'ABS') OR (OriMenu = 'PRE') then
  begin
    if GetControlText('CBLISTE') = 'SAL' then LePrefixe := 'PSA'
    else LePrefixe := 'PCN';
    Where2 := THEdit(GetControl('XX_WHERE2'));
    if Assigned(MonHabilitation) then
      if (LePrefixe <> copy(MonHabilitation.LeSQL, 1, 3)) and (MonHabilitation.LeSQL <> '') then
      begin
        SetControlText('LEPREFIXE', LePrefixe);
        MonHabilitation.LeSQL := '';
        PGAffecteEtabByUser(TFMul(Ecran));
      end;

    if Where2 <> nil then SetControlText('XX_WHERE2', MonHabilitation.LeSQL);
    ChangeListe(Sender);
    MonHabilitation.AvecCreat := 'N';
  // FIN PT7
  end;
  }
  //FIN PT10
end;

procedure TOF_PGMULABSENCE.OnArgument(Arguments: string);
var Num: Integer;
  Defaut: THEdit;
  Arg, Etab, Mois: string;
  aa, mm, jj: word;
  ExerPerEncours, DebPer, FinPer: string;
  Q: TQuery;
  DateF: TDateTime;
  Combo: THValComboBox;
  Check : TCheckBox;  //PT8
begin
  inherited;
  Arg := Arguments;
  Creat := FALSE;
  salarie := Trim(ReadTokenPipe(Arg, ';'));
  Origine := Trim(ReadTokenPipe(Arg, ';')); //Origine :  BULLETIN, MVTABSENCE, MENU
  Etab := Trim(ReadTokenPipe(Arg, ';'));
{ DEB PT5 }
  DateF := IDate1900;
  SSCp := False;
  if Origine = 'BULLETIN' then
  begin
    DateF := StrtoDate(Trim(ReadTokenPipe(Arg, ';')));
    SSCp := True;
  end;
{ FIN PT5 }
  //DEB PT11
  if Origine = 'SAISRUB' then
    DateF := StrtoDate(Trim(ReadTokenPipe(Arg, ';')));
  //FIN PT11
  //DEB PT13
  Action:='';
  if Origine = 'SALARIES' then
    Action := Trim(ReadTokenPipe(Arg, ';'));
  if Action='ACTION=CONSULTATION' then
    SetControlProperty('BInsert','enabled',false);
  //FIN PT13
  OriMenu := Trim(ReadTokenPipe(Arg, ';')); // PT7

  setcontrolenabled('SALARIE', Origine = 'MENU'); { PT6 }
  SetControlText('SALARIE', Salarie); { PT6 }
// SetControlText('PCN_SALARIE_',Salarie); //PT2
  if Salarie = '' then Salarie := GetControlText('SALARIE'); { PT6 }

  if Origine = 'MVTABSENCE' then SetControlText('ETABLISSEMENT', Etab); { PT6 }

  for Num := 1 to 4 do
    VisibiliteChampSalarie(IntToStr(Num), GetControl('TRAVAILN' + IntToStr(Num)), GetControl('T_TRAVAILN' + IntToStr(Num))); { PT6 }
  VisibiliteStat(GetControl('CODESTAT'), GetControl('T_CODESTAT')); { PT6 }

  TFMul(Ecran).FListe.OnDblClick := GrilleDblClick; {PT2} //PT13

  Defaut := ThEdit(getcontrol('SALARIE')); { PT6 }
  if Defaut <> nil then Defaut.OnExit := ExitEdit;
{DEB PT2
Defaut:=ThEdit(getcontrol('PCN_SALARIE_'));
If Defaut<>nil then Defaut.OnExit:=ExitEdit;
FIN PT2}

//TFMul(Ecran).Binsert.OnClick:=OnNewEnr;

  if (Origine <> 'BULLETIN') and (Origine <> 'SAISRUB') then //PT2   //PT11
  begin
    if RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer) then
    begin
      DecodeDate(StrToDate(DebPer), aa, MM, JJ);
      Mois := IntToStr(MM);
      if Length(Mois) = 1 then Mois := '0' + Mois;
      SetControlText('CBXMOIS', Mois);
//  PT1 31/08/2001 V547 PH Recherche de l'année en cours pour eviter de charger
      Q := OpenSQL('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="PGA" AND CO_LIBELLE="' + IntToStr(aa) + '"', TRUE);
      if not Q.EOF then ExerPerEncours := Q.FindField('CO_CODE').AsString;
      Ferme(Q);
// FIN PT1
      SetControlText('CBXANNEE', ExerPerEncours);
    end;
  end
  else //DEB PT2
  begin
    DecodeDate(DateF, aa, MM, JJ);
    Mois := IntToStr(MM);
    if Length(Mois) = 1 then Mois := '0' + Mois;
    SetControlText('CBXMOIS', Mois);
    Q := OpenSQL('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="PGA" AND CO_LIBELLE="' + IntToStr(aa) + '"', TRUE);
    if not Q.EOF then ExerPerEncours := Q.FindField('CO_CODE').AsString;
    Ferme(Q);
    SetControlText('CBXANNEE', ExerPerEncours);
  end; //FIN PT2

  { DEB PT4 }
  if OriMenu = 'PAI' then
  begin
    Q_Mul := THQuery(Ecran.FindComponent('Q'));
    Combo := THValComboBox(GetControl('CBLISTE'));
//@@@    if Assigned(combo) then Combo.Onchange := ChangeListe;
    SetControlText('CBLISTE', 'SAL');
    ChangeListe(nil);
  end
  else if OriMenu <> 'PRE' then   //PT9
  begin
    GblListe := 'ABS';
    SetControlText('CBLISTE', 'ABS');
    SetControlVisible('CBLISTE', False);
    SetControlVisible('HCBLISTE', False);
  end;
  { FIN PT4 }

  //DEB PT8
  SetControlvisible('DATEARRET',True);
  SetControlvisible('TDATEARRET',True);
  SetControlEnabled('DATEARRET',True);
  SetControlEnabled('TDATEARRET',True);
  Check := TCheckBox(GetControl('CKSORTIE'));
  if Check = nil then
  Begin
    SetControlVisible('DATEARRET',False);
    SetControlVisible('TDATEARRET',False);
  End
  else
  begin
    Check.Checked := True;
    Check.OnClick := OnClickSalarieSortie;
  end;
  //FIN PT8
end;

procedure TOF_PGMULABSENCE.ExitEdit(Sender: TObject);
var edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGMULABSENCE.OnLoad;
var
  typeC: THEdit;
  labsal: thlabel;

begin
  inherited;
  ActiveWhere(nil);
  TypeC := THEdit(getcontrol('TYPEC'));
  Labsal := thlabel(getcontrol('T_SALARIE')); { PT6 }
  if TypeC <> nil then
  begin
    if TypeC.Text = 'S' then
    begin
      Ecran.Caption := TraduireMemoire('Consultation des absences par salarié');
      if Salarie <> '' then Ecran.Caption := TraduireMemoire('Consultation des absences de : ' + Salarie);
      if LabSal <> nil then Labsal.Caption := TraduireMemoire('Salarie');
    end;
    if TypeC.Text = 'T' then
    begin
      Ecran.Caption := TraduireMemoire('Consultation des absences par type'); //SB
      if LabSal <> nil then Labsal.Caption := TraduireMemoire('Salarié de');
    end;
    SetControlVisible('NOM2', TypeC.Text = 'T');
    SetControlVisible('LNOM2', TypeC.Text = 'T');
  end;
  if OriMenu = 'PRE' then   //PT9
    Ecran.Caption := TraduireMemoire('Saisie d''évènement');

  UpdateCaption(TFVierge(Ecran));
end;
{ DEB PT4 }

procedure TOF_PGMULABSENCE.ChangeListe(Sender: TObject);
begin
  if not assigned(GetControl('CBLISTE')) then exit;
  if GblListe = GetControlText('CBLISTE') then exit;

  if GetControlText('CBLISTE') = 'SAL' then
  begin
    TFMul(Ecran).SetDBListe('PGMULSALARIE');
//  if Q_Mul <> nil then Q_Mul.Liste := 'PGMULSALARIE';
    SetControlText('PCN_TYPECONGE', '');
    SetControlEnabled('PCN_TYPECONGE', False);
    SetControlEnabled('CBXMOIS', False);
    SetControlEnabled('CBXANNEE', False);
  end
  else
    if GetControlText('CBLISTE') = 'ABS' then
    begin
      // Ancienne liste PGMULMVTABSENCE maintenant on passe par une vue //PT10
      //TFMul(Ecran).SetDBListe('PGMULMVTABSENCE'); //PT10
      TFMul(Ecran).SetDBListe('PGMULABSENCESALAR');
//  if Q_Mul <> nil then Q_Mul.Liste := 'PGMULMVTABSENCE';
      SetControlEnabled('PCN_TYPECONGE', True);
      SetControlEnabled('CBXMOIS', True);
      SetControlEnabled('CBXANNEE', True);
    end;
  GblListe := GetControlText('CBLISTE');
  TFMul(Ecran).BCherche.Click;
end;
  { FIN PT4 }

procedure TOF_PGMULABSENCE.OnClose;
begin
  inherited;
  MonHabilitation.AvecCreat := '';
end;

procedure TOF_PGMULABSENCE.OnClickSalarieSortie(Sender: TObject);
begin
  //DEB PT8
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
  //FIN PT8
end;

//DEB PT13
procedure TOF_PGMULABSENCE.GrilleDblClick(Sender: TObject);
begin
  if TFmul(Ecran).Q.RecordCount <> 0 then
//  if (GetField('PCN_SALARIE') <> '') then
  begin
    if GetControlText('CBLISTE') = 'ABS' then
    begin
      if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
        if Action='ACTION=CONSULTATION' then
          AGLLanceFiche  ('PAY','MVTABSENCE','',GetField('PCN_TYPEMVT')+';'+GetField('PCN_SALARIE')+';'+IntToStr(GetField('PCN_ORDRE')),GetField('PCN_SALARIE') +';A;'+GetField('PCN_ETABLISSEMENT')+';ACTION=CONSULTATION;MONOFICHE')
        else
          AGLLanceFiche  ('PAY','MVTABSENCE','',GetField('PCN_TYPEMVT')+';'+GetField('PCN_SALARIE')+';'+IntToStr(GetField('PCN_ORDRE')),GetField('PCN_SALARIE') +';A;'+GetField('PCN_ETABLISSEMENT')+';ACTION=MODIFICATION;MONOFICHE')
      else
        if Action='ACTION=CONSULTATION' then
          AGLLanceFiche  ('PAY','MVTABSENCE','',GetField('PCN_TYPEMVT')+';'+GetField('PCN_SALARIE')+';'+IntToStr(GetField('PCN_ORDRE')),GetField('PCN_SALARIE') +';A;'+GetField('PCN_ETABLISSEMENT')+';ACTION=CONSULTATION')
        else
          AGLLanceFiche  ('PAY','MVTABSENCE','',GetField('PCN_TYPEMVT')+';'+GetField('PCN_SALARIE')+';'+IntToStr(GetField('PCN_ORDRE')),GetField('PCN_SALARIE') +';A;'+GetField('PCN_ETABLISSEMENT')+';ACTION=MODIFICATION');
    end
    else if GetControlText('CBLISTE') = 'SAL' then
      if Action<>'ACTION=CONSULTATION' then
      begin
        if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
          AGLLanceFiche ('PAY','MVTABSENCE','',GetField('PSA_SALARIE'), GetField('PSA_SALARIE') +';A;'+GetField('PSA_ETABLISSEMENT')+';ACTION=CREATION;MONOFICHE')
        else
          AGLLanceFiche ('PAY','MVTABSENCE','',GetField('PSA_SALARIE'), GetField('PSA_SALARIE') +';A;'+GetField('PSA_ETABLISSEMENT')+';ACTION=CREATION');
      end;
  end;
end;
//FIN PT13

initialization
  registerclasses([TOF_PGMULABSENCE]);
end.

