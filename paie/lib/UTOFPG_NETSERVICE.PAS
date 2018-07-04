{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 04/09/2001
Modifié le ... :   /  /
Description .. : Unit de confection du fichier envoyé pour la saisie déportée
Suite ........ : des éléments de la paie. Cela concerne les rubriques
Suite ........ : saisisables et les congés payés.
Mots clefs ... : PAIE;ABSENCES;PGDEPORTEE
*****************************************************************}
{ Objectif : créer un fichier comportant tous les enregistrements nécessaires à
la saisie déportée des élements d'un paie sur internet.
On fournit donc tous les renseignements concernant :
- la société : paramsoc
- les salaries
- l'historique du bulletin
- les recap salariés des CP et RTT
- certains cumuls salariés pour infos
- les définitions de cumuls
- les motifs d'absences gérés en saisie NETSERVICE = Saisie déportée des absences
On traite la liste des salariés présents et en non suspension de paie.
Par contre, on prend tous les historiques sur la période demandée.
Le traitement est multi établissement cad au niveau du dossier contrairement à ISIS2
}
{
PT1   : 15/10/2001 PH V562 Export des motifs d'absences
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT2   : 10/09/2004 PH V_50 Gestion des transmissions au serveur CWAS NetExpert
PT3   : 16/11/2004 PH V_60 Exportation de tous les champs idem importation avec
                           création de salarié
PT3   : 04/01/2005 PH V_60 Exportation du numéro de dossier dans entête
PT4   : 18/03/2005 PH V_60 On met le répertoire de travail de la pâie pour la
                           génération du fichier
PT5   : 08/06/2005 PH V_60 FQ 12329 Message alerte Exercice social non existant
PT6   : 08/06/2005 PH V_60 FQ 12370 Rajout du nombre de jours travaillés sur
                           chaque enregistrement établissement
PT7   : 13/07/2005 PH V_60 FQ 12276 Prise en compte sélection  établissement
PT8   : 21/09/2005 PH V_60 FQ 12581 Suppression de l'envoi des lignes de
                           histobulletin
PT9   : 05/10/2005 PH V_60 Ergonomie titre du caption dans la fiche
PT10  : 13/06/2006 PH V_65 Nouveaux champs en extraction et selection des motifs
                           et des masques
PT11  : ???????
PT12  : 08/01/2007 GG V_70 user n'ayant pas le droit de voir les confidentiels :
                           dans ce cas, ne pas envoyer de salariés confidentiels
                           dans le fichier
PT13  : 24/04/2007 PH V_70 Nouveautés version 8
PT14  : 11/06/2007 VG V_72 Adaptation nouvelles méthodes CBPPath
PT15  : 20/09/2007 VG V_80 FQ 14781
PT16  : 12/10/2007 PH V_80 FQ 14834 Remis tauxpartiel à la place de prix dans effectif
PT17  : 08/01/2008 PH V_80 Contrôle du contenu du masque pour éviter l'oubli de colonne saisissable
PT18  : 25/02/2008 PH V_80 FQ 15256 Mauvais enreg concernant le masque de saisie par rubrique
}
unit UTofPG_NETSERVICE;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HTB97,
{$IFDEF EAGLCLIENT}
  UtileAGL,
{$ELSE}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  UTOB,
  ParamSoc,
  MailOL,
  Vierge,
  PGRepertoire;

type
  TOF_PG_NETSERVICE = class(TOF)
  public
    BtnLance: TToolbarButton97;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    Destinataire, LeFichier: string;
    MonMemo: TMemo;
    PGNetExpert: Boolean; // Indicateur Si dossier NetExpert = envoi par Synchro CWAS
    procedure ExtractNetService(Sender: TObject);
    procedure FormatMontant(var LaChaine: string; LaTob: TOB; Champ: string);
    function CompleteZeroAGauche(St: string; L: integer): string;
    function EnvoiNetExpert(LeFic: string; NoSeqNet: Integer): Boolean;
  end;

implementation

uses
  uNEActions, UTOBDEBUG, PgOutilseAgl;

procedure TOF_PG_NETSERVICE.OnNew;
begin
  inherited;
end;

procedure TOF_PG_NETSERVICE.OnDelete;
begin
  inherited;
end;

procedure TOF_PG_NETSERVICE.OnUpdate;
begin
  inherited;
end;

procedure TOF_PG_NETSERVICE.OnLoad;
begin
  inherited;
end;

procedure TOF_PG_NETSERVICE.OnArgument(S: string);
var
  NoSEqNet, NoMAx: Integer;
  TNetRecep, TA: TOB;
  formatSq: string;
begin
  inherited;
  BtnLance := TToolbarButton97(GetControl('BVALIDER'));
  if BtnLance <> nil then BtnLance.OnClick := ExtractNetService;
  SetControlProperty('LEFICHIER', 'DataType', 'SAVEFILE(*.*)');
  if not (Ecran is TFVierge) then exit;
  { // PT9
  Ecran.Caption := 'Export S1 ASP';
  UpdateCaption(Ecran);
  }
  {
  if GetParamSoc('SO_PGCHEMINEAGL') = '' then // PGCHEMINEAGL
  begin
    PGIBox('Vous devez renseigner le chemin de stockage dans les paramètres société', Ecran.Caption);
    SetControlEnabled('BVALIDER', FALSE);
    exit;
  end;
  }
  if (V_PGI.ModePcl = '1') and IsDossierNetExpert(V_PGI.NoDossier, NoSEqNet, 'P') then
  begin
    formatSq := '%.03d';
    if NoSeqNet >= 999 then formatSq := '%.04d';
    if NoSeqNet >= 9999 then formatSq := '%.05d';
    if NoSeqNet = 0 then inc(NoSeqNet);
    TNetRecep := TOB.Create('', nil, -1);
    TA := TOB.Create('', TNetRecep, -1);
    TA.AddChampSupValeur('DOMAINE', 'PAIE');
    TA.AddChampSupValeur('CLIID', V_PGI.NoDossier);
    TA.AddChampSupValeur('CLINOM', V_PGI.NomSociete);
    TA.AddChampSupValeur('FICMSQ', 'NSV');
    NoMax := MaxNumSeq(TNetRecep);
    TNetRecep.free;
    if NoMax >= NoSeqNet then NoSeqNet := NoMax + 1;
    // GetEnvVar('TEMP') à la place de GetParamSoc('SO_PGCHEMINEAGL')
    LeFichier := GetEnvVar('TEMP') + '\PG' + V_PGI.NODOSSIER + '_' + Format(formatSq, [NoSeqNet]) + '.nsv';
    PgNetExpert := TRUE;
    SetControlVisible('LEFICHIER', FALSE);
    SetControlVisible('LBLFIC', FALSE);
  end
  else
  begin
//  LeFichier := GetEnvVar('TEMP') + '\PG' + V_PGI.NODOSSIER + '_001.nsv';
//  3eme Chgt on remet le répertoire de travail de la paie
{PT14
    LeFichier := GetParamSoc('SO_PGCHEMINEAGL') + '\PG' + V_PGI.NODOSSIER + '_001.nsv';
}
    LeFichier := VerifieCheminPG(GetParamSoc('SO_PGCHEMINEAGL')) + '\PG' + V_PGI.NODOSSIER + '_001.nsv';
//FIN PT14
    PgNetExpert := FALSE;
  end;
  SetControlText('LEFICHIER', LeFichier);
  Destinataire := GetParamSoc('SO_MAIL'); // Récupération mail dans les coordonnées paramsoc compta
  SetControlText('LEDEST', Destinataire);
  SetControlText('OBJMAIL', 'Saisie des éléments variables de la paie');
  MonMemo := TMemo(GetControl('MEMOMAIL'));
  if MonMemo <> nil then
  begin
    MonMemo.Lines.Add('Veuillez trouver ci-joint le fichier des éléments variables de la paie');
    MonMemo.Lines.Add('pour la période du ' + GetControlText('DD') + ' au ' + GetControlText('DF'));
    MonMemo.Lines.Add('');
    MonMemo.Lines.Add('Cordialement');
  end;
  if (V_PGI.ModePcl = '1') and (PgNetExpert) then
  begin // PT2
    SetControlVisible('LEDEST', FALSE);
    SetControlVisible('LBLOBJMAIL', FALSE);
    SetControlVisible('OBJMAIL', FALSE);
    SetControlVisible('LBLINFO', FALSE);
    SetControlVisible('MEMOMAIL', FALSE);
    //    SetControlText('LBLDEST', 'Transmission automatique des données au serveur NetExpert');
  end
end;

procedure TOF_PG_NETSERVICE.OnClose;
begin
  inherited;
end;

procedure TOF_PG_NETSERVICE.ExtractNetService(Sender: TObject);
var
  LaDate, DateD, DateF, D1, D2, DateDebutExerSoc, DateFinExerSoc: TDateTime;
  Q, QExer: TQuery;
  St, Stt, FileN, StSQLConfidentiel: string; // PT18
  DD, SS, Etab, LeMotif, LeMasq, SQLPlus, LaListe: string; // PT11
  F: TextFile;
  DatD, DatF: THEdit;
  TOB_Sal, T1, TOB_Etab, TOB_Motifs, TOB_MSQ: TOB;
  TRecap: TOB; // PT13
  NomFic: string; // PT13
  i, j, rep, NoSeqNet: Integer;
  ListeJ: HTStrings;
  Erreur: Boolean;
  Body: HTStringList;
// DEB PT17
  LesRub: array[1..7] of string;
  LesTyp: array[1..7] of string;
  LeTyp, lelibelle, LaRub, LeLib : string;
  OkOk: Boolean;

  function VerifRub(Quoi, Rubriq: string): Boolean;
  var k: Integer;
  begin
    result := False;
    for k := 1 to 7 do
    begin
      if Rubriq = LesRub[k] then
        if (pos(Quoi, LesTyp[k])) > 0 then
          begin
            result := TRUE;
            break;
          end;
    end;
  end;
// FIN PT17
begin
  Erreur := FALSE;
  DatD := THEdit(GetControl('DD'));
  DatF := THEdit(GetControl('DF'));
  if (DatD = nil) or (DatF = nil) then exit;
  DD := GetControlText('DD');
  DateD := StrToDate(DD);
  DateF := StrToDate(GetControlText('DF'));
  D1 := DateD; // Memorisation des dates saisies par extraction des historiques
  D2 := DateF;
//  DateD := PLUSMOIS(DateD, 1);
//  DateF := PLUSMOIS(DateF, 1);
  Etab := GetControlText('COMBOETAB'); // PT11
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
//  FileN := V_PGI.DatPath + '\CP' + PgRendNoDossier() + 'P.nsv';
  FileN := GetControlText('LEFICHIER');
  if FileExists(FileN) then
  begin
{
    reponse := HShowMessage('5;;Voulez-vous supprimer le fichier ' + ExtractFileName(FileN) + ';Q;YN;Y;N', '', '');
    if reponse = 6 then}
    DeleteFile(PChar(FileN))
  end;
  AssignFile(F, FileN);
{$I-}ReWrite(F);
{$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
    Exit;
  end;

  BtnLance.Enabled := FALSE;
  // Gestion ENTETE
{  if GetParamSoc ('PGTenueEuro')  then st := ' Monnaie=EUR'
  else st := ' Monnaie=FRF';
  writeln(F, '*** DEBUT *** ');
  if GetParamSoc ('PGTenueEuro') then St := 'EUR|'
  else st := 'FRF|';}
  QExer := OpenSql('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_DATEDEBUT<="' + USDateTime(DateD) + '" ' +
    'AND PEX_DATEFIN>="' + USDateTime(DateF) + '"', True);
  if not QExer.eof then
  begin
    DateDebutExerSoc := QExer.FindField('PEX_DATEDEBUT').AsDateTime;
    DateFinExerSoc := QExer.FindField('PEX_DATEFIN').AsDateTime;
  end
  else
  begin
    DateDebutExerSoc := idate1900;
    DateFinExerSoc := idate1900;
  end;
  Ferme(QExer);
  if (DateDebutExerSoc = idate1900) then // PT5
  begin
    PGIBOX('Le traitement est impossible car vous n''avez pas crée d''exercice social.', Ecran.caption);
    exit;
  end;
  St := 'EUR|';
  st := st + DateToStr(DateD) + '|' + DateToStr(DateF) + '|';
  st := st + ' ' + NomHalley + ' V=1.0 Généré le ' + DateToStr(Date) + '|' + V_PGI.NoDossier; // PT3
  st := st + '|' + DateToStr(DateDebutExerSoc) + '|' + DateToStr(DateFinExerSoc) + '|0800|'; // PT13 Version
  if GetParamSocSecur('SO_PGEWSMODIFSAL', FALSE) then st := st + 'X'
  else st := st + '-';
  // DEB PT13
  st := st + '|' + IntToStr(GetParamSocSecur('SO_PGJJALERTEWS', 31)) + '|' + IntToStr(GetParamSocSecur('SO_PGNBJJALERTEWS', 0));
  st := st + '|' + GetParamSocSecur('SO_PGTEXTEALERTEWS', '');
  // FIN PT13
  writeln(F, '00|' + St);
  st := '';
  // Gestion des etablissements de la paie
  TOB_Etab := TOB.Create('Les Etab', nil, -1);
  St := 'SELECT ETB_ETABLISSEMENT,ETB_LIBELLE,ETB_NBJOUTRAV FROM ETABCOMPL'; // PT6
  if Etab <> '' then St := st + ' WHERE ETB_ETABLISSEMENT="' + Etab + '"'; // PT11
  Q := OpenSql(St, TRUE);
  TOB_Etab.LoadDetailDB('ETABCOMPL', '', '', Q, False, False);
  Ferme(Q);
  for i := 0 to TOB_Etab.detail.Count - 1 do
  begin
    T1 := TOB_Etab.Detail[i];
    st := '01|' + T1.GetValue('ETB_ETABLISSEMENT') + '|' + T1.GetValue('ETB_LIBELLE') + '|' + FloatToStr(T1.GetValue('ETB_NBJOUTRAV')); // PT6
    writeln(F, St);
  end;
  FreeAndNil(TOB_Etab);
  // DEB PTXX
  // GESTION DES MASQUES DE SAISIE
  TOB_MSQ := TOB.Create('Les masques', nil, -1);
  St := 'SELECT * FROM MASQUESAISRUB ';
  // DEB PT10
  LaListe := GetParamSocSecur('SO_PGMASQEWS', '');
  if (LaListe <> '') and (POS('<<', LaListe) = 0) then
  begin
    SqlPlus := ' WHERE PMR_ORDRE IN (';
    LeMasq := ReadTokenSt(LaListe);
    while LeMasq <> '' do
    begin
      SqlPlus := SqlPlus + '"' + LeMasq + '"';
      LeMasq := ReadTokenst(LaListe);
      if LeMasq <> '' then SqlPlus := SqlPlus + ',';
    end;
    SqlPlus := SqlPlus + ')';
    st := st + SQLPlus;
  end;
  // FIN PT10
  Q := OpenSql(St, TRUE);
  TOB_MSQ.LoadDetailDB('MASQUESAISRUB', '', '', Q, False, False);
  Ferme(Q);
  OkOk := TRUE; // PT17
  for i := 0 to TOB_MSQ.detail.Count - 1 do
  begin
    T1 := TOB_MSQ.Detail[i];
    lelibelle := T1.GetValue('PMR_LIBELLE'); // PT17
    st := '02|' + T1.GetValue('PMR_ORDRE') + '|' + T1.GetValue('PMR_LIBELLE');
    for j := 1 to 7 do
    begin
      SS := IntToStr(j);
      if T1.GetValue('PMR_COL' + SS) <> '' then
      begin
        LesRub[j] := T1.GetValue('PMR_COL' + SS);
        LesTyp[j] := T1.GetValue('PMR_TYPECOL' + SS);
      end
      else
      begin
        LesRub[j] := '';
        LesTyp[j] := '';
      end;
      st := st + '|' + T1.GetValue('PMR_COL' + SS) + '|' + T1.GetValue('PMR_LIBCOL' + SS) + '|' + T1.GetValue('PMR_TYPECOL' + SS);
      st := st + '|' + T1.GetValue('PMR_AIDECOL' + SS); //PT15
    end;
    // DEB PT17 En fin de masque on vérifie si tous les champs saisissables de la rubrique sont dans le masque
    for j := 1 to 7 do
    begin
      if LesRub[j] = '' then break;
      // DEB PT18
      stt := 'SELECT PRM_TYPEBASE,PRM_TYPETAUX,PRM_TYPECOEFF,PRM_TYPEMONTANT,PRM_LIBELLE FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE="' + LesRub[j] + '"';
      Q := OpenSql(stt, true);
      // FIN PT18
      if not Q.EOF then
      begin
        LaRub := LesRub [j];
        LeLib := Q.FindField('PRM_LIBELLE').AsString;
        LeTyp := Q.FindField('PRM_TYPEBASE').AsString;
        if (LeTyp = '00') or (LeTyp = '01') then OkOk := VerifRub('B', LesRub[j]);
        if not OkOk then break; // Le masque ne comporte pas les tous champs saisissables
        LeTyp := Q.FindField('PRM_TYPETAUX').AsString;
        if (LeTyp = '00') or (LeTyp = '01') then OkOk := VerifRub('T', LesRub[j]);
        if not OkOk then break; // Le masque ne comporte pas les tous champs saisissables
        LeTyp := Q.FindField('PRM_TYPECOEFF').AsString;
        if (LeTyp = '00') or (LeTyp = '01') then OkOk := VerifRub('C', LesRub[j]);
        if not OkOk then break; // Le masque ne comporte pas les tous champs saisissables
        LeTyp := Q.FindField('PRM_TYPEMONTANT').AsString;
        if (LeTyp = '00') or (LeTyp = '01') then OkOk := VerifRub('M', LesRub[j]);
        if not OkOk then break; // Le masque ne comporte pas les tous champs saisissables
      end;
      Ferme(Q);
    end;
    if not OkOk then break; // Masque pas OK on arrête le traitement
  // FIN PT17
{    for j := 1 to 7 do   PT15
    begin
      SS := IntToStr(j);
      st := st + '|' + T1.GetValue('PMR_AIDECOL' + SS);  //PT10
    end;
}
    writeln(F, St);
  end;
  FreeAndNil(TOB_MSQ);
  if Not OkOk then // PT17
  begin
    st := 'Le masque '+lelibelle+ 'est mal paramétré #13#10'+
          ' car tous les champs saisissables de la rubrique '+LaRub+ ' '+Lelib+ ' #13#10';
    st := st + 'ne sont pas présents dans le masque.';
    PgiBox (St ,'Traitement abandonné '+Ecran.caption);
    exit;
  end;
  // FIN PTXX
  // GESTION DES SALARIES PRESENTS sur la paie à venir
  TOB_Sal := TOB.Create('Les Salaries', nil, -1);
  LaDate := DateD;
  if GetParamSocSecur('SO_PGSALSORTISEWS', FALSE) then
    LaDate := PLUSMOIS(LaDate, GetParamSocSecur('SO_PGMOISSALEWS', 0) * -1);
  // DEB PT3
  St := 'SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,PSA_ADRESSE1,PSA_ADRESSE2,PSA_ADRESSE3,';
  St := St + 'PSA_CODEPOSTAL,PSA_VILLE,PSA_TELEPHONE,PSA_PORTABLE,PSA_DATENAISSANCE,PSA_NATIONALITE,PSA_SEXE,PSA_NUMEROSS,PSA_CIVILITE,';
  st := st + 'PSA_NUMEROBL,PSA_NATIONALITE,PSA_DEPTNAISSANCE,PSA_COMMUNENAISS,PSA_PAYSNAISSANCE,PSA_CARTESEJOUR,PSA_DELIVPAR,PSA_DATEXPIRSEJOUR,'; // PT10
  st := st + 'PSA_SITUATIONFAMIL,PSA_PERSACHARGE,PSA_MOTIFENTREE,PSA_HEUREMBAUCHE,PSA_MOTIFSORTIE,PSA_CATDADS,PSA_REGIMESS,PSA_DADSCAT,PSA_DADSPROF,'; // PT10
  st := st + 'PSA_UNITETRAVAIL,PSA_CONDEMPLOI,PSA_TAUXPARTIEL,PSA_SORTIEDEFINIT '; // PT10 , PSA_UNITEPRISEFF PT16
  st := st + ' FROM SALARIES WHERE PSA_SUSPENSIONPAIE <> "X"';
  //PT12
  StSQLConfidentiel := '';
  StSQLConfidentiel := SqlConf('SALARIES');
  if StSQLConfidentiel <> '' then
    st := st + ' AND ' + StSQLConfidentiel;
  //Fin PT12
  St := St + ' AND (PSA_DATEENTREE <="' + UsDateTime(DateF) + '") AND ((PSA_DATESORTIE >="' + UsDateTime(LaDate) +
    '") OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "' + UsDateTime(iDate1900) + '")) AND (PSA_DATEENTREE <="' + UsDateTime(DateF) + '")';
  St := St + ' AND PSA_NUMEROSS <> "" AND PSA_NUMEROSS IS NOT NULL'; // PT11
  if Etab <> '' then St := st + ' AND PSA_ETABLISSEMENT="' + Etab + '"'; // PT11
  Q := OpenSql(St, TRUE);
  TOB_Sal.LoadDetailDB('SALARIES', '', '', Q, False, False);
  Ferme(Q);
  for i := 0 to TOB_Sal.detail.Count - 1 do
  begin
    T1 := TOB_Sal.Detail[i];
    st := 'A0|' + T1.GetValue('PSA_ETABLISSEMENT') + '|' + T1.GetValue('PSA_SALARIE') + '|' + T1.GetValue('PSA_LIBELLE') + '|' + T1.GetValue('PSA_PRENOM') +
      '|' + DateToStr(T1.GetValue('PSA_DATEENTREE')) + '|' + DateToStr(T1.GetValue('PSA_DATESORTIE'));
    st := st + '|' + T1.GetValue('PSA_ADRESSE1') + '|' + T1.GetValue('PSA_ADRESSE2') + '|' + T1.GetValue('PSA_ADRESSE3') + '|' + T1.GetValue('PSA_CODEPOSTAL') + '|' + T1.GetValue('PSA_VILLE') +
      '|' + T1.GetValue('PSA_TELEPHONE') + '|' + T1.GetValue('PSA_PORTABLE') + '|' + DateToStr(T1.GetValue('PSA_DATENAISSANCE')) + '|' + T1.GetValue('PSA_NATIONALITE') + '|';
    st := st + T1.GetValue('PSA_SEXE') + '|' + T1.GetValue('PSA_NUMEROSS') + '|' + T1.GetValue('PSA_CIVILITE') + '|';
    // DEB PT10
    st := st + T1.GetValue('PSA_NUMEROBL') + '|' + T1.GetValue('PSA_NATIONALITE') + '|' + T1.GetValue('PSA_DEPTNAISSANCE') + '|';
    st := st + T1.GetValue('PSA_COMMUNENAISS') + '|' + T1.GetValue('PSA_PAYSNAISSANCE') + '|' + T1.GetValue('PSA_CARTESEJOUR') + '|';
    st := st + T1.GetValue('PSA_DELIVPAR') + '|' + DateToStr(T1.GetValue('PSA_DATEXPIRSEJOUR')) + '|' + T1.GetValue('PSA_SITUATIONFAMIL') + '|';
    st := st + IntToStr(T1.GetValue('PSA_PERSACHARGE')) + '|' + T1.GetValue('PSA_MOTIFENTREE') + '|' + TimeToStr(T1.GetValue('PSA_HEUREMBAUCHE')) + '|';
    st := st + T1.GetValue('PSA_MOTIFSORTIE') + '|' + T1.GetValue('PSA_CATDADS') + '|' + T1.GetValue('PSA_REGIMESS') + '|' + T1.GetValue('PSA_DADSCAT') + '|';
    st := st + T1.GetValue('PSA_DADSPROF') + '|' + T1.GetValue('PSA_UNITETRAVAIL') + '|' + T1.GetValue('PSA_CONDEMPLOI') + '|';
    st := st + StrfMontant(T1.GetValue('PSA_TAUXPARTIEL'), 2, 0, '', TRUE) + '|' + T1.GetValue('PSA_SORTIEDEFINIT');
    // FIN PT10
    writeln(F, St);
    // FIN PT3
  end;
{ // DEB PT8
  // GESTION DES REMUNERATIONS SALARIES du mois de référence
  TOB_Hist := TOB.Create('Les lignes bulletins', nil, -1);
  St := 'SELECT PHB_SALARIE,PHB_RUBRIQUE,PHB_LIBELLE,PHB_BASEREM,PHB_TAUXREM,PHB_COEFFREM,PHB_MTREM,PRM_CODECALCUL,PRM_TYPEBASE,' +
    'PRM_TYPETAUX,PRM_TYPECOEFF,PRM_TYPEMONTANT,PRM_DECBASE,PRM_DECTAUX,PRM_DECCOEFF,PRM_DECMONTANT FROM HISTOBULLETIN ' +
    ' right outer join REMUNERATION ON PRM_RUBRIQUE=PHB_RUBRIQUE WHERE ##PRM_PREDEFINI## (PRM_TYPEBASE="00" OR PRM_TYPETAUX="00" OR ' +
    ' PRM_TYPECOEFF="00" OR PRM_TYPEMONTANT="00" OR PRM_TYPEBASE="01" OR PRM_TYPETAUX="01" OR ' +
    ' PRM_TYPECOEFF="01" OR PRM_TYPEMONTANT="01") AND PHB_NATURERUB="AAA" AND PHB_DATEDEBUT >= "' + UsDateTime(D1) + '"' +
    ' AND PHB_DATEFIN <= "' + UsDateTime(D2) + '"';
  if Etab <> '' then St := st + ' AND PHB_ETABLISSEMENT="'+Etab+'"'; // PT11
  st := st + ' ORDER BY PHB_SALARIE,PHB_RUBRIQUE'; // PT11
  Q := OpenSql(St, TRUE);
  TOB_Hist.LoadDetailDB('HISTOBULLETIN', '', '', Q, False, False);
  Ferme(Q);
  for i := 0 to TOB_Hist.detail.Count - 1 do
  begin
    T1 := TOB_Hist.Detail[i];
    TS := TOB_Sal.FindFirst(['PSA_SALARIE'], [T1.GetVAlue('PHB_SALARIE')], FALSE);
    if TS = nil then Continue;
    LeMontant := CompleteZeroAGauche(FloatToStr(T1.GetValue('PHB_BASEREM')), 13); // Montant du cumul * 100 pour eviter les , ou .
    st := 'B0|' + T1.GetValue('PHB_SALARIE') + '|' + T1.GetValue('PHB_RUBRIQUE') + '|' + T1.GetValue('PHB_LIBELLE') + '|' + T1.GetValue('PRM_CODECALCUL') +
      '|' + T1.GetValue('PRM_TYPEBASE') + '|' + T1.GetValue('PRM_TYPETAUX') + '|' + T1.GetValue('PRM_TYPECOEFF') + '|' + T1.GetValue('PRM_TYPEMONTANT');
    FormatMontant(st, T1, 'PHB_BASEREM');
    FormatMontant(st, T1, 'PHB_TAUXREM');
    FormatMontant(st, T1, 'PHB_COEFFREM');
    FormatMontant(st, T1, 'PHB_MTREM');
    St := St + '|' + CompleteZeroAGauche(FloatToStr(T1.GetValue('PRM_DECBASE')), 1);
    St := St + '|' + CompleteZeroAGauche(FloatToStr(T1.GetValue('PRM_DECTAUX')), 1);
    St := St + '|' + CompleteZeroAGauche(FloatToStr(T1.GetValue('PRM_DECCOEFF')), 1);
    St := St + '|' + CompleteZeroAGauche(FloatToStr(T1.GetValue('PRM_DECMONTANT')), 1);
    writeln(F, St);
  end;
  FreeAndNil (TOB_Hist);
 // FIN PT8
  }
  // DEB PTYY
  // Gestion des recapitulatifs congés payés et RTT des salariés
  ChargeTobSalRecap('N'); // Calcul et Chargement des récapitulatifs en fonction des périodes CP
{PT14
  NomFic := GetParamSoc('SO_PGCHEMINEAGL') + '\recapsal.txt';
}
  NomFic := VerifieCheminPG(GetParamSoc('SO_PGCHEMINEAGL')) + '\recapsal.txt';
//FIN PT14
  if not FileExists(NomFic) then
  begin
    PGIBox('Fichier inaccessible : ' + NomFic, 'Abandon du traitement');
    writeln(F, '*** Erreur de traitement des Cumuls CP *** ');
    writeln(F, '*** FIN *** ');
    CloseFile(F);
    Exit;
  end;
  TRecap := tob.create('', nil, -1);
  TobLoadfromfile(NomFic, nil, TRecap);
  for i := 0 to TRecap.detail.Count - 1 do
  begin
    T1 := TRecap.Detail[i];
    st := 'C0|' + T1.GetValue('PRS_SALARIE');
    FormatMontant(st, T1, 'PRS_ACQUISN1');
    FormatMontant(st, T1, 'PRS_PRISN1');
    FormatMontant(st, T1, 'PRS_RESTN1');
    FormatMontant(st, T1, 'PRS_ACQUISN');
    FormatMontant(st, T1, 'PRS_PRISN');
    FormatMontant(st, T1, 'PRS_RESTN');
    FormatMontant(st, T1, 'PRS_CUMRTTACQUIS');
    FormatMontant(st, T1, 'PRS_CUMRTTPRIS');
    writeln(F, St);
  end;
  TRecap.Free;
//  TRecap := nil;
  // FIN PTYY
  // GESTION DES CUMULS SALARIES du mois de référence
{ PT13 suppression des cumuls salariés + des cumuls de paie
  TOB_CumSal := TOB.Create('Les cumuls salaries', nil, -1);
  St := 'SELECT PHC_SALARIE,PHC_CUMULPAIE,PHC_MONTANT,PCL_LIBELLE FROM HISTOCUMSAL LEFT JOIN CUMULPAIE ON PCL_CUMULPAIE=PHC_CUMULPAIE WHERE ##PCL_PREDEFINI## ' +
    'PHC_DATEDEBUT >= "' + UsDateTime(D1) + '" AND PHC_DATEFIN <= "' + UsDateTime(D2) + '" AND ((PHC_CUMULPAIE ="01") OR ' +
    '(PHC_CUMULPAIE ="09") OR (PHC_CUMULPAIE ="10") OR (PHC_CUMULPAIE ="20"))';
  if Etab <> '' then St := st + ' AND PHC_ETABLISSEMENT="'+Etab+'"'; // PT11
  st := st + ' ORDER BY PHC_SALARIE,PHC_CUMULPAIE'; //PT11
  Q := OpenSql(St, TRUE);
  TOB_CumSal.LoadDetailDB('HISTOCUMSAL', '', '', Q, False, False);
  Ferme(Q);
  for i := 0 to TOB_CumSal.detail.Count - 1 do
  begin
    T1 := TOB_CumSal.Detail[i];
    TS := TOB_Sal.FindFirst(['PSA_SALARIE'], [T1.GetValue('PHC_SALARIE')], FALSE);
    if TS = nil then Continue;
    st := 'D0|' + T1.GetValue('PHC_SALARIE') + '|' + T1.GetValue('PHC_CUMULPAIE');
    FormatMontant(st, T1, 'PHC_MONTANT');
    writeln(F, St);
  end;
  FreeAndNil ( TOB_CumSal);
  FreeAndNil (TOB_Sal);
  st := '';
  // Gestion des cumuls de la paie
  TOB_Cum := TOB.Create('Les Cumuls', nil, -1);
  St := 'SELECT PCL_CUMULPAIE,PCL_LIBELLE FROM CUMULPAIE WHERE PCL_CUMULPAIE="01" OR PCL_CUMULPAIE="09" OR PCL_CUMULPAIE="10" OR PCL_CUMULPAIE="20"';
  Q := OpenSql(St, TRUE);
  TOB_Cum.LoadDetailDB('CUMULPAIE', '', '', Q, False, False);
  Ferme(Q);
  for i := 0 to TOB_Cum.detail.Count - 1 do
  begin
    T1 := TOB_Cum.Detail[i];
    st := 'E0|' + T1.GetValue('PCL_CUMULPAIE') + '|' + T1.GetValue('PCL_LIBELLE');
    writeln(F, St);
  end;
  FreeAndNil (TOB_Cum);
  }
  // PT1 : 15/10/2001 V562 PH Export des motifs d'absences
  // Gestion des motifs absences de la paie
  TOB_Motifs := TOB.Create('Les Motifs', nil, -1);
  St := 'SELECT PMA_MOTIFABSENCE,PMA_LIBELLE FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_MOTIFEAGL = "X" ';
  // DEB PT10
  LaListe := GetParamSocSecur('SO_PGMOTIFEWS', '');
  if (LaListe <> '') and (POS('<<', LaListe) = 0) then
  begin
    SqlPlus := ' AND PMA_MOTIFABSENCE IN (';
    LeMotif := ReadTokenst(LaListe);
    while LeMotif <> '' do
    begin
      SqlPlus := SqlPlus + '"' + LeMotif + '"';
      LeMotif := ReadTokenst(LaListe);
      if LeMotif <> '' then SqlPlus := SqlPlus + ',';
    end;
    SqlPlus := SqlPlus + ')';
    st := st + SQLPlus;
  end;
  // FIN PT10
  Q := OpenSql(St, TRUE);
  TOB_Motifs.LoadDetailDB('MOTIFABSENCE', '', '', Q, False, False);
  Ferme(Q);
  for i := 0 to TOB_Motifs.detail.Count - 1 do
  begin
    T1 := TOB_Motifs.Detail[i];
    st := 'F0|' + T1.GetValue('PMA_MOTIFABSENCE') + '|' + T1.GetValue('PMA_LIBELLE');
    writeln(F, St);
  end;
  FreeAndNIL(TOB_Motifs);
  // Gestion FIN DE FICHIER
  writeln(F, '*** FIN *** ');
  CloseFile(F);

  if (V_PGI.ModePcl = '1') and IsDossierNetExpert(V_PGI.NoDossier, NoSEqNet) then
    Erreur := EnvoiNetExpert(FileN, NoSeqNet) // PT2
  else
  begin
    if GetControlText('LEDEST') <> '' then
    begin
      rep := PGIAsk('L''extraction est terminée, voulez vous transmettre le fichier par mail ?', Ecran.Caption);
      if rep = mryes then
      begin
        ListeJ := HTStringList.Create;
//        i := MonMemo.Lines.Count;
        for i := 0 to MonMemo.Lines.Count do
        begin
          ListeJ.Add(MonMemo.lines[i]);
        end;

        SendMail(GetControlText('OBJMAIL'), GetControlText('LEDEST'), '', ListeJ, FileN, FALSE);
        ListeJ.Clear;
        ListeJ.Free;
        if GetControlText('LEDEST') <> GetParamSoc('SO_MAIL') then SetParamSOc('SO_MAIL', GetControlText('LEDEST'));
      end;
    end
    else
      PGIBOX('Le fichier a été généré sur disque', Ecran.Caption);
  end;
  if (V_PGI.ModePcl = '1') and Erreur then
  begin
    Body := HTStringList.Create;
    Body.add('Veuillez trouver ci-joint le rapport d''erreur correspondant au dossier ' + V_PGI.NoDossier);
    Body.add('Le traitement a échoué, le fichier n''a pas été déposé');
    SendMail('Rapport PAIE-GRH PGI Expert', 'synchro-S1-S5@cegid.fr', '', Body, '', true, 1, '', '');
    BoDy.free;
  end;
  SetParamSoc('SO_PGDATEEWS', Date);
end;
// Fonction unique de formatage des double de façon à avoir une cohérence avec tous les champs de type double

procedure TOF_PG_NETSERVICE.FormatMontant(var LaChaine: string; LaTob: TOB; Champ: string);
begin
  LaChaine := LaChaine + '|' + CompleteZeroAGauche(FloatToStr(LaTob.GetValue(Champ)), 13); // Montant du cumul pour eviter les , ou .
end;

function TOF_PG_NETSERVICE.EnvoiNetExpert(LeFic: string; NoSeqNet: Integer): Boolean;
var
  TNetEnvoi, TA, OutTOB: TOB;
  Filearchive, Ext, LeFichier, formatSq: string;
  i: integer;
  Info: string;
  Erreur: Boolean;
begin
  //$|0||
  //=|0||
  //$|1||DOMAINE|CLIID |CLINOM|FICNOM |FICMSQ|FICTYPE|DATEARRET|FICSEQ|CREDAT|FICCTRL|DESCRIPTIF|
  //=|1||PAIE |10001|RUBI|PG10001_002|NSV|DOS|2004-04-21|002|2004-04-30 02:10:12|-|descriptif du fichier|
  if not FileExists(LeFic) then
  begin
    {    LeFichier := ExtractFileName(LeFic);
        LeFichier := GetEnvVar('TEMP') + '\' + LeFichier;
        if not RenameFile(LeFic, LeFichier) then
        begin}
//    Erreur := FALSE;
    result := FALSE;
    PGiError('Le fichier ' + LeFic + ' ne peut pas être transmis', Ecran.Caption);
    exit;
    //    end;
  end;
  LeFichier := ExtractFileName(LeFic);
  TNetEnvoi := TOB.Create('', nil, -1);
  TA := TOB.Create('', TNetEnvoi, -1);
  TA.AddChampSupValeur('DOMAINE', 'PAIE');
  TA.AddChampSupValeur('CLIID', V_PGI.NoDossier);
  TA.AddChampSupValeur('CLINOM', V_PGI.NomSociete);
  Ext := ExtractFileName(LeFichier);
  Filearchive := ReadTokenPipe(Ext, '.');
  TA.AddChampSupValeur('FICNOM', Filearchive);
  Ext := ExtractFileExt(LeFichier);
  TA.AddChampSupValeur('FICMSQ', Copy(Ext, 2, length(Ext)));
  TA.AddChampSupValeur('FICTYPE', 'NSV');
  TA.AddChampSupValeur('DATEARRET', FormatDateTime('yyyy-mm-dd', StrToDate(GetControlText('DF'))));
  formatSq := '%.03d';
  if NoSeqNet >= 999 then formatSq := '%.04d';
  if NoSeqNet >= 9999 then formatSq := '%.05d';
  TA.AddChampSupValeur('FICSEQ', Format(formatSq, [NoSeqNet]));
  TA.AddChampSupValeur('CREDAT', FormatDateTime('yyyy-mm-dd hh:nn', NowH));
  TA.AddChampSupValeur('FICCTRL', '-');
  Info := 'Dossier : ' + V_PGI.NoDossier + ' éléments variables de paie du : ' +
    FormatDateTime('dd/mm/yyyy', StrToDate(GetControlText('DD'))) +
    ' au ' + FormatDateTime('dd/mm/yyyy', StrToDate(GetControlText('DF')));
  TA.AddChampSupValeur('DESCRIPTIF', Info);
  //  if V_PGI.SAV then TOBDEBUG (TNetEnvoi);
  OutTOB := NEEnvoi(TNetEnvoi);
  Erreur := FALSE;
  if OutTOB.GetValue('ERROR') = 0 then
  begin
    for i := 0 to OutTOB.detail.count - 1 do
    begin
      if OutTOB.detail[i].GetValue('ERROR') <> '0' then
      begin
        PGIInfo(OutTOB.detail[i].GetValue('ERRORLIB'));
        Erreur := TRUE;
        DeleteFile(LeFichier);
        break;
      end;
    end;
    if not Erreur then
    begin
      DeleteFile(LeFichier);
      inc(NoSeqNet);
      {      ExecuteSQL('UPDATE DOSSIER SET DOS_NECPSEQ=' + IntToStr(NoSeqNet) +
              ' Where DOS_NODOSSIER="' + V_PGI.NoDossier + '"');}
      ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA=' + IntToStr(NoSeqNet) +
        ' Where SOC_NOM = "SO_NEPAISEQ"');

    end;
  end
  else
  begin
    DeleteFile(LeFichier);
    PGIBox(OutTOB.GetValue('ERRORLIB'));
    TNetEnvoi.Free;
    exit;
  end;
  if not Erreur then PGIInfo('Le fichier a été envoyé');
  TNetEnvoi.Free;
  result := Erreur;
end;

function TOF_PG_NETSERVICE.CompleteZeroAGauche(St: string; L: integer): string;
begin
  St := Copy(St, 1, L);
  while Length(St) < L do St := '0' + St;
  Result := St;
end;

initialization
  registerclasses([TOF_PG_NETSERVICE]);
end.

