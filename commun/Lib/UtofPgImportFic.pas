 {***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 14/06/2001
Modifié le ... :   /  /
Description .. : Unit de traitement du fichier extérieur à la paie
Suite ........ : Utile pour traiter des absences, des saisies d'éléments
Suite ........ : de paie provenant d'un autre logiciel,
Suite ........ : de heures de présences et/ou d'absences provenant d'une
Suite ........ : pointeuse
Suite ........ : Utilisé pour récupérer  les absences venant d'une saisie
Suite ........ : déportée.
Mots clefs ... : PAIE;ABSENCES;IMPORT
*****************************************************************}
{ PT1 14/06/01 PH
      Rajout libellés complémentaires dans le traitement des lignes absences pour
      prendre en compte les libellés complémentaires lors de la récup des absences
      déportées
  PT2 09/10/01 PH Rajout traitement ligne Analytique
  PT3 26/11/01 PH Récupération du libellé saisi de l'absence et non pas un libellé forcé
  PT4 26/11/01 PH Incrementation du numéro d'ordre des mvts de type CP
  PT5 27/11/01 PH Initialisation de tous les champs de la table absencesalarie
  PT6 27/11/01 PH Si on ne gère pas les cp et que l'on a des lignes de cp,on perd les lignes
                  sans que l'utilisateur le sache
  PT7 18/12/01 PH remplacement strtofloat par valeur sinon erreur de conversion
  PT8 15/02/02 PH test pour les mvts de cp sur mois clos N et N+1 pour prendre les mvts
                  saisis sur le mois precedant (même si clos) au dernier mois
  PT9 28/10/2002 SB On considère que c'est le salarie qui saisie les mvts absences de l'import
  PT10 30/10/2002 SB On tronque le libellé de l'absence à 35 caractères
PT11-1 02/12/2002 PH Regroupement lecture du fichier de façon à trouver matin et après midi
                     pour le controle que l'absence n'a pas été déjà saisie
PT11-2 02/12/2002 PH Si Salarié sorti et absence déjà saisie alors OK
PT12   28/01/2003 SB Erreur appli si type separateur différent de celui utilisé
PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
PT14   04/06/2003 PH Création automatique des sections analytiques lors de l'import si autorisé
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT15   25/08/2003 PH Correction prise en compte traitement analytique ANA et VEN
PT16   28/08/2003 PH Création des sections en récupérant le libellé de la section
PT17   08/09/2003 PH Destruction des mouvements ANA
PT18   10/09/2003 PH Controle du séparateur du fichier import et creation automatique fichier rapport
PT19   16/12/2003 PH OPtimisation de code pour éviter requêtes répétitives + traitement des absences sur périodes closes
PT20   06/01/2004 PH Suppression des ventilations analytiques salarie uniquement sur le même axe que celles contenues dans le fichier d'import
PT21   12/03/2004 SB V_50 FQ 11162 Encodage de de la date de cloture erroné si fin fevrier
PT22   09/04/2004 SB V_50 FQ 11136 Ajout Gestion des congés payés niveau salarié
PT22   26/05/2004 PH V_50 Non traitement des absences portant sur des périodes closes mias ne provoquant pas de rejet du fichier
PT23   03/06/2004 PH V_50 FQ 11271 Erreur SAV TOB.getvalue sur un champ non existant ==> provoque une erreur ?
PT24   08/09/2004 PH V_50 Modif pour traitement automatique fichier NetService PCL
PT25   20/09/2004 PH V_50 Source sans uses spécifiques de la paie
PT26   02/11/2004 PH V_60 Prise en compte des rubriques de type permanent dans le fichier d'import
PT27   02/11/2004 PH V_60 FQ 11409 Traitement Analytique salarié pour un mois + Création de salarié
PT28   30/11/2004 PH V_60 FQ 11794 Fin de traitement OK
}
unit UtofPgImportFic;

interface
uses
  Variants,

  SysUtils,
  HCtrls,
  HEnt1,
  Classes,
  StdCtrls,
{$IFNDEF EAGLCLIENT}
  Db,
  HDB,
  Hqry,
  Fiche,
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  DBCtrls,
  QRe,
{$ELSE}
  eFiche,
{$ENDIF}
  Controls,
  forms,
  Vierge,
  ComCtrls,
  HMsgBox,
  Math,
  UTOM,
  UTOB,
  UTOF,
  HTB97,
  ParamSoc,
  Dialogs,
  HRichOLE,
  HStatus,
{$IFDEF AFFAIRE}
  {$IFNDEF OGC}
  PgOutils2,
  {$ENDIF}
{$ENDIF}
  ed_tools;
type
  TOF_PGIMPORTFIC = class(TOF)
  private
    MarqueDebut, MarqueFin, typebase, typetaux, typecoeff, typemontant, SalariePrec, RubPrec: string;
    TFinPer, DateEntree, DateSortie: TDatetime;
    Separateur: char;
    MarqueFinOk, Erreur_10, ExisteErreur, LibErreur, libanomalie, AbsExclue: boolean;
    RapportErreur: THRichEditOLE;
    Tob_sal, Tob_SalATT, tob_rubrique, Tob_section, Tob_Abs, Tob_HSR, Tob_OrdreAbs, Tob_OrdreHSR, Tob_PaieVentil: tob;
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
    Tob_Ventil, Tob_Etab: TOB; // PT27 Rajout traitement creation de salarie vérif etablissement
    FR: TextFile;
    F: TextFile;
    taille: integer;
    Sep: Thvalcombobox; // PT18
    NomDuFichier: string;
    TypImport: string;
    Cloture: string; // PT19
    MatSouche: Integer; // Matricule souche pour la création automatique
    TopCreat : Boolean ; // Top salariés déjà dejà crées ou modifiés
    DebExer, FinExer: tdatetime;
    procedure FicChange(Sender: TObject);
    procedure SepChange(Sender: TObject);
    procedure Traiteligne(chaine, S: string);
    procedure Erreur(Chaine, S: string; NoErreur: integer; WW: string = '');
    procedure Anomalie(Chaine, S: string; NoErreur: integer);
    function ControleEnrDossier(chaine, S: string): integer;
    procedure ControleEnregMHE_MFP_MLB(TypeEnr, chaine, S: string; Ecrit: boolean);
    procedure ControleEnregMAB(chaine, S: string; Ecrit: boolean);
    procedure ControleEnregANA(TypeEnr, chaine, S: string; Ecrit: boolean);
    procedure ControleEnregVSA(TypeEnr, chaine, S: string; Ecrit: boolean); // PT27
    procedure ControleEnregSAL(TypeEnr, chaine, S: string; Ecrit: boolean); // PT27
    procedure ControleEnregVEN(TypeEnr, chaine, S: string; Ecrit: boolean);
    {$IFNDEF EAGLCLIENT}
    procedure ImprimeClick(Sender: TObject);
    {$ENDIF !EAGLCLIENT}
    function ExisteRubrique(Rubrique: string): boolean;
    function ExisteSection(Axe, Section, LibSect: string): boolean;
    procedure ControleSalarie(TypeEnr, chaine, S: string; var CodeSalarie: string; Ecrit: Boolean);
    procedure ConsultErreurClick(Sender: Tobject);
    procedure ValideEcran(Top: boolean);
    procedure ImporteFichier;
    function IsMarqueDebut(S: string): boolean;
    function IsMarqueFin(S: string): boolean;
    procedure FermeTout;
    function IsPeriodecloture(Dated, DateF: TDateTime; CasCp: string = '-'): integer;
    function RendNumOrdre(TypeMvt, CodeSalarie: string): integer;
    function PgImportDef(LeNomduFichier, LeTypImp: string; LeSeparateur: Char = '|'): Boolean;
    procedure ControleFichierImport(Sender: Tobject);
    function ImportRendPeriodeEnCours(var ExerPerEncours, DebPer, FinPer: string): Boolean;
    function ImportRendExerSocialEnCours(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime): Boolean;
    function ImportColleZeroDevant(Nombre, LongChaine: integer): string;
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
  end;
  {
  NoErreur :
  5 : Manque marque début fichier

  }

implementation
uses TiersUtil,UtilPgi,ENt1;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : fonction d'initialisation d'un mouvement congés payés
Suite ........ : A appeler avant tout renseignement de nouveau mvt
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure ImportInitialiseTobAbsenceSalarie(TCP: tob);
begin
  TCP.Putvalue('PCN_TYPEMVT', 'CPA');
  TCP.Putvalue('PCN_SALARIE', '');
  TCP.Putvalue('PCN_DATEDEBUT', 0);
  TCP.Putvalue('PCN_DATEFIN', 0);
  TCP.Putvalue('PCN_ORDRE', 0);
  TCP.Putvalue('PCN_TYPECONGE', '');
  TCP.Putvalue('PCN_SENSABS', '');
  TCP.PutValue('PCN_PERIODEPY', -1);
  TCP.Putvalue('PCN_LIBELLE', '');
  TCP.Putvalue('PCN_DATEMODIF', 0);
  TCP.Putvalue('PCN_DATESOLDE', 0);
  TCP.Putvalue('PCN_DATEVALIDITE', 0);
  TCP.PutValue('PCN_PERIODECP', -1);
  TCP.Putvalue('PCN_DATEDEBUTABS', 0);
  TCP.Putvalue('PCN_DATEFINABS', 0);
  TCP.Putvalue('PCN_DATEPAIEMENT', 0);
  TCP.Putvalue('PCN_CODETAPE', '...'); //PT-13
  TCP.Putvalue('PCN_JOURS', 0);
  TCP.Putvalue('PCN_BASE', 0);
  TCP.Putvalue('PCN_NBREMOIS', 0);
  TCP.Putvalue('PCN_CODERGRPT', 0);
  TCP.PutValue('PCN_MVTDUPLIQUE', '-');
  TCP.Putvalue('PCN_ABSENCE', 0);
  TCP.Putvalue('PCN_MODIFABSENCE', '-');
  TCP.Putvalue('PCN_APAYES', 0);
  TCP.Putvalue('PCN_VALOX', 0);
  TCP.Putvalue('PCN_VALOMS', 0);
  TCP.Putvalue('PCN_VALORETENUE', 0);
  TCP.Putvalue('PCN_MODIFVALO', '-');
  TCP.Putvalue('PCN_PERIODEPAIE', '');
  TCP.Putvalue('PCN_ETABLISSEMENT', '');
  TCP.Putvalue('PCN_TRAVAILN1', '');
  TCP.Putvalue('PCN_TRAVAILN2', '');
  TCP.Putvalue('PCN_TRAVAILN3', '');
  TCP.Putvalue('PCN_TRAVAILN4', '');
  TCP.PutValue('PCN_GENERECLOTURE', '-');
  // a voir   TCP.PutValue('PCN_CONFIDENTIEL','-');
end;
// DEB PT25

function PGImportEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime;
begin
  Result := Idate1900;
  if IsValidDate(IntToStr(JJ) + '/' + IntToStr(MM) + '/' + IntToStr(AA)) then
    Result := encodedate(AA, MM, JJ);
  if (MM = 2) and ((JJ = 28) or (JJ = 29)) then //Année bissextile
  begin
    Result := encodedate(AA, MM, 1);
    Result := FindeMois(Result);
  end;
end;
// FIN PT25

function ImportCalculPeriode(DTClot, DTValidite: TDatetime): integer;
var
  Dtdeb, DtFin, DtFinS: TDATETIME;
  aa, mm, jj: word;
  i: integer;
begin
  result := -1;
  if DTClot <= idate1900 then exit; //PT53
  Decodedate(DTclot, aa, mm, jj);
  DtDeb := PGImportEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT21 } // PT25
  DtFin := DtClot;
  DtFinS := PGImportEncodeDateBissextile(aa + 1, mm, jj); { PT21 } // PT25
  if Dtvalidite > Dtfins then
  begin
    result := -9;
    exit;
  end;
  if DtValidite > DtClot then exit;
  result := 0;
  i := 0;
  while not ((DTValidite >= DtDeb) and (DTValidite <= DtFin)) do
  begin
    i := i + 1;
    if i > 50 then exit; // pour ne pas boucler au cas où....
    result := result + 1;
    DtFin := DtDeb - 1;
    Decodedate(DTFin, aa, mm, jj);
    DtDeb := PGImportEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT21 } // PT25
  end;
end;

function ImportRendNoDossier(): string;
begin
{ if (V_PGI_Env = nil) or (V_PGI_Env.NoDossier = '') then result := '000000'
  else result := V_PGI_Env.NoDossier; }
  if V_PGI.NoDossier <> '' then result := V_PGI.NoDossier
  else result := '000000';
end;

function PgImpnotVide(T: tob; const MulNiveau: boolean): boolean;
var
  tp: tob;
begin
  result := false;
  if t = nil then
    exit;
  tp := t.findfirst([''], [''], MulNiveau);
  result := (tp <> nil);
end;

procedure TOF_PGIMPORTFIC.OnArgument(Arguments: string);
var
  NomFic, FicRapport: ThEdit;
  {$IFNDEF EAGLCLIENT}
  bImprime : TToolbarbutton97;
  {$ENDIF !EAGLCLIENT}
  BValider, ConsultErreur: TToolbarbutton97;
  st, LeNomFic, LeMode: string;
begin
  //  inherited;
  MatSouche := -999;
  TopCreat := FALSE ;
  TFVierge(Ecran).Retour := 'NON'; // PT24 Initialisation avec un code retour
  st := Arguments;
  if st <> '' then
  begin
    LeNomFic := ReadTokenSt(St);
    LeMode := ReadTokenSt(St);
    if not (ctxAffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte) and not (ctxBTP in V_PGI.PGIContexte) then
      Ecran.Caption := Ecran.Caption + ' NetExpert';
    PgImportDef(LeNomFic, LeMode);
  end
  else
  begin
    NomFic := THedit(GetControl('EDTFICIMPORT'));
    taille := 0;
    if Nomfic <> nil then
    begin
      NomFic.OnChange := FicChange;
      Nomfic.DataType := 'OPENFILE(*.txt;*.*)'
    end;
    FicRapport := THedit(GetControl('FICRAPPORT'));
    if FicRapport <> nil then
    begin
      FicRapport.OnExit := FicChange;
      FicRapport.DataType := 'OPENFILE(*.txt;*.*)';
    end;
  end; // PT24
  BValider := TToolbarbutton97(GetControl('BVALIDER'));
  if BValider <> nil then
    BValider.OnClick := ControleFichierImport;

  {$IFNDEF EAGLCLIENT}
  BImprime := ttoolbarbutton97(getcontrol('BIMPRIMER'));
  if Bimprime <> nil then
    Bimprime.Onclick := ImprimeClick;
  {$ENDIF !EAGLCLIENT}

  Salarieprec := '';
  RapportErreur := THRichEditOLE(getcontrol('LISTEERREUR'));

  MarqueDebut := '***DEBUT***';
  MarqueFin := '***FIN***';
  Sep := THValComboBox(getcontrol('SEPARATEUR'));
  if (Sep <> nil) and (TypImport = '') then // PT24
  begin
    Sep.value := 'TAB';
    // PT18   10/09/2003 PH Controle du séparateur du fichier import et creation automatique fichier rapport
    Sep.OnExit := SepChange;
    Separateur := #9;
  end;
  ConsultErreur := TToolbarbutton97(getcontrol('CONSULTERREUR'));
  if ConsultErreur <> nil then
    ConsultErreur.Onclick := ConsultErreurClick;
  //  PT24  end;
  // DEB PT27
  Tob_Etab := TOB.Create('ETABCOMPL', nil, -1);
  Tob_Etab.LoadDetailDB('ETABCOMPL', '', '', nil, FALSE);
  //  FIN PT27
end;

procedure TOF_PGIMPORTFIC.ConsultErreurClick(Sender: Tobject);
var
  FicRapport: THedit;
  F: TextFile;
  S: string;
  TPC: TPageControl;
  T: TTabSheet;
  i: integer;
  okok: Boolean;
begin
  if RapportErreur = nil then exit;
  FicRapport := THedit(GetControl('FICRAPPORT'));
  if FicRapport = nil then exit;
  RapportErreur.Clear;
  if FicRapport.Text = '' then exit;
  if not FileExists(FicRapport.text) then exit;

  valideEcran(false);
  AssignFile(F, FicRapport.Text);
  Reset(F);
  i := 0;
  InitMoveProgressForm(nil, 'Chargement du fichier d''erreurs à l''écran', 'Veuillez patienter SVP ...', taille, TRUE, TRUE);
  while not eof(F) do
  begin
    i := i + 1;
    Readln(F, S);
    okok := MoveCurProgressForm(IntToStr(i));
    if not okok then
    begin
      closeFile(F);
      FiniMoveProgressForm;
    end;
    RapportErreur.lines.Add(S);
  end;

  closeFile(F);
  ValideEcran(true);
  FiniMoveProgressForm;

  T := TTabSheet(GetControl('TT2'));
  if T <> nil then
  begin
    TPC := TPageControl(getcontrol('PAGECONTROL1'));
    if TPC <> nil then
      TPC.ActivePage := T;
  end;
end;

procedure TOF_PGIMPORTFIC.valideEcran(Top: boolean);
begin
  SetControlEnabled('EDTFICIMPORT', top);
  SetControlEnabled('FICRAPPORT', top);
  SetControlEnabled('SEPARATEUR', top);
  SetControlEnabled('BVALIDER', top);
  SetControlEnabled('BFERME', top);
  SetControlEnabled('CONSULTERREUR', top);
end;

procedure TOF_PGIMPORTFIC.ControleFichierImport(Sender: Tobject);
var
  NomFic, FicRapport: ThEdit;
  S, st: string;
  NoErreur, i, ll: integer;
  ExerPerEncours, DebPer, FinPer: string;
  Q: TQuery;
  LeFic, LeRapport: string;
begin
  ExisteErreur := false;
  AbsExclue := false; // PT22
  MarqueFinOk := false;
  i := 0;
  if TypImport <> 'AUTO' then
  begin
    NomFic := THedit(GetControl('EDTFICIMPORT'));
    if NomFic = nil then exit;
    if NomFic.Text = '' then
    begin
      PGIBox('Fichier d''import obligatoire', 'Import Fichier');
      exit;
    end;
    if not FileExists(NomFic.text) then
    begin
      PGIBox('Fichier d''import inexistant', 'Import Fichier');
      exit;
    end;
    FicRapport := THedit(GetControl('FICRAPPORT'));
    if FicRapport = nil then exit;
    if FicRapport.Text = '' then
    begin
      PGIBox('Fichier de rapport obligatoire', 'Import Fichier');
      exit;
    end;
    if not FileExists(FicRapport.text) then
    begin
      // PT18    PGIBox('Fichier de rapport inexistant', 'Import Fichier'); exit ;
    end;
    valideEcran(false);
    LeFic := NomFic.Text;
    LeRapport := FicRapport.Text;
  end
  else
  begin
    LeFic := NomDuFichier;
    ll := Pos('.NSV', LeFic);
    if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) or (ctxBTP in V_PGI.PGIContexte) then
      LeRapport:=ExtractFilePath(NomDuFichier) + 'GenerationPaie.log' //AB-20050105-Import Gestion Affaire
    else
      LeRapport := Copy(LeFic, 1, ll - 1) + '.log'; // nom du fichier de rapport = .log au lieu de txt
    if not FileExists(LeRapport) then
    begin
      AssignFile(FR, LeRapport); // PT24
{$I-}
      ReWrite(FR);
{$I+}
      if IoResult <> 0 then
      begin
        PGIBox('Fichier rapport inaccessible : ' + LeFic, 'Abandon du traitement');
        Exit;
      end;
      closeFile(FR);
    end;
  end;
  AssignFile(F, LeFic);
  Reset(F);
  AssignFile(FR, LeRapport);
  Reset(FR);
  if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) or (ctxBTP in V_PGI.PGIContexte) then
    Append(FR)          //AB-20050105-Import Gestion d'affaire
  else
    ReWrite(FR);
  liberreur := true;
  S := '';
  while ((not IsMarqueDebut(s)) and (not eof(F))) do
  begin
    Readln(F, S);
  end;
  if eof(F) then
  begin
    Erreur(s, S, 5);
    CloseFile(F);
    CloseFile(FR);
    Exit;
  end;
  InitMoveProgressForm(nil, 'Contrôle du fichier d''import', 'Veuillez patienter SVP ...', taille, False, TRUE);

  if not Eof(F) then
  begin
    i := i + 1;
    MoveCurProgressForm(IntToStr(i));
    Readln(F, S);
    Noerreur := ControleEnrDossier(S, S);
    if NoErreur <> 0 then
      Erreur(s, S, NoErreur);
  end;
  ExerPerEncours := '';
  DebPer := '';
  FinPer := '';
  TFinPer := 0;
  if not ImportRendPeriodeEnCours(ExerPerEncours, DebPer, FinPer) then
    Erreur('', '', 7);
  if FinPer <> '' then
    if Isvaliddate(FinPer) then
      TFinPer := strtodate(FinPer) else
      Erreur('', '', 7);
  st := 'SELECT PSA_SALARIE,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SUSPENSIONPAIE,PSA_ETABLISSEMENT, ' +
    'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL,PSA_CONGESPAYES,' + { PT22 }
  'ETB_DATECLOTURECPN,ETB_CONGESPAYES FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT = ETB_ETABLISSEMENT';
  Q := opensql(st, true);
  if (not Q.eof) then
  begin
    Tob_sal := TOB.Create('Table des Salariés', nil, -1);
    if tob_sal <> nil then
      Tob_sal.loaddetaildb('INFOS SALARIES', '', '', Q, false, false, -1, 0);
  end;
  Ferme(Q);
  st := 'SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE <>""';
  Q := opensql(st, true);
  if not Q.eof then
  begin
    Tob_rubrique := TOB.Create('Table des rubriques', nil, -1);
    if tob_rubrique <> nil then
      Tob_rubrique.loaddetaildb('REMUNERATION', '', '', Q, false, false, -1, 0);
  end;
  Ferme(Q);
  //   PT2 09/10/01 PH Rajout traitement ligne Analytique
  if GetParamSoc('SO_PGANALYTIQUE') then // PT24
  begin
    if VH^.LiaisonY2ViaShare then
    begin
      st := 'SELECT CSP_AXE,CSP_SECTION FROM CSECTION ORDER BY CSP_AXE,CSP_SECTION';
      Q := opensql(st, true);
      if not Q.eof then
      begin
        Tob_section := TOB.Create('Table des sections', nil, -1);
        if tob_section <> nil then
          Tob_section.loaddetaildb('CSECTION', '', '', Q, false, false, -1, 0);
      end;
      Ferme(Q);
    end else
    begin
      st := 'SELECT S_AXE,S_SECTION FROM SECTION ORDER BY S_AXE,S_SECTION';
      Q := opensql(st, true);
      if not Q.eof then
      begin
        Tob_section := TOB.Create('Table des sections', nil, -1);
        if tob_section <> nil then
          Tob_section.loaddetaildb('SECTION', '', '', Q, false, false, -1, 0);
      end;
      Ferme(Q);
    end;
  end;
  // FIN PT2
  while not Eof(F) do //Deuxieme passage
  begin
    i := i + 1;
    MoveCurProgressForm(IntToStr(i));
    Readln(F, S);
    if MarqueFinOk then
    begin
      Erreur_10 := true;
      Erreur(s, S, 10);
    end;
    St := S;
    St := readtokenPipe(St, Separateur);
    if St = MarqueFin then
    begin
      MarqueFinOk := true;
      continue
    end;
    Traiteligne(S, S);
  end;
  if not MarqueFinOk then Erreur('', '', 15);
  FiniMoveProgressForm;

  if not ExisteErreur then
  begin
    Reset(F);
    ImporteFichier();
    TFVierge(Ecran).Retour := ''; // PT24
    if RapportErreur <> nil then
    begin
      // PT22   26/05/2004 PH V_50 Non traitement des absences portant sur des périodes closes mias ne provoquant pas de rejet du fichier
      if not AbsExclue then Writeln(FR, 'VOTRE IMPORT S''EST BIEN PASSE')
      else
      begin
        Writeln(FR, 'VOTRE IMPORT S''EST BIEN PASSE mais des lignes absences ont été exclues');
        PGIBox('Des lignes absences ont été exclues car elles sont hors période #13#10' +
          'Pour en consulter la liste, cliquez sur le bouton liste', 'Import fichier');
      end;
      // FIN PT22
            //RapportErreur.Lines.add('VOTRE IMPORT S''EST BIEN PASSE');
      PGIBox('Le fichier a été importé avec succès.#13#10' +
        ' Lancez la génération de bulletins pour l''intégration dans les paies.', 'Import fichier');
    end else
    if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) or (ctxBTP in V_PGI.PGIContexte) then
    begin
      Writeln(FR,'La génération dans la Paie s''est bien passée');
      PGIInfo('Les lignes d''activité sont transférées dans la paie avec succès.#13#10'+
      ' Vous pouvez faire la saisie ou la préparation des bulletins de paie.','Génération Paie');
    end;
  end
  else
  begin
    TFVierge(Ecran).Retour := 'ERR'; // PT24
    if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) or (ctxBTP in V_PGI.PGIContexte) then
    begin  //AB-20050105-Import Gestion Affaire
      PGIBox('Des erreurs détectées n''ont pas permis l''intégration dans la paie.#13#10'+
      ' Pour consulter les erreurs, cliquez sur bouton W pour ouvrir le fichier.log');
      Writeln(FR,'Des erreurs ont été détectées dans l''intégration dans la paie.');
    end else
      PGIBox('Des erreurs détectées n''ont pas permis l''import de votre fichier.#13#10' +
        ' Pour consulter la liste d''erreurs, cliquez sur le bouton liste', 'Import de fichier');
  end;
  FermeTout;
  valideEcran(true);
end;

function TOF_PGIMPORTFIC.IsMarqueDebut(S: string): boolean;
begin
  result := true;
  if S = MarqueDebut then exit;
  S := readtokenPipe(S, Separateur);
  if S = MarqueDebut then exit
  else result := false;
end;

function TOF_PGIMPORTFIC.IsMarqueFin(S: string): boolean;
begin
  result := true;
  if S = MarqueFin then exit;
  S := readtokenPipe(S, Separateur);
  if S = MarqueFin then exit
  else result := false;
end;

procedure TOF_PGIMPORTFIC.FermeTout;
begin
  CloseFile(F);
  CloseFile(FR);
  if PgImpnotVide(Tob_sal, true) then
    Tob_sal.Free;
  if PgImpnotVide(Tob_rubrique, true) then
    Tob_rubrique.Free;
  //   PT2 09/10/01 PH Rajout traitement ligne Analytique
  if Tob_section <> nil then
  begin
    Tob_section.free;
    tob_section := nil;
  end;
  // DEB PT27
  if Tob_Etab <> nil then
  begin
    Tob_Etab.free;
    Tob_Etab := nil;
  end;
  if Tob_SalATT <> nil then
  begin
    Tob_SalATT.free;
    Tob_SalATT := nil;
  end;
  // FIN PT27
end;

procedure TOF_PGIMPORTFIC.ImporteFichier;
var
  i, ii: integer;
  LaNature, MesAxes: string;
  S, CodeEnr, st, tempo, CodeS, salP: string;
  Tob_S, T: tob;
  Q: TQuery;
  DD, DF: TDateTime;
begin
  DD := iDate1900;
  DF := iDate2099;
  SalP := '';
  st := 'SELECT PCN_SALARIE,PCN_TYPEMVT,MAX(PCN_ORDRE) AS NUMORDRE FROM ABSENCESALARIE GROUP BY PCN_SALARIE,PCN_TYPEMVT ORDER BY PCN_SALARIE';
  Q := Opensql(st, true);
  //  if Q.eof then
  //     begin  Ferme(Q);   exit;     end;
  Tob_OrdreABS := Tob.create('Table des ordre absences maxi', nil, -1);
  Tob_ordreAbs.loaddetaildb('LES ABSENCESALARIES', '', '', Q, false);
  Ferme(Q);
  if Tob_ordreABS = nil then
  begin
    exit;
  end;

  st := 'SELECT PSD_SALARIE,MAX(PSD_ORDRE) AS PSD_ORDRE FROM HISTOSAISRUB GROUP BY PSD_SALARIE';
  Q := Opensql(st, true);
  //  if Q.eof then
  //     begin     Ferme(Q);   exit;     end;
  Tob_OrdreHSR := tob.create('Table des ordre maxi', nil, -1);
  Tob_ordreHSR.addchampsup('PSD_SALARIE', True);
  Tob_ordreHSR.addchampsup('PSD_ORDRE', True);
  Tob_ordreHSR.loaddetaildb('HISTOSAISRUB', '', '', Q, false);
  Ferme(Q);
  if Tob_ordreHSR = nil then
  begin
    exit;
  end;

  InitMoveProgressForm(nil, 'Ecriture de la base', 'Veuillez patienter SVP ...', taille, False, TRUE);

  Tob_S := Tob.create('Tob lignes import', nil, -1);
  Tob_Abs := tob.create('ABSENCES A INSERER', Tob_S, -1);
  Tob_HSR := tob.create('HISTOSAISRUB A INSERER', Tob_S, -1);
  // PT2 09/10/01 PH Rajout traitement ligne Analytique
  if GetParamSoc('SO_PGANALYTIQUE') then // PT24
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
  begin
    Tob_PaieVentil := TOB.Create('Les Ventils Analytiques', nil, -1);
    Tob_Ventil := TOB.Create('Les Préventils Analytiques', nil, -1);
  end;
  i := 0;
  while not Eof(F) do //Deuxieme passage
  begin
    i := i + 1;
    MoveCurProgressForm(IntToStr(i));
    Readln(F, S);
    if isMarqueDebut(S) then
      continue;
    if isMarqueFin(S) then
      break;
    CodeEnr := readtokenPipe(S, Separateur);
    tempo := S;
    if CodeEnr = '000' then
      continue;
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
    // DEB PT27 Création automatique des salariés en attente
    if (Tob_SalATT <> nil) and (Tob_SalATT.Detail.Count - 1 >= 0) then ControleEnregSAL('', '', '', TRUE); // Creation des salariés
    if CodeEnr = 'SAL' then continue;
    // FIN PT27
    // DEB PT27 Rajout ventilation type salarié uniquement valables dans le mois
    if (CodeEnr = 'VEN') or (CodeEnr = 'VSA') then
    begin
      CodeS := readtokenPipe(tempo, Separateur);
      if CodeS <> SalP then
      begin
        if (CodeEnr = 'VEN') then st := 'Salarie ' + CodeS + ' : enregistrement(s) pré-ventilations analytiques traité(s)'
        else st := 'Salarie ' + CodeS + ' : enregistrement(s) pré-ventilations analytiques salarié pour la période traité(s)';
        Anomalie(st, st, 98);
        SalP := CodeS;
      end;
      if (CodeEnr = 'VEN') then ControleEnregVEN(CodeEnr, S, S, True)
      else ControleEnregVSA(CodeEnr, S, S, True);
      // FIN PT27
    end
    else
    begin
      // FIN PT13
      if CodeEnr = 'MAB' then
      begin // Traitement ligne absences
        CodeS := readtokenPipe(tempo, Separateur);
        if CodeS <> SalP then
        begin
          st := 'Salarie ' + CodeS + ' : enregistrement(s) Absence traité(s)';
          Anomalie(st, st, 98);
          SalP := CodeS;
        end;
        ControleEnregMAB(S, S, True);
      end // Fin traitement ligne absences
      else
      begin
        CodeS := readtokenPipe(tempo, Separateur);
        if CodeS <> SalP then
        begin
          if CodeEnr <> 'ANA' then st := 'Salarie ' + CodeS + ' : enregistrement(s) MHE ou MLB ou MFP traité(s)'
          else st := 'Salarie ' + CodeS + ' : enregistrement(s) Analytique(s) traité(s)';
          Anomalie(st, st, 98);
          SalP := CodeS;
        end;

        if CodeEnr <> 'ANA' then ControleEnregMHE_MFP_MLB(CodeEnr, S, S, True)
        else ControleEnregANA(CodeEnr, S, S, True);
      end;
    end;
  end; // Fin boucle lecture fichier import
  try
    BeginTrans;

    if Tob_HSR <> nil then
    begin
      Tob_HSR.SetAllModifie(TRUE);
      Tob_HSR.InsertOrUpdateDB(false);
    end;
    if Tob_ABS <> nil then
    begin
      Tob_ABS.SetAllModifie(TRUE);
      Tob_ABS.InsertOrUpdateDB(false);
    end;
    if Tob_PaieVentil <> nil then
    begin
      // PT17   08/09/2003 PH Destruction des mouvements ANA
      SalP := '';
      T := Tob_PaieVentil.FindFirst([''], [''], FALSE);
      // PT20 Suppression des ventilations analytiques salarie sur axe dans le fichier d'import
      MesAxes := '-----';
      while T <> nil do
      begin
        CodeS := T.GetValue('PAV_COMPTE');
        DD := T.GetValue('PAV_DATEDEBUT');
        DF := T.GetValue('PAV_DATEFIN');
        if CodeS <> SalP then
        begin
          st := 'DELETE FROM PAIEVENTIL WHERE (PAV_NATURE LIKE "PG%" OR PAV_NATURE LIKE "VS%") AND PAV_COMPTE="' + SalP + '"' +
            ' AND PAV_DATEDEBUT >= "' + UsDateTime(DD) + '" AND PAV_DATEFIN <= "' + UsDateTime(DF) + '"';
          ExecuteSQL(St);
          SalP := CodeS;
          MesAxes := '-----';
        end;
        // PT23 V_NATURE remplacé par PAV_NATURE
        ii := ValeurI(Copy(T.GetValue('PAV_NATURE'), 3, 1));
        MesAxes[ii] := 'X';
        T := Tob_PaieVentil.FindNext([''], [''], FALSE);
      end;
      // FIN PT17
      // Traitement du dernier salarié
      st := 'DELETE FROM PAIEVENTIL WHERE (PAV_NATURE LIKE "PG%" OR PAV_NATURE LIKE "VS%") AND PAV_COMPTE="' + SalP + '"' +
        ' AND PAV_DATEDEBUT >= "' + UsDateTime(DD) + '" AND PAV_DATEFIN <= "' + UsDateTime(DF) + '"';
      ExecuteSQL(St);
      // FIN PT20
      Tob_PaieVentil.SetAllModifie(TRUE);
      Tob_PaieVentil.InsertOrUpdateDB(false);
    end;
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
    if Tob_Ventil <> nil then
    begin
      // PT15   25/08/2003 PH Correction destruction des pré-ventilations analytiques
      SalP := '';
      T := Tob_Ventil.FindFirst([''], [''], FALSE);
      // PT20 Suppression des ventilations analytiques salarie sur axe dans le fichier d'import
      MesAxes := '-----';
      while T <> nil do
      begin
        CodeS := T.GetValue('V_COMPTE');
        if CodeS <> SalP then
        begin
          // PT15   25/08/2003 PH Correction destruction des pré-ventilations analytiques
          for ii := 1 to 5 do
          begin
            if MesAxes[ii] = 'X' then
            begin //  on traite tjrs le salarié en cours cad celui qui traité avt rupture de matricule
              LaNature := 'SA' + IntToStr(ii);
              ExecuteSQL('DELETE FROM VENTIL WHERE V_NATURE LIKE "' + LaNature + '" AND V_COMPTE="' + SalP + '"');
            end;
          end;
          SalP := CodeS;
          MesAxes := '-----';
          // FIN PT20
        end;
        ii := ValeurI(Copy(T.GetValue('V_NATURE'), 3, 1));
        MesAxes[ii] := 'X'; // DEB PT20
        T := Tob_Ventil.FindNext([''], [''], FALSE);
      end;
      // PT20 Suppression des ventilations analytiques salarie sur axe dans le fichier d'import
      for ii := 1 to 5 do
      begin // Pour traiter le cas du dernier salarié
        if MesAxes[ii] = 'X' then
        begin
          LaNature := 'SA' + IntToStr(ii);
          ExecuteSQL('DELETE FROM VENTIL WHERE V_NATURE LIKE "' + LaNature + '" AND V_COMPTE="' + SalP + '"');
        end;
      end;
      // FIN PT20
      Tob_Ventil.SetAllModifie(TRUE);
      Tob_Ventil.InsertOrUpdateDB(false);
    end;
    // FIN PT13
    // FIN PT2
    FiniMoveProgressForm();
    if Tob_ordreHSR <> nil then Tob_ordreHSR.free;
    if Tob_ordreABS <> nil then Tob_ordreABS.free;
    if Tob_S <> nil then Tob_S.free;
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
    if Tob_Ventil <> nil then Tob_Ventil.free;
    if Tob_PaieVentil <> nil then Tob_PaieVentil.free;
    // FIN PT13

    CommitTrans;
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de l''écriture dans la base', 'Ecriture fichier import');
  end;

end;
// ligne absence
{function TOF_PGIMPORTFIC.EcritligneMAB(S:string,T:tob);
var
begin

end;}

procedure TOF_PGIMPORTFIC.Traiteligne(chaine, S: string);
var
  CodeEnr: string;
begin

  CodeEnr := readtokenPipe(S, Separateur);
  if (CodeEnr = 'MHE') or (CodeEnr = 'MFP') or (CodeEnr = 'MLB') then
    ControleEnregMHE_MFP_MLB(CodeEnr, chaine, S, false)
  else
    if CodeEnr = 'MAB' then ControleEnregMAB(chaine, S, false)
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
  else
    if CodeEnr = 'VEN' then ControleEnregVEN(CodeEnr, chaine, S, false)
  else
    // PT15   25/08/2003 PH Correction prise en compte traitement analytique ANA
    if CodeEnr = 'ANA' then ControleEnregANA(CodeEnr, chaine, S, false)
  else
    // DEB PT27
    if CodeEnr = 'VSA' then ControleEnregVSA(CodeEnr, chaine, S, false)
  else
    if CodeEnr = 'SAL' then ControleEnregSAL(CodeEnr, chaine, S, false)
    // FIN PT27
  else Erreur('', S, 20);
end;

function TOF_PGIMPORTFIC.ControleEnrDossier(chaine, S: string): integer;
var
  NoDoss, DossEnCours, ent: string;
  DatedebES, DateFinES, DebExer, FinExer: TDateTime;
  MoisE, AnneeE, ComboExer: string;
  aa, mm, jj: word;
  rep: Integer;
begin
  ent := readtokenPipe(S, Separateur);
  if ent <> '000' then Erreur(chaine, DossEnCours, 220);
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  DossEnCours := ImportRendNoDossier();
  NoDoss := readtokenPipe(S, Separateur);
  if (DossEnCours <> NoDoss) and (DossEnCours <> '000000') and (DossEnCours <> '') then
    Erreur(chaine, DossEnCours, 230);
  // format US vers format Français
  {DEB PT12 transtypage incorrect si S=''}
  DecodeDate(idate1900, aa, mm, jj);
  if S <> '' then jj := strtoint(readtokenPipe(S, '/'));
  if S <> '' then mm := strtoint(readtokenPipe(S, '/'));
  if S <> '' then aa := strtoint(readtokenPipe(S, Separateur));
  DatedebES := PGImportEncodeDateBissextile(aa, mm, jj); { PT21 } // PT25
  DecodeDate(idate1900, aa, mm, jj);
  if S <> '' then jj := strtoint(readtokenPipe(S, '/'));
  if S <> '' then mm := strtoint(readtokenPipe(S, '/'));
  if S <> '' then aa := strtoint(readtokenPipe(S, Separateur));
  DatefinES := PGImportEncodeDateBissextile(aa, mm, jj); { PT21 } // PT25
  {FIN PT12}

  MoisE := '';
  AnneeE := '';
  ComboExer := '';
  DebExer := 0;
  FinExer := 0;
  ImportRendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);
  if ((datedebES <> DebExer) or (dateFinES <> FinExer)) and ((datedebES > Idate1900) and (dateFinES > Idate1900)) then
  begin
    Rep := PGIAsk('Les dates d''exercice social du fichier sont différentes de celles de votre entreprise#13#10Voulez-vous quand même prendre en compte votre fichier?',
      'Incohérence des dates d''exercice social.');
    if rep <> mrYes then Erreur(chaine, DossEnCours, 240);
  end;
  result := 0;
end;

procedure TOF_PGIMPORTFIC.ControleEnregMHE_MFP_MLB(TypeEnr, chaine, S: string; Ecrit: boolean);
var
  Codesalarie, DateDebutperiode, DateFinperiode, SNoOrdre: string;
  Rubrique, typeAlim, libelle, sBase, staux, sCoefficient, sMontant, ll: string;
  TDateDeb, TDateFin: tdatetime;
  Base, taux, coeff, montant: double;
  Sal, TS, T: tob;
  Ordre, TopCloture: integer;
begin
  libErreur := false;
  libanomalie := false;
  Ordre := 1;
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if CodeSalarie <> Salarieprec then
  begin
    ControleSalarie(TypeEnr, chaine, S, CodeSalarie, Ecrit);
  end;
  Salarieprec := Codesalarie;
  DateDebutperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DatedebutPeriode) then
    Erreur(chaine, DatedebutPeriode, 30);
  DateFinperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateFinPeriode) then
    Erreur(chaine, DateFinPeriode, 35);

  TDateDeb := 0;
  TDateFin := 0;

  if (Isvaliddate(DatedebutPeriode)) and (Isvaliddate(DatefinPeriode)) then
  begin
    TDateDeb := strtodate(DatedebutPeriode);
    TDateFin := strtodate(DatefinPeriode);
  end;

  if TDateFin < TDateDeb then
    Erreur(chaine, DateDebutPeriode + '//' + DateFinPeriode, 40);

  Topcloture := IsPeriodecloture(TDatedeb, TDateFin);
  // PT22   26/05/2004 PH V_50 Non traitement des absences portant sur des périodes closes mias ne provoquant pas de rejet du fichier
  if not Ecrit then ll := 'NF' else ll := '';
  if TopCloture = 1 then Erreur(Chaine, '', 44)
  else
    if TopCloture = 2 then Erreur(Chaine, '', 47, ll)
  else
    if TopCloture = 3 then Erreur(Chaine, '', 48, ll);
  // FIN PT22
  if (DateSortie > 10) and (DateSortie < TDateDeb) then
    Erreur(Chaine, '', 46);

  SNoOrdre := readtokenPipe(S, Separateur);
  if SnoOrdre = '' then
    Erreur(chaine, SNoOrdre, 45);

  Rubrique := readtokenPipe(S, Separateur);
  if TypeEnr <> 'MLB' then RubPrec := Rubrique;

  TypeBase := '';
  TypeTaux := '';
  TypeCoeff := '';
  TypeMontant := '';
  if (TypeEnr = 'MHE') or (typeEnr = 'MFP') then
  begin
    if not ExisteRubrique(Rubrique) then
      Erreur(chaine, Rubrique, 50);
  end
  else
  begin
    if Rubrique <> rubPrec then
      Erreur(Chaine, Rubrique, 52);
  end;
  libelle := readtokenPipe(S, Separateur);

  TypeAlim := readtokenPipe(S, Separateur);
  if (TypeAlim = '') and (TypeEnr <> 'MLB') then
    Erreur(chaine, Rubrique, 55);
  // DEB PT26
  if ((TypeBase <> '00') and (TypeBase <> '01') and (pos('B', typeAlim) > 0) or
    (((typeBase = '00') or (typeBase = '01')) and (pos('B', typeAlim) = 0)) or
    ((TypeTaux <> '00') and (TypeTaux <> '01') and (pos('T', typeAlim) > 0)) or
    (((typeTaux = '00') or (typeTaux = '01')) and (pos('T', typeAlim) = 0)) or
    ((TypeCoeff <> '00') and (TypeCoeff <> '01') and (pos('C', typeAlim) > 0)) or
    (((typeCoeff = '00') or (typeCoeff = '01')) and (pos('C', typeAlim) = 0)) or
    ((TypeMontant <> '00') and (TypeMontant <> '01') and (pos('M', typeAlim) > 0)) or
    (((typeMontant = '00') or (typeMontant = '01')) and (pos('M', typeAlim) = 0))) and (TypeEnr <> 'MLB') then
    Erreur(chaine, Rubrique, 57);
  // FIN PT26
  sBase := readtokenPipe(S, Separateur);
  if pos('B', typeAlim) > 0
    then
  begin
    if sbase = '' then
      Erreur(chaine, sBase, 60)
    else
      if not Isnumeric(sBase) then
      Erreur(chaine, sBase, 60);
  end
  else
  begin
    if sBase <> '' then
      Erreur(chaine, sBase, 65);
  end;

  staux := readtokenPipe(S, Separateur);
  if pos('T', typeAlim) > 0
    then
  begin
    if staux = '' then
      Erreur(chaine, staux, 70)
    else
      if not Isnumeric(staux) then
      Erreur(chaine, staux, 70);
  end
  else
  begin
    if staux <> '' then
      Erreur(chaine, staux, 75);
  end;


  sCoefficient := readtokenPipe(S, Separateur);
  if pos('C', typeAlim) > 0
    then
  begin
    if sCoefficient = '' then
      Erreur(chaine, sCoefficient, 80)
    else
      if not Isnumeric(sCoefficient) then
      Erreur(chaine, sCoefficient, 80);
  end
  else
  begin
    if sCoefficient <> '' then
      Erreur(chaine, sCoefficient, 85);
  end;

  sMontant := readtokenPipe(S, Separateur);
  if pos('M', typeAlim) > 0
    then
  begin
    if sMontant = '' then
      Erreur(chaine, sMontant, 90)
    else
      if not Isnumeric(sMontant) then
      Erreur(chaine, sMontant, 90);
  end
  else
  begin
    if sMontant <> '' then
      Erreur(chaine, sMontant, 95);
  end;

  if not Ecrit then exit;
  Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
  if Sal = nil then exit;
  TS := Tob_OrdreHSR.findfirst(['PSD_SALARIE'], [CodeSalarie], true);
  if TS <> nil then
  begin
    Ordre := TS.getvalue('PSD_ORDRE');
    TS.putvalue(('PSD_ORDRE'), Ordre + 1);
  end
  else
  begin
    TS := Tob.create('HISTOSAISRUB', Tob_OrdreHSR, -1);
    if Ts <> nil then
    begin
      Ts.putvalue('PSD_SALARIE', CodeSalarie);
      Ts.putValue('PSD_Ordre', 1);
    end;
  end;
  if not Ecrit then exit;
  Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
  if (Sal = nil) then exit;
  T := Tob.create('HISTOSAISRUB', TOB_HSR, -1);
  if T = nil then exit;
  TS := Tob_OrdreHSR.findfirst(['PSD_SALARIE'], [CodeSalarie], true);
  if TS <> nil then
  begin
    Ordre := TS.getvalue('PSD_ORDRE');
    TS.putvalue(('PSD_ORDRE'), Ordre + 1);
  end
  else
  begin
    TS := Tob.create('HISTOSAISRUB', Tob_OrdreHSR, -1);
    if Ts <> nil then
    begin
      Ts.putvalue('PSD_SALARIE', CodeSalarie);
      Ts.putValue('PSD_Ordre', 1);
    end;
  end;
  Base := 0;
  Taux := 0;
  Coeff := 0;
  Montant := 0;
  T.PutValue('PSD_ORIGINEMVT', TypeEnr);
  T.PutValue('PSD_SALARIE', CodeSalarie);
  T.PutValue('PSD_DATEDEBUT', TDateDeb);
  T.PutValue('PSD_DATEFIN', TDateFin);
  T.PutValue('PSD_RUBRIQUE', Rubrique);
  T.PutValue('PSD_ORDRE', Ordre);
  T.PutValue('PSD_LIBELLE', Libelle);
  T.PutValue('PSD_RIBSALAIRE', '');
  T.PutValue('PSD_BANQUEEMIS', '');
  T.PutValue('PSD_TOPREGLE', '-');
  T.PutValue('PSD_TYPALIMPAIE', typeAlim);
  T.PutValue('PSD_ETABLISSEMENT', sal.getvalue('PSA_ETABLISSEMENT'));
  if pos('B', typeAlim) > 0 then Base := Valeur(sBase);
  if pos('T', typeAlim) > 0 then Taux := Valeur(sTaux);
  if pos('C', typeAlim) > 0 then Coeff := Valeur(sCoefficient);
  if pos('M', typeAlim) > 0 then Montant := Valeur(sMontant);

  T.PutValue('PSD_BASE', Base);
  T.PutValue('PSD_TAUX', Taux);
  T.PutValue('PSD_COEFF', Coeff);
  T.PutValue('PSD_MONTANT', Montant);
  T.PutValue('PSD_DATEINTEGRAT', 0);
  T.PutValue('PSD_DATECOMPT', 0);
  T.PutValue('PSD_AREPORTER', '');
  T.Putvalue('PSD_CONFIDENTIEL', sal.getvalue('PSA_CONFIDENTIEL'));

end;
//   PT2 09/10/01 PH Rajout traitement ligne Analytique Salarie/rubrique

procedure TOF_PGIMPORTFIC.ControleEnregANA(TypeEnr, chaine, S: string; Ecrit: boolean);
var
  Codesalarie, DateDebutperiode, DateFinperiode, SNoOrdre: string;
  Rubrique, Axe, Section, CodeAxe, sMontant, LibSect: string;
  TDateDeb, TDateFin: tdatetime;
  T: tob;
  TopCloture: integer;
begin
  libErreur := false;
  libanomalie := false;
  if not GetParamSoc('SO_PGANALYTIQUE') then Erreur(chaine, '', 101); // PT24
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if CodeSalarie <> Salarieprec then
  begin
    ControleSalarie(TypeEnr, chaine, S, CodeSalarie, Ecrit);
  end;
  Salarieprec := Codesalarie;
  DateDebutperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DatedebutPeriode) then
    Erreur(chaine, DatedebutPeriode, 30);
  DateFinperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateFinPeriode) then
    Erreur(chaine, DateFinPeriode, 35);
  TDateDeb := 0;
  TDateFin := 0;
  if (Isvaliddate(DatedebutPeriode)) and (Isvaliddate(DatefinPeriode)) then
  begin
    TDateDeb := strtodate(DatedebutPeriode);
    TDateFin := strtodate(DatefinPeriode);
  end;
  if TDateFin < TDateDeb then
    Erreur(chaine, DateDebutPeriode + '//' + DateFinPeriode, 40);
  Topcloture := IsPeriodecloture(TDatedeb, TDateFin);
  if TopCloture = 1 then
    Erreur(Chaine, '', 44);
  if TopCloture = 2 then
    Erreur(Chaine, '', 47);
  if TopCloture = 3 then
    Erreur(Chaine, '', 48);
  if (DateSortie > 10) and (DateSortie < TDateDeb) then
    Erreur(Chaine, '', 46);

  Rubrique := readtokenPipe(S, Separateur);
  RubPrec := Rubrique;
  if not ExisteRubrique(Rubrique) then Erreur(chaine, Rubrique, 50);
  if Rubrique <> rubPrec then Erreur(Chaine, Rubrique, 52);
  Axe := readtokenPipe(S, Separateur);
  if (Axe < 'A1') or (Axe > 'A5') then Erreur(chaine, Axe, 103);
  Section := readtokenPipe(S, Separateur);

  Smontant := readtokenPipe(S, Separateur);
  if sMontant = '' then
    Erreur(chaine, sMontant, 90)
  else
    if not Isnumeric(sMontant) then Erreur(chaine, sMontant, 90);
  SNoOrdre := readtokenPipe(S, Separateur);
  if SNoOrdre = '' then
    Erreur(chaine, SNoOrdre, 102)
  else
    if not Isnumeric(SNoOrdre) then Erreur(chaine, SNoOrdre, 102);

  // PT16   28/08/2003 PH Création des sections en récupérant le libellé de la section
  if S <> '' then LibSect := readtokenPipe(S, Separateur)
  else LibSect := '';
  if not ExisteSection(Axe, Section, LibSect) then Erreur(Chaine, Section, 100);
  // Fin PT16

  if not Ecrit then exit;
  T := Tob.create('PAIEVENTIL', TOB_Paieventil, -1);
  if T = nil then exit;
  CodeAxe := Copy(Axe, 2, 1);
  T.PutValue('PAV_NATURE', 'PG' + CodeAxe);
  T.PutValue('PAV_COMPTE', CodeSalarie + ';' + Rubrique);
  T.PutValue('PAV_DATEDEBUT', strtodate(DateDebutperiode));
  T.PutValue('PAV_DATEFIN', strtodate(DateFinperiode));
  T.PutValue('PAV_SECTION', Section);
  T.PutValue('PAV_TAUXMONTANT', Valeur(sMontant));
  T.PutValue('PAV_TAUXQTE1', 0);
  T.PutValue('PAV_TAUXQTE2', 0);
  T.PutValue('PAV_NUMEROVENTIL', StrToInt(SNoOrdre));
  T.PutValue('PAV_SOCIETE', GetParamSoc('SO_SOCIETE'));
  T.PutValue('PAV_MONTANT', 0);
  T.PutValue('PAV_SOUSPLAN1', '');
  T.PutValue('PAV_SOUSPLAN2', '');
  T.PutValue('PAV_SOUSPLAN3', '');
  T.PutValue('PAV_SOUSPLAN4', '');
  T.PutValue('PAV_SOUSPLAN5', '');
  T.PutValue('PAV_SOUSPLAN6', '');

end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Traitement d'une ligne du fichier import qui remplacera les
Suite ........ : ventilations analytiques salarié qui sont dans la table
Suite ........ : VENTIL.
Suite ........ :
Suite ........ : Elles correspondent  aux pré-ventilations analytiques.
Mots clefs ... :
*****************************************************************}
//PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés

procedure TOF_PGIMPORTFIC.ControleEnregVEN(TypeEnr, chaine, S: string; Ecrit: boolean);
var
  Codesalarie, SNoOrdre, LibSect: string;
  Axe, Section, CodeAxe, sMontant: string;
  T: tob;
begin
  libErreur := false;
  libanomalie := false;
  if not GetParamSoc('SO_PGANALYTIQUE') then Erreur(chaine, '', 101); // PT24
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if CodeSalarie <> Salarieprec then ControleSalarie(TypeEnr, chaine, S, CodeSalarie, Ecrit);
  Salarieprec := Codesalarie;
  Axe := readtokenPipe(S, Separateur);
  if (Axe < 'A1') or (Axe > 'A5') then Erreur(chaine, Axe, 103);
  Section := readtokenPipe(S, Separateur);
  Smontant := readtokenPipe(S, Separateur);
  if sMontant = '' then
    Erreur(chaine, sMontant, 90)
  else
    if not Isnumeric(sMontant) then Erreur(chaine, sMontant, 90);
  SNoOrdre := readtokenPipe(S, Separateur);
  if SNoOrdre = '' then
    Erreur(chaine, SNoOrdre, 102)
  else
    if not Isnumeric(SNoOrdre) then Erreur(chaine, SNoOrdre, 102);
  // PT16   28/08/2003 PH Création des sections en récupérant le libellé de la section
  if S <> '' then LibSect := readtokenPipe(S, Separateur)
  else LibSect := '';
  if not ExisteSection(Axe, Section, LibSect) then Erreur(Chaine, Section, 100);
  // Fin PT16
  if not Ecrit then exit;
  T := Tob.create('VENTIL', TOB_Ventil, -1);
  if T = nil then exit;
  CodeAxe := Copy(Axe, 2, 1);
  T.PutValue('V_NATURE', 'SA' + CodeAxe);
  T.PutValue('V_COMPTE', CodeSalarie);
  T.PutValue('V_SECTION', Section);
  T.PutValue('V_TAUXMONTANT', Valeur(sMontant));
  T.PutValue('V_TAUXQTE1', 0);
  T.PutValue('V_TAUXQTE2', 0);
  T.PutValue('V_NUMEROVENTIL', StrToInt(SNoOrdre));
  T.PutValue('V_SOCIETE', GetParamSoc('SO_SOCIETE'));
  T.PutValue('V_MONTANT', 0);
  T.PutValue('V_SOUSPLAN1', '');
  T.PutValue('V_SOUSPLAN2', '');
  T.PutValue('V_SOUSPLAN3', '');
  T.PutValue('V_SOUSPLAN4', '');
  T.PutValue('V_SOUSPLAN5', '');
  T.PutValue('V_SOUSPLAN6', '');
end;

// Test existance section axe analytique

function TOF_PGIMPORTFIC.ExisteSection(Axe, Section, LibSect: string): boolean;
var
  T, T1: tob;
  ll: Integer;
  SousSection : string;
begin
  result := false;
  T := Tob_section.findfirst(['S_AXE', 'S_SECTION'], [Axe, Section], true);
  if T = nil then
  begin
    // PT14   04/06/2003 PH Création automatique des sections analytiques lors de l'import si autorisé
    if GetParamSoc('SO_PGCREATIONSECTION') then // PT24
    begin
      if VH^.LiaisonY2ViaShare then
      begin
        SousSection := GetSousSectionUnique (Axe);
        T := TOB.Create('LES SECTIONS', nil, -1);
        T1 := TOB.Create('CSECTION', T, -1);
        T1.PutValue('CSP_SECTION', Section);
        if LibSect <> '' then
        begin
          T1.PutValue('CSP_LIBELLE', LibSect);
          ll := StrLen(PChar(LibSect));
          if ll > 17 then ll := 17;
          if ll > 0 then T1.PutValue('CSP_ABREGE', Copy(LibSect, 1, ll))
          else T1.PutValue('CSP_ABREGE', 'Section paie ');
        end
        else
        begin
          T1.PutValue('CSP_LIBELLE', 'Section paie ' + Section);
          T1.PutValue('CSP_ABREGE', 'Section paie ');
        end;
        T1.PutValue('CSP_SECTION',SousSection) ;
        T1.PutValue('CSP_SENS', 'M');
        T1.PutValue('CSP_SOLDEPROGRESSIF', 'X');
        T1.PutValue('CSP_AXE', Axe);
        // CEGID V9
        T1.PutValue('CSP_INVISIBLE','-') ;
        T.InsertDB(nil, FALSE);
        T1.ChangeParent(Tob_section, -1);
        Tob_section.Detail.Sort('CSP_AXE;CSP_SECTION');
        FreeAndNil(T);
        result := true;
      end else
      begin
        T := TOB.Create('LES SECTIONS', nil, -1);
        // PT15   25/08/2003 PH Correction creation automatique des sections analytiques
        T1 := TOB.Create('SECTION', T, -1);
        T1.PutValue('S_SECTION', Section);
        //PT16   28/08/2003 PH Création des sections en récupérant le libellé de la section
        if LibSect <> '' then
        begin
          T1.PutValue('S_LIBELLE', LibSect);
          ll := StrLen(PChar(LibSect));
          if ll > 17 then ll := 17;
          if ll > 0 then T1.PutValue('S_ABREGE', Copy(LibSect, 1, ll))
          else T1.PutValue('S_ABREGE', 'Section paie ');
        end
        else
        begin
          T1.PutValue('S_LIBELLE', 'Section paie ' + Section);
          T1.PutValue('S_ABREGE', 'Section paie ');
        end;
        // FIN PT16
        T1.PutValue('S_SENS', 'M');
        T1.PutValue('S_SOLDEPROGRESSIF', 'X');
        T1.PutValue('S_AXE', Axe);
        // CEGID V9
        T1.PutValue('S_INVISIBLE','-') ;
        // ---

        T.InsertDB(nil, FALSE);
        T1.ChangeParent(Tob_section, -1);
        Tob_section.Detail.Sort('S_AXE;S_SECTION');
        FreeAndNil(T);
        // PT15   25/08/2003 PH Creation section OK
        result := true;
      end;
    end
    else // AB 16/10/03
      Writeln(FR, 'Vous n''êtes pas autoriser à créer des sections analytiques dans les paramètres société comptable PAIE.');
    // FIN PT14
  end
  else result := true;
end;

// FIN PT2
// Fonction de récupération d'un mvt de type Absence

procedure TOF_PGIMPORTFIC.ControleEnregMAB(chaine, S: string; Ecrit: boolean);
var
  CodeSalarie, DateDebutAbs, DateFinAbs, st: string;
  TDateDeb, TDateFin, Dateclot: tdatetime;
  NbJ, NbH, MotifAbs, Libelle: string;
  Q: tquery;
  T, Sal: tob;
  Ordre, TopCloture: integer;
  LibCompl1, LibCompl2, Dj, Fj, ll: string;
  OkOk, SalSorti: Boolean;
  TAtt: TOB;
begin
  TAtt := nil;
  libErreur := false;
  libanomalie := false;
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if CodeSalarie <> Salarieprec then
  begin
    ControleSalarie('MAB', Chaine, S, CodeSalarie, Ecrit);
  end;
  Salarieprec := Codesalarie;
  DateDebutAbs := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateDebutAbs) then
    Erreur(chaine, DateDebutAbs, 110);
  DateFinAbs := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateFinAbs) then
    Erreur(chaine, DateFinAbs, 120);
  TDateDeb := 0;
  TDateFin := 0;
  if (Isvaliddate(DateDebutAbs)) and (Isvaliddate(DateFinAbs)) then
  begin
    TDateDeb := strtodate(DateDebutAbs);
    TDateFin := strtodate(DateFinAbs);
  end;
  if TDateFin < TDateDeb then
    Erreur(chaine, DateDebutAbs + '//' + DateFinAbs, 130);
  //  PT8 15/02/02 PH test pour les mvts de cp sur mois clos N et N+1
  Topcloture := IsPeriodecloture(TDatedeb, TDateFin, 'X');
  // PT22   26/05/2004 PH V_50 Non traitement des absences portant sur des périodes closes mias ne provoquant pas de rejet du fichier
  if not Ecrit then ll := 'NF' else ll := '';
  if TopCloture = 1 then Erreur(Chaine, '', 44)
  else
    if TopCloture = 2 then Erreur(Chaine, '', 47, ll)
  else
    if TopCloture = 3 then Erreur(Chaine, '', 48, ll);
  // FIN PT22
  if (DateSortie > 10) and (DateSortie < TDateDeb) then Erreur(Chaine, '', 46);
  //  PT11-1 02/12/2002 PH Regroupement lecture du fichier pour controle existance absence avec
  //  les bornes matin après midi
  NbJ := readtokenPipe(S, Separateur);
  if not isNumeric(Nbj) then
    if Nbj <> '' then Erreur(chaine, Nbj, 150);

  NbH := readtokenPipe(S, Separateur);
  if not isNumeric(NbH) then
    if ((NbH <> '') and (isNumeric(NbJ))) then
      Erreur(chaine, NbH, 160);

  MotifAbs := readtokenPipe(S, Separateur);
  Libelle := readtokenPipe(S, Separateur);
  // PT1 14/06/01 PH Rajout récup libellé complémentaires
  LibCompl1 := readtokenPipe(S, Separateur);
  LibCompl2 := readtokenPipe(S, Separateur);
  Dj := readtokenPipe(S, Separateur);
  Fj := readtokenPipe(S, Separateur);
  if Dj = '' then Dj := 'MAT';
  if Fj = '' then Fj := 'PAM';
  //  PT11 02/12/2002 PH Modification de la requete pour prendre en compte les demi journées
  st := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_TYPECONGE,PCN_TYPEMVT FROM ABSENCESALARIE ' +
    ' WHERE PCN_SALARIE = "' + CodeSalarie + '" AND PCN_DEBUTDJ="' + DJ + '" AND PCN_FINDJ="' + FJ + '"' +
    ' AND (PCN_TYPECONGE = "PRI" OR ' +
    ' (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-"))' +
    ' AND ((PCN_DATEDEBUTABS >="' + usdatetime(TDateDeb) + '" AND PCN_DATEDEBUTABS <= "' +
    usdatetime(TDateFin) + '")' +
    'OR (PCN_DATEFINABS >="' + usdatetime(TDateDeb) + '" AND PCN_DATEDEBUTABS <= "' +
    usdatetime(TDateFin) + '")' +
    'OR (PCN_DATEFINABS >="' + usdatetime(TDateFin) + '" AND PCN_DATEDEBUTABS <= "' +
    usdatetime(TDateDeb) + '")' +
    'OR(PCN_DATEFINABS <="' + usdatetime(TDateFin) + '" AND PCN_DATEDEBUTABS >= "' +
    usdatetime(TDateDeb) + '"))';
  // FIN PT11-1
  OkOk := TRUE;
  SalSorti := FALSE;
  Q := opensql(st, True);
  if not Q.eof then
  begin // Absence déjà existante
    Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
    // PT28 Rajout du test si Tob_SalATT <> NIL
    if Tob_SalATT <> nil then TAtt := Tob_SalATT.findfirst(['PIM_SALARIE'], [Codesalarie], true); // PT27
    //  PT11-2 02/12/2002 PH si les absences sont déjà saisies et que le salarié est sorti alors OK
    if (Sal = nil) and (TAtt = nil) then OkOk := FALSE // PT27
    else
    begin
      if (TAtt = nil) then
      begin
        if ((Sal.GetValue('PSA_DATESORTIE')) <> NULL) or (Sal.GetValue('PSA_DATESORTIE') > Idate1900)
          then SalSorti := TRUE else OkOk := FALSE;
      end;
    end;
    if not OkOk then Erreur(chaine, '', 140);
    // FIN PT11-2
  end;
  ferme(Q);

  Q := opensql('SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_MOTIFABSENCE ="'
    + Motifabs + '"', true);
  if Q.eof then Erreur(Chaine, MotifAbs, 170);
  ferme(Q);

  if not Ecrit then exit;
  //  PT11-2 02/12/2002 PH si les absences sont déjà saisies et que le salarié est sorti alors OK
  if OkOk and SalSorti then exit; // On ne traite pas la ligne du fichier d'import

  Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
  if Sal = nil then exit;
  // PT6 27/11/01 PH Si on ne gère pas les cp et que l'on a des lignes de cp,on perd les lignes
  if (Sal.getvalue('ETB_CONGESPAYES') <> 'X') and (Sal.getvalue('PSA_CONGESPAYES') <> 'X') and (MotifAbs = 'PRI') then { PT22 }
  begin
    Erreur(chaine, CodeSalarie, 27);
    exit;
  end;
  T := Tob.create('ABSENCESALARIE', TOB_ABS, -1);
  if T = nil then exit;

  //  Ordre := IncrementeSeqNoOrdre('CPA',codeSalarie);
  //   PT4 26/11/01 PH Incrementation du numéro d'ordre des mvts de type CP
  if MotifAbs = 'PRI' then Ordre := RendNumOrdre('CPA', CodeSalarie)
  else Ordre := RendNumOrdre('ABS', CodeSalarie);
  ImportInitialiseTobAbsenceSalarie(T);
  T.PutValue('PCN_SALARIE', Codesalarie);
  if MotifAbs = 'PRI' then
    T.PutValue('PCN_TYPEMVT', 'CPA')
  else T.PutValue('PCN_TYPEMVT', 'ABS');
  T.PutValue('PCN_ORDRE', Ordre);
  T.PutValue('PCN_CODERGRPT', Ordre);
  T.PutValue('PCN_DATEDEBUT', TDateDeb);
  if MotifAbs <> 'PRI' then T.PutValue('PCN_DATEFIN', TDateFin)
  else T.PutValue('PCN_DATEFIN', 2); // On force la date de fin pourf les CP à idate1900
  T.PutValue('PCN_DATEDEBUTABS', TDateDeb);
  T.PutValue('PCN_DATEFINABS', TDateFin);
  T.PutValue('PCN_TYPECONGE', MotifAbs);
  if MotifAbs = 'PRI' then
    T.PutValue('PCN_MVTPRIS', MotifAbs);
  T.PutValue('PCN_SENSABS', '-');
  // PT3 26/11/01 PH Récupération du libellé saisi de l'absence
  //  T.PutValue('PCN_LIBELLE'      ,'Mvt d''import');
  T.PutValue('PCN_LIBELLE', Copy(Libelle, 1, 35)); //PT10
  // PT1 14/06/01 PH Rajout récup libellé complémentaires et matin/AP
  T.PutValue('PCN_LIBCOMPL1', LibCompl1);
  T.PutValue('PCN_LIBCOMPL2', LibCompl2);
  T.PutValue('PCN_DEBUTDJ', Dj);
  T.PutValue('PCN_FINDJ', Fj);

  T.PutValue('PCN_DATEMODIF', Date);
  T.PutValue('PCN_DATEVALIDITE', TDateFin);
  Dateclot := Sal.getvalue('ETB_DATECLOTURECPN');
  T.PutValue('PCN_PERIODECP', ImportCalculPeriode(Dateclot, TDatefin));
  if Nbj = '' then NBj := '0';
  //   PT7 18/12/01 PH remplacement strtofloaat par valeur sinon erreur de conversion
  T.PutValue('PCN_JOURS', VALEUR(Nbj));
  if NbH = '' then NBH := '0';
  T.PutValue('PCN_HEURES', VALEUR(NbH));
  //   PT5 27/11/01 PH Initialisation de tous les champs de la table absencesalarie
  T.PutValue('PCN_SAISIEDEPORTEE', 'X');
  T.PutValue('PCN_VALIDRESP', 'VAL');
  T.PutValue('PCN_VALIDSALARIE', 'SAL'); //PT9
  T.PutValue('PCN_EXPORTOK', 'X');

  T.PutValue('PCN_ETABLISSEMENT', Sal.getvalue('PSA_ETABLISSEMENT'));
  T.Putvalue('PCN_TRAVAILN1', Sal.getvalue('PSA_TRAVAILN1'));
  T.Putvalue('PCN_TRAVAILN2', Sal.getvalue('PSA_TRAVAILN2'));
  T.Putvalue('PCN_TRAVAILN3', Sal.getvalue('PSA_TRAVAILN3'));
  T.Putvalue('PCN_TRAVAILN4', Sal.getvalue('PSA_TRAVAILN4'));
  T.Putvalue('PCN_CODESTAT', Sal.getvalue('PSA_CODESTAT'));
  T.Putvalue('PCN_CONFIDENTIEL', Sal.getvalue('PSA_CONFIDENTIEL'));

end;
// PT22   26/05/2004 PH V_50 Non traitement des absences portant sur des périodes closes mias ne provoquant pas de rejet du fichier

procedure TOF_PGIMPORTFIC.Erreur(Chaine, S: string; NoErreur: integer; WW: string = '');
var
  st: string;
begin
  if ((NoErreur = 48) or (NoErreur = 47)) and GetParamSoc('SO_IFDEFCEGID') then
  begin
    if WW = 'NF' then
    begin
      Writeln(FR, '');
      Writeln(FR, 'l''enregistrement absence suivant sera exclus car ces dates sont hors période');
      Writeln(FR, Chaine);
      taille := taille + 3;
      AbsExclue := TRUE;
    end;
    exit;
  end;
  // FIN PT22
  ExisteErreur := true;
  if not liberreur then
    {       RapportErreur.lines.Add('Erreur sur l''enregistrement ||   '+ chaine);}
  begin
    Writeln(FR, '');
    Writeln(FR, 'Erreur sur l''enregistrement ||   ' + chaine);
    taille := taille + 2;
  end;
  Liberreur := true;
  case NoErreur of
    5: st := 'Marque de début de fichier non trouvée';
    7: st := 'Impossible de récupérer l''exercice social en cours - rejet du fichier';
    10: st := 'Des enregistrements ont été détectés après la marque de fin de fichier';
    15: st := 'Marque de fin de fichier non trouvée';
    20: st := 'Code Enregistrement inexistant';
    25: st := 'Code salarié inexistant dans ce dossier';
    26: st := 'Salarie en suspension de paye';
    27: st := 'L''établissement auquel appartient le salarié ne gère pas les CP alors que des CP pris sont trouvés';
    30: st := 'Date début de période non valide';
    35: st := 'Date fin de période non valide';
    40: st := 'la date fin de période est antérieure à la date début de période';
    44: st := 'Il n''y a pas d''exercice social actif correspondant à ces dates';
    45: st := 'Numéro d''ordre obligatoire';
    46: st := 'Attention, salarié sorti';
    47: st := 'Le mois de la date début est une période cloturée';
    48: st := 'Le mois de la date fin est une période cloturée';
    50: st := 'Rubrique inexistante';
    52: st := 'La rubrique d''une ligne de commentaire doit être précédée d''un enregistrement faisant référence à cette même rubrique';
    53: st := 'La rubrique d''une ligne analytique doit être précédée d''un enregistrement faisant référence à cette même rubrique';
    55: st := 'Type alimentation obligatoire';
    57: st := 'L''alimentation de la rubrique est différente de celui du plan de paie';
    60: st := 'La base doit être renseignée et être numérique';
    65: st := 'La base ne doit pas être renseignée';
    70: st := 'Le taux doit être renseigné et être numérique';
    75: st := 'Le taux ne doit pas être renseigné';
    80: st := 'Le coefficient doit être renseigné et être numérique';
    85: st := 'Le coefficient ne doit pas être renseigné';
    90: st := 'Le montant doit être renseigné et être numérique';
    95: st := 'Le montant ne doit pas être renseigné';
    100: st := 'Section analytique inconnue';
    101: st := 'Il y a des lignes analytiques alors que vous ne gerez pas l''analytique';
    102: st := 'Le numéro de ventilation analytique doit être renseigné et être numérique';
    110: st := 'Date début absence non valide';
    120: st := 'Date fin absence non valide';
    130: st := 'La date début doit être antérieure à la date de fin';
    140: st := 'il existe déjà une absence pour ce salarié à cette période';
    150: st := 'Nombre de jours non numérique';
    160: st := 'Nombre d''heures non numérique';
    170: st := 'Code motif inexistant dans la table';
    220: st := 'Code enregistrement de la ligne Dossier-société erroné';
    230: st := 'Le numéro de dossier est erroné';
    240: st := 'Erreur de structure de l''enregistrement';
    250: st := 'Impossible de créer le salarié car le matricule existe déjà';
    251: st := 'Impossible de créer le salarié car il manque la date d''entrée';
    252: st := 'Impossible de créer le salarié car il manque son nom';
    253: st := 'Impossible de créer le salarié car il manque son prénom';
    255: st := 'Impossible de créer le salarié car plusieurs salariés en attente de création ont le même matricule';
    256: st := 'Création des salariés interrompue car il n''existe aucun salarié';
    257: st := 'Création des salariés interrompue car l''établissement indiqué n''existe pas';
    258: st := 'il est impossible de créer des salariés si la condification n''est pas numérique';
  else st := 'une erreur a été détectée dans votre fichier,veuillez le vérifier'
  end;
  Writeln(FR, st);
  taille := taille + 1;
  {   if RapportErreur <> nil then
        RapportErreur.lines.Add(St+' ||  '+S);}
end;

procedure TOF_PGIMPORTFIC.Anomalie(Chaine, S: string; NoErreur: integer);
var
  st: string;
begin
  if not libanomalie then
  begin
    if Noerreur <> 98 then
      Writeln(FR, 'Erreur non bloquante sur l''enregistrement ||   ' + chaine);
  end;
  Libanomalie := true;
  case NoErreur of
    26: st := 'Salarie en suspension de paye';
    98: st := S;
  else st := 'une erreur a été détectée dans votre fichier,veuillez le vérifier'
  end;
  Writeln(FR, st);
  taille := taille + 1;
end;

procedure TOF_PGIMPORTFIC.SepChange(Sender: TObject);
begin
  // PT18   10/09/2003 PH Controle du séparateur du fichier import et creation automatique fichier rapport
  if Sep <> nil then
    if (Sep.value = 'TAB') or (Copy(Sep.Value, 1, 3) = 'Tab') then Separateur := #9
    else if (Sep.value = 'PIP') or (Copy(Sep.Value, 1, 3) = 'Pip') then Separateur := '|'
    else if (Sep.value = 'PVI') or (Copy(Sep.Value, 1, 3) = 'Poi') then Separateur := ';';
end;

function TOF_PGIMPORTFIC.ExisteRubrique(Rubrique: string): boolean;
var
  T: tob;
begin
  result := false;
  if not PgImpnotVide(tob_rubrique, true) then exit;

  T := Tob_rubrique.findfirst(['PRM_RUBRIQUE', 'PRM_PREDEFINI'], [rubrique, 'DOS'], true);
  if T = nil then
    T := Tob_rubrique.findfirst(['PRM_RUBRIQUE', 'PRM_PREDEFINI'], [rubrique, 'STD'], true);
  if T = nil then
    T := Tob_rubrique.findfirst(['PRM_RUBRIQUE', 'PRM_PREDEFINI'], [rubrique, 'CEG'], true);
  if T = nil then exit;
  TypeBase := T.getvalue('PRM_TYPEBASE');
  TypeTaux := T.getvalue('PRM_TYPETAUX');
  TypeCoeff := T.getvalue('PRM_TYPECOEFF');
  TypeMontant := T.getvalue('PRM_TYPEMONTANT');
  result := true;
end;

procedure TOF_PGIMPORTFIC.FicChange(Sender: TObject);
var
  LeFic, St: string;
  F: TextFile;
begin
  // PT18   10/09/2003 PH Controle du séparateur du fichier import et creation automatique fichier rapport
  if not Assigned(Sender) then exit;
  LeFic := THEdit(Sender).Text;
  if LeFic = '' then exit;
  if THEdit(Sender).Name = 'FICRAPPORT' then
  begin
    if not FileExists(LeFic) then
    begin
      AssignFile(F, LeFic);
{$I-}ReWrite(F);
{$I+}if IoResult <> 0 then
      begin
        PGIBox('Fichier inaccessible : ' + LeFic, 'Abandon du traitement');
        Exit;
      end;
      closeFile(F);
    end;
  end
  else
  begin
    AssignFile(F, LeFic);
    reset(F);
    ReadLn(F, St);
    ReadLn(F, St);
    if Sep <> nil then
    begin
      if Pos(';', St) > 0 then
      begin
        Sep.Value := 'PVI';
        Separateur := ';';
      end
      else if Pos('|', St) > 0 then
      begin
        Sep.Value := 'PIP';
        Separateur := '|';
      end;
    end;
    closeFile(F);
  end;
  // FIN PT18
end;

{$IFNDEF EAGLCLIENT}
procedure TOF_PGIMPORTFIC.ImprimeClick(Sender: TObject);
var
  MPages: tpagecontrol;
begin
  MPages := TPageControl(getcontrol('PAGECONTROL1'));
  if MPages <> nil then
    PrintPageDeGarde(MPages, TRUE, TRUE, FALSE, Ecran.Caption, 0);
end;
{$ENDIF !EAGLCLIENT} 

procedure TOF_PGIMPORTFIC.ControleSalarie(TypeEnr, chaine, S: string; var CodeSalarie: string; Ecrit: Boolean);
var
  T, TAtt: tob;
begin
  TAtt := nil;
  DateEntree := 0;
  DateSortie := 0;
  if tob_sal = nil then
    Erreur(chaine, Codesalarie, 25) else
  begin
    T := Tob_Sal.findfirst(['PSA_SALARIE'], [Codesalarie], true);
    if T = nil then TAtt := Tob_SalATT.findfirst(['PIM_SALARIE'], [Codesalarie], true); // PT27
    if (T = nil) and (TAtt = nil) then Erreur(chaine, Codesalarie, 25) // PT27
    else
      if (T <> nil) and (T.GetValue('PSA_SUSPENSIONPAIE') = 'X') then Anomalie(chaine, Codesalarie, 26);
      // DEB PT27
    if Ecrit and (TAtt <> nil) then
    begin
      if (TAtt.GetValue('TRAITSAL') = 'C') then CodeSalarie := TAtt.GetValue('CodeS5');
    end;
    // FIN PT27
  end;
end;
// La periode contenent ces dates est elle cloturée ?
// result = 0 : non
//          1 : pas d'exercice social trouvé
//          2 : Mois date début cloturé
//          3 : Mois date de fin cloturé
//
//  PT8 15/02/02 PH test pour les mvts de cp sur mois clos N et N+1

function TOF_PGIMPORTFIC.IsPeriodecloture(Dated, DateF: TDateTime; CasCp: string = '-'): integer;
var
  Q: TQuery;
  aa, mmd, mmf, jj: WORD;
  trouve: boolean;
  d, f: integer;
begin
  // PT19
  if (Cloture = '') or (Dated < DebExer) or (DateF > FinExer) then
  begin
    trouve := false;
    Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER', TRUE);
    while not Q.EOF do
    begin
      DebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
      FinExer := Q.FindField('PEX_DATEFIN').AsDateTime;
      if (debExer <= DateD) and (FinExer >= Datef) then
      begin
        trouve := true;
        break;
      end;
      Q.Next;
    end;
    Cloture := Q.FindField('PEX_CLOTURE').asstring;
    Ferme(Q);
    if not trouve then
    begin
      result := 1; // pas d'exercice social actif correspondant à ces dates
      exit;
    end;
  end;
  // FIN PT19
  decodedate(Dated, aa, mmd, jj);
  decodedate(dateF, aa, mmf, jj);
  if GetParamSoc('SO_PGDECALAGE') = TRUE then // PT24
  begin
    if mmd = 12 then d := 1 else d := mmd + 1;
    if mmf = 12 then f := 1 else f := mmf + 1;
  end
  else
  begin
    d := mmd;
    f := mmf;
  end;
  //  PT8 15/02/02 PH test pour les mvts de cp sur mois clos N et N+1
  if CasCp = 'X' then
  begin
    result := 0;
    if (Cloture[d] = 'X') then
    begin
      if d < 12 then
      begin
        if Cloture[d + 1] = 'X' then result := 2;
      end else result := 2;
    end;
    if (result = 0) and (Cloture[f] = 'X') then
    begin
      if f < 12 then
      begin
        if Cloture[f + 1] = 'X' then result := 3;
      end else result := 3;
    end;
  end
  else
  begin
    if Cloture[d] = 'X' then result := 2
    else
      if Cloture[f] = 'X' then result := 3
    else
      result := 0;
  end;
  //  FIN PT8
// PT19
  if (result = 2) or (result = 3) then
  begin
    // Rajouter le test d'un paramsoc pour outrepasser les controles SPECIFCEGID
  end;
  // FIN PT19
end;

function TOF_PGIMPORTFIC.RendNumOrdre(TypeMvt, CodeSalarie: string): integer;
var
  Ordre: Integer;
  T1: TOB;
begin
  T1 := Tob_OrdreAbs.findFirst(['PCN_SALARIE', 'PCN_TYPEMVT'], [CodeSalarie, TypeMvt], FALSE);
  if T1 <> nil then
  begin
    Ordre := T1.GetValue('NUMORDRE') + 1;
    T1.PutValue('NUMORDRE', Ordre);
  end
  else
  begin
    Ordre := 1;
    T1 := TOB.Create('LES NUMORDRES ABSENCESALARIE', Tob_OrdreAbs, -1);
    T1.AddChampSup('PCN_SALARIE', FALSE);
    T1.AddChampSup('PCN_TYPEMVT', FALSE);
    T1.AddChampSup('NUMORDRE', FALSE);
    T1.PutValue('PCN_SALARIE', CodeSalarie);
    T1.PutValue('PCN_TYPEMVT', TypeMvt);
    T1.PutValue('NUMORDRE', Ordre);
  end;
  result := Ordre;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dev Paie
Créé le ...... : 11/09/2003
Modifié le ... :   /  /
Description .. : Le nom du fichier obligatoirement renseigné avec une
Suite ........ : extension tq .t (.txt,.trt,.ttt ....)
Suite ........ : Fichier existant obligatoirement
Suite ........ : Le Type d'import = 'AUTO' sinon considéré comme appli paie
Suite ........ : Le séparateur est un Char := |
Suite ........ :
Suite ........ : Test si creation du fichier de rapport peut être créer mais
Suite ........ : il aura une extension en .log au lieu de .trt ou .txt ou .nsv (NetExpert PCL)
Mots clefs ... : PAIE;GIGA
*****************************************************************}

function TOF_PGIMPORTFIC.PgImportDef(LeNomduFichier, LeTypImp: string; LeSeparateur: Char = '|'): Boolean;
var
  LeRapport: string;
  ll: Integer;
  Sep: THValComboBox;
begin
  result := FALSE;
  if LeNomDuFichier = '' then exit;
  Separateur := LeSeparateur; //AB-20050104-Import Gestion Affaire
  NomDuFichier := LeNomDuFichier;
  TypImport := LeTypImp;
  //  if not FileExists(LeNomduFichier) then exit;
  ll := Pos('.NSV', LeNomduFichier);
  if ll <= 0 then exit; // PT24
  if LeTypImp <> 'AUTO' then exit;
  if (LeSeparateur <> '|') then exit;
  SetControlText('EDTFICIMPORT', LeNomDuFichier);
  SetControlEnabled('EDTFICIMPORT', FALSE);
  LeRapport := Copy(LeNomDuFichier, 1, ll) + '.log';
  SetControlText('FICRAPPORT', LeRapport);
  SetControlEnabled('FICRAPPORT', FALSE);
  SetControlEnabled('SEPARATEUR', FALSE);
  Sep := THValComboBox(getcontrol('SEPARATEUR'));
  if (Sep <> nil) then Sep.value := 'PIP';
  result := TRUE;
end;

function TOF_PGIMPORTFIC.ImportRendPeriodeEnCours(var ExerPerEncours, DebPer, FinPer: string): Boolean;
var
  Q: TQuery;
begin
  result := FALSE;
  DebPer := DateToStr(idate1900);
  FinPer := DateToStr(idate1900);
  Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE);
  if not Q.EOF then
  begin
    DebPer := Q.FindField('PEX_DEBUTPERIODE').AsString;
    FinPer := Q.FindField('PEX_FINPERIODE').AsString;
    ExerPerEncours := Q.FindField('PEX_EXERCICE').AsString;
    result := TRUE;
    //  PT1
  end;
  Ferme(Q);
end;

function TOF_PGIMPORTFIC.ImportRendExerSocialEnCours(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime): Boolean;
var
  Q: TQuery;
  DatF: TDateTime;
  Jour, Mois, Annee: WORD;
begin
  result := FALSE;
  DebExer := idate1900;
  FinExer := idate1900;
  Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE);
  if not Q.EOF then
  begin
    DatF := Q.FindField('PEX_FINPERIODE').AsFloat; //Q.Fields[7].AsFloat; // Recup date de fin periode en cours
    DecodeDate(DatF, Annee, Mois, Jour);
    MoisE := ImportColleZeroDevant(Mois, 2);
    ComboExer := Q.FindField('PEX_EXERCICE').AsString; //Q.Fields[0].AsString; // recup Combo identifiant exercice
    AnneeE := Q.FindField('PEX_ANNEEREFER').AsString; //Q.Fields[8].AsString; // recup Annee de exercice
    DebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
    FinExer := Q.FindField('PEX_DATEFIN').AsDateTime;
    result := TRUE;
    // PT1
  end;
  Ferme(Q);
end;

function TOF_PGIMPORTFIC.ImportColleZeroDevant(Nombre, LongChaine: integer): string;
var
  tabResult: string;
  TabInt: string;
  i, j: integer;
begin
  tabResult := '';
  for i := 1 to LongChaine do
  begin
    if Nombre < power(10, i) then
    begin
      TabInt := inttostr(Nombre);
      // colle (LongChaine-i zéro devant]
      for j := 0 to (LongChaine - i - 1)
        do insert('0', TabResult, j);
      result := concat(TabResult, Tabint);
      exit;
    end;
    if i > LongChaine then result := inttostr(Nombre);
  end;
end;

// DEB PT27
procedure TOF_PGIMPORTFIC.ControleEnregVSA(TypeEnr, chaine, S: string; Ecrit: boolean);
var
  Codesalarie, DateDebutperiode, DateFinperiode, SNoOrdre: string;
  Axe, Section, CodeAxe, sMontant, LibSect: string;
  TDateDeb, TDateFin: tdatetime;
  T: tob;
  TopCloture: integer;
begin
  libErreur := false;
  libanomalie := false;
  if not GetParamSoc('SO_PGANALYTIQUE') then Erreur(chaine, '', 101); // PT24
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if CodeSalarie <> Salarieprec then
  begin
    ControleSalarie(TypeEnr, chaine, S, CodeSalarie, Ecrit);
  end;
  Salarieprec := Codesalarie;
  DateDebutperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DatedebutPeriode) then
    Erreur(chaine, DatedebutPeriode, 30);
  DateFinperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateFinPeriode) then
    Erreur(chaine, DateFinPeriode, 35);
  TDateDeb := 0;
  TDateFin := 0;
  if (Isvaliddate(DatedebutPeriode)) and (Isvaliddate(DatefinPeriode)) then
  begin
    TDateDeb := strtodate(DatedebutPeriode);
    TDateFin := strtodate(DatefinPeriode);
  end;
  if TDateFin < TDateDeb then
    Erreur(chaine, DateDebutPeriode + '//' + DateFinPeriode, 40);
  Topcloture := IsPeriodecloture(TDatedeb, TDateFin);
  if TopCloture = 1 then
    Erreur(Chaine, '', 44);
  if TopCloture = 2 then
    Erreur(Chaine, '', 47);
  if TopCloture = 3 then
    Erreur(Chaine, '', 48);
  if (DateSortie > 10) and (DateSortie < TDateDeb) then
    Erreur(Chaine, '', 46);
  Axe := readtokenPipe(S, Separateur);
  if (Axe < 'A1') or (Axe > 'A5') then Erreur(chaine, Axe, 103);
  Section := readtokenPipe(S, Separateur);
  Smontant := readtokenPipe(S, Separateur);
  if sMontant = '' then
    Erreur(chaine, sMontant, 90)
  else
    if not Isnumeric(sMontant) then Erreur(chaine, sMontant, 90);
  SNoOrdre := readtokenPipe(S, Separateur);
  if SNoOrdre = '' then
    Erreur(chaine, SNoOrdre, 102)
  else
    if not Isnumeric(SNoOrdre) then Erreur(chaine, SNoOrdre, 102);

  // PT16   28/08/2003 PH Création des sections en récupérant le libellé de la section
  if S <> '' then LibSect := readtokenPipe(S, Separateur)
  else LibSect := '';
  if not ExisteSection(Axe, Section, LibSect) then Erreur(Chaine, Section, 100);
  // Fin PT16

  if not Ecrit then exit;
  T := Tob.create('PAIEVENTIL', TOB_Paieventil, -1);
  if T = nil then exit;
  CodeAxe := Copy(Axe, 2, 1);
  T.PutValue('PAV_NATURE', 'VS' + CodeAxe);
  T.PutValue('PAV_COMPTE', CodeSalarie);
  T.PutValue('PAV_DATEDEBUT', strtodate(DateDebutperiode));
  T.PutValue('PAV_DATEFIN', strtodate(DateFinperiode));
  T.PutValue('PAV_SECTION', Section);
  T.PutValue('PAV_TAUXMONTANT', Valeur(sMontant));
  T.PutValue('PAV_TAUXQTE1', 0);
  T.PutValue('PAV_TAUXQTE2', 0);
  T.PutValue('PAV_NUMEROVENTIL', StrToInt(SNoOrdre));
  T.PutValue('PAV_SOCIETE', GetParamSoc('SO_SOCIETE'));
  T.PutValue('PAV_MONTANT', 0);
  T.PutValue('PAV_SOUSPLAN1', '');
  T.PutValue('PAV_SOUSPLAN2', '');
  T.PutValue('PAV_SOUSPLAN3', '');
  T.PutValue('PAV_SOUSPLAN4', '');
  T.PutValue('PAV_SOUSPLAN5', '');
  T.PutValue('PAV_SOUSPLAN6', '');
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 10/12/2004
Modifié le ... :   /  /
Description .. : Fonction de création automatique de salarié
Suite ........ : 10/12/2004 Ajout de fonctionnalité : prise en compte de
Suite ........ : la modification des infos salariés
Mots clefs ... : PAIE;
*****************************************************************}
// DEB PT27
procedure TOF_PGIMPORTFIC.ControleEnregSAL(TypeEnr, chaine, S: string; Ecrit: boolean);
var
  T, T1, T2, TSalACreer, TEtab : TOB;
  CodeSalarie, DatEnt, SNom, SPrenom, AD1, AD2, AD3, CP, Ville, Tel, DNai, Nat, Sexe, NOSS: string;
  Etab, DSortie, Civilite, Portable, st, Etablis: string;
  Anomalie: Boolean;
  i, LongLib, LongPre: integer;
  Q: TQuery;
  InfTiers: Info_Tiers;
  CodeAuxi, Libell, Prenom, LeRapport, ConvColl, TraitSal, Trait: string;
  QQuery: TQuery;
  IMax, MaxInterim: integer;

  function TraiteErreur(Chaine, CodeSal: string; Noerreur: Integer): Boolean;
  begin
    Erreur(Chaine, CodeSal, NoErreur);
    Result := TRUE;
  end;

begin
  if not Ecrit then
  begin
    libErreur := false;
    libanomalie := false;
    if Tob_SalATT = nil then Tob_SalATT := Tob.Create('Mes_Salaries_A_CREER', nil, -1);
    Codesalarie := readtokenPipe(S, Separateur);
    if (Copy(Codesalarie, 1, 3) = 'S1_') then TraitSal := 'C' else
      traitSal := 'M';
    if (TraitSal = 'C') then
    begin
      if (not GetParamSoc('SO_PGINCSALARIE')) or (GetParamSoc('SO_PGTYPENUMSAL') <> 'NUM') then
      begin
        Erreur(Chaine, Codesalarie, 258);
        exit;
      end;
    end;
    // ATTENTION, si modification alors reporter les modifs dans UTOMSalaries
    if MatSouche = -999 then
    begin
      MaxInterim := 0;
      if (GetParamSoc('SO_PGCODEINTERIM') = True) and (GetParamSoc('SO_PGINTERIMAIRES') = True) then
      begin
        QQuery := OpenSQL('SELECT MAX(PSI_INTERIMAIRE) FROM INTERIMAIRES', True);
        if (not QQuery.EOF) and (QQuery.Fields[0].AsString <> '') then MaxInterim := StrToInt(QQuery.Fields[0].AsString)
        else MaxInterim := 0;
        Ferme(QQuery);
      end;

      QQuery := OpenSQL('SELECT MAX(PSA_SALARIE) FROM SALARIES', TRUE);
      if not QQuery.Eof then
      begin // Attention, respecter le type du numéro salarié soit alpha soit numérique avec increment auto
        if QQuery.Fields[0].AsString <> '' then Imax := StrToInt(QQuery.Fields[0].AsString) // On récupère le max du numéro de salarié
        else Imax := 0;
        if MaxInterim > IMax then IMax := MaxInterim;
        if IMax < 2147483647 then
          MatSouche := Imax;
      end
      else MatSouche := 0;
      Ferme(QQuery);
    end;
    // contrôle existance du salarié
    T := Tob_sal.FindFirst(['PSA_SALARIE'], [CodeSalarie], true);
    if (T <> nil) and (TraitSal = 'C') then // Mode creation salarie et déjà existant
    begin
      Erreur(Chaine, Codesalarie, 250);
      exit;
    end;
    T := Tob_SalATT.FindFirst(['PMI_SALARIE'], [CodeSalarie], true);
    if (T <> nil) and (TraitSal = 'C') then // Salarié déjà présent ==> doublon de code salarié à créer
    begin
      Erreur(Chaine, Codesalarie, 255);
      exit;
    end;
    Anomalie := FALSE;
    DatEnt := readtokenPipe(S, Separateur);
    if (DatEnt = '') or (not IsValidDate(DatEnt)) then Anomalie := TraiteErreur(Chaine, Codesalarie, 251);
    SNom := readtokenPipe(S, Separateur);
    if Snom = '' then Anomalie := TraiteErreur(Chaine, Codesalarie, 252);
    SPrenom := readtokenPipe(S, Separateur);
    if SPrenom = '' then Anomalie := TraiteErreur(Chaine, Codesalarie, 253);
    if Anomalie then exit;
    AD1 := readtokenPipe(S, Separateur);
    AD2 := readtokenPipe(S, Separateur);
    AD3 := readtokenPipe(S, Separateur);
    CP := readtokenPipe(S, Separateur);
    Ville := readtokenPipe(S, Separateur);
    Tel := readtokenPipe(S, Separateur);
    DNai := readtokenPipe(S, Separateur);
    Nat := readtokenPipe(S, Separateur);
    Sexe := readtokenPipe(S, Separateur);
    NOSS := readtokenPipe(S, Separateur);
    Civilite := readtokenPipe(S, Separateur);
    Portable := readtokenPipe(S, Separateur);
    DSortie := readtokenPipe(S, Separateur);
    Etablis := readtokenPipe(S, Separateur);
    if Etablis <> '' then
    begin
      TEtab := Tob_Etab.FindFirst(['ETB_ETABLISSEMENT'], [Etablis], TRUE);
      if TEtab = nil then
      begin
        Erreur(Chaine, Etablis, 257);
        exit;
      end;
    end;
    T1 := TOB.Create('SALATTENTE', TOB_SalATT, -1);
    T1.PutValue('PIM_SALARIE', CodeSalarie);
    T1.PutValue('PIM_DATEENTREE', StrToDate(DatEnt));
    T1.PutValue('PIM_LIBELLE', Snom);
    T1.PutValue('PIM_PRENOM', SPrenom);
    T1.PutValue('PIM_ADRESSE1', AD1);
    T1.PutValue('PIM_ADRESSE2', AD2);
    T1.PutValue('PIM_ADRESSE3', AD3);
    T1.PutValue('PIM_CODEPOSTAL', CP);
    T1.PutValue('PIM_VILLE', Ville);
    T1.PutValue('PIM_TELEPHONE', Tel);
    if IsValidDate(DNai) then T1.PutValue('PIM_DATENAISSANCE', StrToDate(DNai))
    else T1.PutValue('PIM_DATENAISSANCE', IDate1900);
    T1.PutValue('PIM_NATIONALITE', Nat);
    T1.PutValue('PIM_SEXE', Sexe);
    T1.PutValue('PIM_NUMEROSS', NOSS);
    T1.AddChampSupValeur('PORTABLE', Portable, False);
    T1.AddChampSupValeur('CIVILITE', Civilite, False);
    if IsValidDate(DSortie) then T1.AddChampSupValeur('DATESORTIE', StrToDate(DSortie), FALSE)
    else T1.AddChampSupValeur('DATESORTIE', IDate1900, FALSE);
    if Etablis <> '' then T1.AddChampSupValeur('ETABLISSEMENT', Etablis, False)
    else T1.AddChampSupValeur('ETABLISSEMENT', '', False);
    T1.AddChampSupValeur('TRAITSAL', TraitSal, False); // Stockage du type de traitement à faire pour le salarié
    {$IFDEF AFFAIRE}
      {$ifndef OGC}
        if TraitSal = 'C' then // Phase creation alors ajout 1 ligne dans la table de correspondance
        begin
          MatSouche := MatSouche + 1;
          T1.AddChampSupValeur('CodeS5', ColleZeroDevant(MatSouche, 10), False);
        end
        else
          T1.AddChampSupValeur('CodeS5', CodeSalarie, False);
      {$endif OGC}
    {$ENDIF AFFFAIRE}
  end // Fin du cas controle des enregistrements salaries
  else
  begin
    if (Tob_SalATT <> nil) AND (NOT TopCreat) then
    begin
      TopCreat := TRUE;
      TSalACreer := TOB.Create('les_salaries_acreer', nil, -1);
      Etab := GetParamSoc('SO_ETABLISDEFAUT');
      if Etab <> '' then TEtab := Tob_Etab.FindFirst(['ETB_ETABLISSEMENT'], [Etab], TRUE)
      else TEtab := Tob_Etab.FindFirst([''], [''], TRUE);
      if TEtab = nil then
      begin
        Erreur(Chaine, Codesalarie, 256);
        exit;
      end;
      Etab := TEtab.GetValue('ETB_ETABLISSEMENT');
      ConvColl := TEtab.GetValue('ETB_CONVENTION');

      for i := 0 to Tob_SalATT.detail.count - 1 do
      begin
        T1 := Tob_SalATT.detail[i];
        Trait := T1.GetValue('TRAITSAL');
        if Trait = 'C' then
        begin
          T2 := TOB.Create('SALARIES', TSalACreer, -1);
          T2.PutValue('PSA_SALARIE', T1.GetValue('CODES5'));
          Etablis := T1.GetValue('ETABLISSEMENT');
          T2.PutValue('PSA_ETABLISSEMENT', Etab);
          T2.PutValue('PSA_CONVENTION', Convcoll);
          if Etablis <> '' then
          begin
            if Etablis <> Etab then
            begin
              TEtab := Tob_Etab.FindFirst(['ETB_ETABLISSEMENT'], [Etablis], TRUE);
              if TEtab <> nil then
              begin
                T2.PutValue('PSA_ETABLISSEMENT', Etablis);
                T2.PutValue('PSA_CONVENTION', TEtab.GetValue('ETB_CONVENTION'));
              end;
            end;
          end;
          T2.PutValue('PSA_DATEENTREE', T1.GetValue('PIM_DATEENTREE'));
          T2.PutValue('PSA_DATEANCIENNETE', Getfield('PSA_DATEENTREE'));
          T2.PutValue('PSA_ANCIENPOSTE', T2.GetValue('PSA_DATEENTREE'));
          T2.PutValue('PSA_DATENAISSANCE', Idate1900);
          T2.PutValue('PSA_DATESORTIE', Idate1900);
          T2.PutValue('PSA_DATEENTREEPREC', Idate1900);
          T2.PutValue('PSA_DATESORTIEPREC', Idate1900);
          T2.PutValue('PSA_DATELIBRE1', Idate1900);
          T2.PutValue('PSA_DATELIBRE2', Idate1900);
          T2.PutValue('PSA_DATELIBRE3', Idate1900);
          T2.PutValue('PSA_DATELIBRE4', Idate1900);
          T2.PutValue('PSA_ORDREAT', '1');
          T2.PutValue('PSA_DADSFRACTION', '1');
          T2.PutValue('PSA_REGIMESS', '200');
          T2.PutValue('PSA_TAUXPARTSS', 0);
          T2.PutValue('PSA_TAUXPARTIEL', 0);
          T2.PutValue('PSA_DADSDATE', IDate1900);
          T2.PutValue('PSA_NATIONALITE', 'FRA');
          T2.PutValue('PSA_PAYSNAISSANCE', 'FRA');
          T2.PutValue('PSA_PRISEFFECTIF', 'X');
          T2.PutValue('PSA_UNITEPRISEFF', 1);
          T2.PutValue('PSA_LIBELLE', T1.GetValue('PIM_LIBELLE'));
          T2.PutValue('PSA_PRENOM', T1.GetValue('PIM_PRENOM'));
          T2.PutValue('PSA_ADRESSE1', T1.GetValue('PIM_ADRESSE1'));
          T2.PutValue('PSA_ADRESSE2', T1.GetValue('PIM_ADRESSE2'));
          T2.PutValue('PSA_ADRESSE2', T1.GetValue('PIM_ADRESSE3'));
          T2.PutValue('PSA_CODEPOSTAL', T1.GetValue('PIM_CODEPOSTAL'));
          T2.PutValue('PSA_VILLE', T1.GetValue('PIM_VILLE'));
          T2.PutValue('PSA_TELEPHONE', T1.GetValue('PIM_TELEPHONE'));
          T2.PutValue('PSA_DATENAISSANCE', T1.GetValue('PIM_DATENAISSANCE'));
          T2.PutValue('PSA_NATIONALITE', T1.GetValue('PIM_NATIONALITE'));
          T2.PutValue('PSA_SEXE', T1.GetValue('PIM_SEXE'));
          T2.PutValue('PSA_NUMEROSS', T1.GetValue('PIM_NUMEROSS'));
          T2.PutValue('PSA_CIVILITE', T1.GetValue('CIVILITE'));
          T2.PutValue('PSA_PORTABLE', T1.GetValue('PORTABLE'));
          T2.PutValue('PSA_DATESORTIE', T1.GetValue('DATESORTIE'));
          T2.PutValue('PSA_TYPPROFIL', 'ETB');
          T2.PutValue('PSA_TYPPROFILREM', 'ETB');
          T2.PutValue('PSA_TYPPERIODEBUL', 'ETB');
          T2.PutValue('PSA_TYPPROFILRBS', 'ETB');
          T2.PutValue('PSA_TYPPROFILAFP', 'ETB');
          T2.PutValue('PSA_TYPPROFILAPP', 'ETB');
          T2.PutValue('PSA_TYPPROFILRET', 'ETB');
          T2.PutValue('PSA_TYPPROFILMUT', 'ETB');
          T2.PutValue('PSA_TYPPROFILPRE', 'ETB');
          T2.PutValue('PSA_TYPPROFILTSS', 'ETB');
          T2.PutValue('PSA_TYPPROFILCGE', 'ETB');
          T2.PutValue('PSA_TYPPROFILANC', 'ETB');
          T2.PutValue('PSA_TYPEDITBULCP', 'ETB');
          T2.PutValue('PSA_CPACQUISMOIS', 'ETB');
          T2.PutValue('PSA_CPACQUISSUPP', 'ETB');
          T2.PutValue('PSA_TYPNBACQUISCP', 'ETB');
          T2.PutValue('PSA_CPTYPEMETHOD', 'ETB');
          T2.PutValue('PSA_CPTYPERELIQ', 'ETB');
          T2.PutValue('PSA_CPACQUISANC', 'ETB');
          T2.PutValue('PSA_DATANC', 'ETB');
          T2.PutValue('PSA_TYPPROFILTRANS', 'ETB');
          T2.PutValue('PSA_TYPPROFILFNAL', 'DOS');
          T2.PutValue('PSA_STANDCALEND', 'ETB');
          T2.PutValue('PSA_TYPREDREPAS', 'ETB');
          T2.PutValue('PSA_TYPREDRTT1', 'ETB');
          T2.PutValue('PSA_TYPREDRTT2', 'ETB');
          T2.PutValue('PSA_TYPDADSFRAC', 'ETB');
          T2.PutValue('PSA_TYPPRUDH', 'ETB');
          T2.PutValue('PSA_CPTYPEVALO', 'ETB');
          T2.PutValue('PSA_TYPREGLT', 'ETB');
          T2.PutValue('PSA_TYPVIRSOC', 'ETB');
          T2.PutValue('PSA_TYPDATPAIEMENT', 'ETB');
          T2.PutValue('PSA_TYPPAIACOMPT', 'ETB');
          T2.PutValue('PSA_TYPACPSOC', 'ETB');
          T2.PutValue('PSA_TYPPAIFRAIS', 'ETB');
          T2.PutValue('PSA_TYPFRAISSOC', 'ETB');
          T2.PutValue('PSA_TYPJOURHEURE', 'DOS');
          T2.PutValue('PSA_TYPEDITORG', 'ETB');
          T2.PutValue('PSA_TYPACTIVITE', 'ETB');
          T2.PutValue('PSA_ACTIVITE', TEtab.GetValue('ETB_ACTIVITE'));
          T2.PutValue('PSA_EDITORG', TEtab.GetValue('ETB_EDITORG'));
          T2.PutValue('PSA_PROFILANCIEN', TEtab.GetValue('ETB_PROFILANCIEN'));
          T2.PutValue('PSA_PROFIL', TEtab.GetValue('ETB_PROFIL'));
          T2.PutValue('PSA_PROFILREM', TEtab.GetValue('ETB_PROFILREM'));
          T2.PutValue('PSA_PERIODBUL', TEtab.GetValue('ETB_PERIODBUL'));
          T2.PutValue('PSA_PROFILRBS', TEtab.GetValue('ETB_PROFILRBS'));
          T2.PutValue('PSA_REDREPAS', TEtab.GetValue('ETB_REDREPAS'));
          T2.PutValue('PSA_REDRTT1', TEtab.GetValue('ETB_REDRTT1'));
          T2.PutValue('PSA_REDRTT2', TEtab.GetValue('ETB_REDRTT2'));
          T2.PutValue('PSA_PROFILAFP', TEtab.GetValue('ETB_PROFILAFP'));
          T2.PutValue('PSA_PCTFRAISPROF', TEtab.GetValue('ETB_PCTFRAISPROF'));
          T2.PutValue('PSA_PROFILAPP', TEtab.GetValue('ETB_PROFILAPP'));
          T2.PutValue('PSA_PROFILRET', TEtab.GetValue('ETB_PROFILRET'));
          T2.PutValue('PSA_PROFILMUT', TEtab.GetValue('ETB_PROFILMUT'));
          T2.PutValue('PSA_PROFILPRE', TEtab.GetValue('ETB_PROFILPRE'));
          T2.PutValue('PSA_PROFILTSS', TEtab.GetValue('ETB_PROFILTSS'));
          T2.PutValue('PSA_PROFILTRANS', TEtab.GetValue('ETB_PROFILTRANS'));
          T2.PutValue('PSA_CALENDRIER', TEtab.GetValue('ETB_STANDCALEND'));
          T2.PutValue('PSA_PROFILFNAL', GetParamSoc('SO_PGPROFILFNAL'));
          T2.PutValue('PSA_PROFILCGE', TEtab.GetValue('ETB_PROFILCGE'));
          T2.PutValue('PSA_NBREACQUISCP', TEtab.GetValue('ETB_NBREACQUISCP'));
          T2.PutValue('PSA_NBACQUISCP', TEtab.GetValue('ETB_NBACQUISCP'));
          T2.PutValue('PSA_NBRECPSUPP', TEtab.GetValue('ETB_NBRECPSUPP'));
          T2.PutValue('PSA_VALORINDEMCP', TEtab.GetValue('ETB_VALORINDEMCP'));
          T2.PutValue('PSA_RELIQUAT', TEtab.GetValue('ETB_RELIQUAT'));
          T2.PutValue('PSA_BASANCCP', TEtab.GetValue('ETB_BASANCCP'));
          T2.PutValue('PSA_TYPDATANC', TEtab.GetValue('ETB_TYPDATANC'));
          T2.PutValue('PSA_MVALOMS', TEtab.GetValue('ETB_MVALOMS'));
          T2.PutValue('PSA_VALODXMN', TEtab.GetValue('ETB_VALODXMN'));
          T2.PutValue('PSA_EDITBULCP', TEtab.GetValue('ETB_EDITBULCP'));
          T2.PutValue('PSA_JOURHEURE', TEtab.GetValue('ETB_JOURHEURE'));
          T2.PutValue('PSA_JOURHEURE', GetParamSoc('SO_PGMETODHEURES'));
          T2.PutValue('PSA_PRUDHCOLL', TEtab.GetValue('ETB_PRUDHCOLL'));
          T2.PutValue('PSA_PRUDHSECT', TEtab.GetValue('ETB_PRUDHSECT'));
          T2.PutValue('PSA_PRUDHVOTE', TEtab.GetValue('ETB_PRUDHVOTE'));
          T2.PutValue('PSA_PGMODEREGLE', TEtab.GetValue('ETB_PGMODEREGLE'));
          T2.PutValue('PSA_RIBVIRSOC', TEtab.GetValue('ETB_RIBSALAIRE'));
          T2.PutValue('PSA_PAIACOMPTE', TEtab.GetValue('ETB_PAIACOMPTE'));
          T2.PutValue('PSA_RIBACPSOC', TEtab.GetValue('ETB_RIBACOMPTE'));
          T2.PutValue('PSA_PAIFRAIS', TEtab.GetValue('ETB_PAIFRAIS'));
          T2.PutValue('PSA_RIBFRAISSOC', TEtab.GetValue('ETB_RIBFRAIS'));
          T2.PutValue('PSA_MOISPAIEMENT', TEtab.GetValue('ETB_MOISPAIEMENT'));
          T2.PutValue('PSA_JOURPAIEMENT', TEtab.GetValue('ETB_JOURPAIEMENT'));
          if (GetParamSoc('SO_PGCONGES')) and (TEtab.GetValue('ETB_CONGESPAYES') = 'X') then T2.PutValue('PSA_CONGESPAYES', 'X')
          else T2.PutValue('PSA_CONGESPAYES', '-');
          // Traitement de création automatique des tiers salariés
          if (GetParamSoc('SO_PGTYPENUMSAL') = 'NUM') and (GetParamSoc('SO_PGTIERSAUXIAUTO') = TRUE) then
          begin
            Libell := T2.GetValue('PSA_LIBELLE');
            LongLib := Length(Libell);
            Prenom := T2.GetValue('PSA_PRENOM');
            LongPre := Length(Prenom);
            Libell := Copy(Libell, 1, LongLib) + ' ' + Copy(Prenom, 1, LongPre);
            with InfTiers do
            begin
              Libelle := Copy(Libell, 1, 35);
              Adresse1 := T2.GetValue('PSA_ADRESSE1');
              Adresse2 := T2.GetValue('PSA_ADRESSE2');
              Adresse3 := T2.GetValue('PSA_ADRESSE3');
              Ville := T2.GetValue('PSA_VILLE');
              Telephone := T2.GetValue('PSA_TELEPHONE');
              CodePostal := T2.GetValue('PSA_CODEPOSTAL');
              Pays := '';
            end;
            CodeAuxi := CreationTiers(InfTiers, LeRapport, 'SAL', T2.GetValue('PSA_SALARIE'));
            if CodeAuxi <> '' then T2.PutValue('PSA_AUXILIAIRE', CodeAuxi);
          end;
        end // Fin du cas création des salaries
        else
        begin // Cas de la modification des salariés
          st := 'UPDATE SALARIES SET PSA_DATEENTREE="' + UsDateTime(T1.GetValue('PIM_DATEENTREE')) + '",';
          st := st + 'PSA_LIBELLE="' + T1.GetValue('PIM_LIBELLE') + '",';
          st := st + 'PSA_PRENOM="' + T1.GetValue('PIM_PRENOM') + '",';
          st := st + 'PSA_ADRESSE1="' + T1.GetValue('PIM_ADRESSE1') + '",';
          st := st + 'PSA_ADRESSE2="' + T1.GetValue('PIM_ADRESSE2') + '",';
          st := st + 'PSA_ADRESSE3="' + T1.GetValue('PIM_ADRESSE3') + '",';
          st := st + 'PSA_CODEPOSTAL="' + T1.GetValue('PIM_CODEPOSTAL') + '",';
          st := st + 'PSA_VILLE="' + T1.GetValue('PIM_VILLE') + '",';
          st := st + 'PSA_TELEPHONE="' + T1.GetValue('PIM_TELEPHONE') + '",';
          st := st + 'PSA_DATENAISSANCE="' + UsDateTime(T1.GetValue('PIM_DATENAISSANCE'))+'",';
          st := st + 'PSA_NATIONALITE="' + T1.GetValue('PIM_NATIONALITE') + '",';
          st := st + 'PSA_SEXE="' + T1.GetValue('PIM_SEXE') + '",';
          st := st + 'PSA_NUMEROSS="' + T1.GetValue('PIM_NUMEROSS') + '",';
          st := st + 'PSA_CIVILITE="' + T1.GetValue('CIVILITE') + '",';
          st := st + 'PSA_PORTABLE="' + T1.GetValue('PORTABLE') + '",';
          st := st + 'PSA_DATESORTIE="' + UsDateTime(T1.GetValue('DATESORTIE')) + '",';
          st := st + 'PSA_DATEMODIF="' + UsDateTime(Date) + '"';
          st := st + ' WHERE PSA_SALARIE="' + T1.GetValue('PIM_SALARIE') + '"';
          ExecuteSQL(St);
        end;
      end; // Fin de la boucle de traitement des salariés à créer
      TSalACreer.InsertDB(nil, false);
      TSalACreer.free;
//    end;
    if Tob_sal <> nil then
    begin
      Tob_sal.Free;
      Tob_sal := nil;
    end;
    // permet le rechargement des salariés nouvellement crées
    st := 'SELECT PSA_SALARIE,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SUSPENSIONPAIE,PSA_ETABLISSEMENT, ' +
      'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL,PSA_CONGESPAYES,' + { PT22 }
    'ETB_DATECLOTURECPN,ETB_CONGESPAYES FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT = ETB_ETABLISSEMENT';
    Q := opensql(st, true);
    if (not Q.eof) then
    begin
      Tob_sal := TOB.Create('Table des Salariés', nil, -1);
      if tob_sal <> nil then
        Tob_sal.loaddetaildb('INFOS SALARIES', '', '', Q, false, false, -1, 0);
    end;
    Ferme(Q);
    end;
  end;
end;

procedure TOF_PGIMPORTFIC.OnClose;
begin
  inherited;
  if Tob_SalATT <> nil then
  begin
    Tob_SalATT.Free;
    Tob_SalATT := nil;
  end;
end;
// FIN PT27

initialization
  registerclasses([TOF_PGIMPORTFIC]);

end.

