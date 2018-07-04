unit Paie_cjdc_lib;

interface

uses
{$IFNDEF eAGLClient}
  Db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF eAGLClient}
  Controls,
  Forms,
  StdCtrls,
  HMsgBox,
  Windows,
  FileCtrl,
  SysUtils,
  Classes,
  Inifiles,
  Dialogs,
  Consts,
  hctrls,
  HEnt1,
  UDossierSelect,
{$IFDEF EAGLCLIENT}
  galOutil, //cv4 le 05/10/06 -> PathDos ne fonctionne pas en eAGL !
{$ENDIF}
  uTob;
  
const
  MODE_LECTURE = 1;
  MODE_ECRITURE = 2;
  MODE_AJOUT = 3;
  MODE_CREATE = 0;
  MODE_APPEND = 1;
  MODE_READWRITE = 2;

  // Web Service
  TDFC = 1;
  TVA = 2;
  // a compléter pour social à partir de 10
  DUCS = 10; // à revoir
  DADSU = 11;

  // Récupération pièces sur JDC : Type de pièces
  ADS = '00';
  ACS = '01';
  ARS = '02';
  ACS_ARS = '98';
  TOUS = '99';
  // Récupération pièces sur JDC : statut
  NOUVELLES = '00';
  TOUTES = '01';

  RECUP = '-1'; //Récupération  pièce (WS)


Type

    TTRANSMISSION = record
    ModeEnvoi: string;
//    AliasDest: string;
//    CodeApp: string;
   	FichierChemin: string;   //chemin complet du fichier à envoyer
    FichierRequete: string;  // zone de travail
    FichierEnvoi: string;    // nom du fichier à envoyer
    Teleprocedure: Integer;
    idWS: string;
    mpWS: string;
    urlWS: string;
  end;

  INFOFICHIER = record
    Chemin: string; // Chemin d'accés
    NomFichier: string; // Nom du fichier
    FListe: TStringList; // Liste de chaines issues du fichier
    ModeAcces: Integer; // Mode d'accés au fichier
    NbLigne: LongWord; // Nombre de ligne texte
    PosCur: LongWord; // Position de l'indice de parcours
  end;

  RPConfig = record // copie de TDI_ConfigGlobale
     //--- Section Generalite
    RepTmp: string; //Répertoire de travail de l'utilisateur (Document and setting...)
    ViaWS: string; // 'True' si utilisation Web Service
    RecupWS: string; // 'True' si récupération automatique des ICR
  end;

function ChargerListeEmetteurs(ComboEmetteur: THVALCOMBOBOX; Item1: string): string;
//@@function InitInfoWS(P_SiretEmetteur: string; var Transmission: TTRANSMISSION): boolean;
function InitInfoWS(P_SiretEmetteur: string; var Transmission: TTRANSMISSION; Social : boolean = false): boolean;    //@@
function TDI_FinFichierCR(var Fichier: INFOFICHIER): Boolean;
function TDI_EcrireFichierCR(var Fichier: INFOFICHIER; var Chaine: string): Boolean;
function TDI_LireFichierCR(var Fichier: INFOFICHIER; var Chaine: string): Boolean;
function TDI_FermerFichierCR(var Fichier: INFOFICHIER): Boolean;
function TDI_OuvrirFichierCR(var Fichier: INFOFICHIER; PChemin: string;PNomFichier: string; PModeAcces: Integer): Boolean;
function TDI_CopierFichier(CheminSource, CheminDestination: string): boolean;
function TDI_OuvrirFichier(var Fichier: TextFile; PChemin: string; PNomFichier: string; PModeAcces: Integer): Boolean;
function TDI_FermerFichier(var Fichier: TextFile): Boolean;
function TDI_EcrireFichier(var Fichier: TextFile; var Chaine: string): Boolean;
function TDI_LireFichier(var Fichier: TextFile; var Chaine: string): Boolean;
function TDI_ObtenirDateFormat(): string;
function TDI_Executer(ChParametre: string): Boolean;
function TDI_ObtenirDateHeure(): string;
function EcrireTabletteConfig(Tablette, Code, Libelle, Abrege, Libre: string): boolean; //@@


implementation

{----------------------------------------------------------
 --- Nom       : FONCTION TDI_Executer
 --- Objet     : Execute un process
 ----------------------------------------------------------}

function TDI_Executer(ChParametre: string): Boolean;
var
  Si: TStartupInfo;
  Pi: TProcessInformation;
  Resultat: Boolean;
begin
  ZeroMemory(@si, sizeof(si));
  si.cb := sizeof(si);
  Resultat := CreateProcess(nil, PChar(ChParametre), nil, nil, true, 0, nil, nil, si, pi);
  WaitForSingleObject(pi.hprocess, INFINITE);
  TDI_Executer := Resultat;
end;

function TDI_OuvrirFichierCR(var Fichier: INFOFICHIER; PChemin: string;
  PNomFichier: string; PModeAcces: Integer): Boolean;
var
  Resultat: Boolean;
begin
  Resultat := False;
  with Fichier do
  begin
    //--- Initialisation des données
    if PChemin = '' then Chemin := PNomFichier
    else
      Chemin := PChemin + '\' + PNomFichier;
    NomFichier := PNomFichier;
    FListe := TStringList.Create;
    ModeAcces := PModeAcces;
    NbLigne := 0;
    PosCur := 0;
    //--- Ouverture du fichier en mode lecture ou ajout
    if ((ModeAcces = MODE_LECTURE) and (FileExists(Chemin))) then
    begin
      FListe.loadFromFile(Chemin);
      NbLigne := FListe.Count;
      Resultat := True;
      PosCur := 0;
    end;

    //   if ((ModeAcces=MODE_AJOUT) and (FileExists (Chemin))) then
    if (ModeAcces = MODE_AJOUT) then
    begin
      if (FileExists(Chemin)) then
      begin
        FListe.loadFromFile(Chemin);
        NbLigne := FListe.Count;
      end
      else
        NbLigne := 0;
      PosCur := NbLigne;
      Resultat := True;
    end;

    //--- Ouverture du fichier en mode ecriture
    if (ModeAcces = MODE_ECRITURE) then
    begin
      NbLigne := 0;
      PosCur := 0;
      Resultat := True;
    end;
  end;
  TDI_OuvrirFichierCR := Resultat;
end;

{-----------------------------------------------------------
 --- Nom       : FONCTION TDI_FermerFichierCR
 --- Objet     : Libère la mémoire. Transfert les données
 ---             de la mémoire vers le fichier.
 --- Retour    : True si l'opération a réussie
 ---           : False si l'opération a échouée
 -----------------------------------------------------------}

function TDI_FermerFichierCR(var Fichier: INFOFICHIER): Boolean;
begin
  with Fichier do
  begin
    if ((ModeAcces = MODE_ECRITURE) or (ModeAcces = MODE_AJOUT)) then
      FListe.SavetoFile(Chemin);
    FListe.Free;
  end;
  TDI_FermerFichierCR := True;
end;

{-----------------------------------------------------------------
 --- Nom       : FONCTION TDI_LireFichierCR
 --- Objet     : Lit une chaine de caractère. Aprés la lecture
 ---             l'indice est automatiquement augmenté de 1.
 ---             Si l'indice est inférieur à 0 ou s'il est
 ---             supérieur au nombre d'éléments contenu dans la
 ---             liste la fonction renvoie une chaine vide et
 ---             l'indice n'est pas augmenté.
 --- Retour    : True si l'opération a réussie
 ---           : False si l'opération a échouée
 -----------------------------------------------------------------}

function TDI_LireFichierCR(var Fichier: INFOFICHIER; var Chaine: string): Boolean;
var
  Resultat: Boolean;
begin
  with Fichier do
    if ((PosCur >= 0) and (PosCur <= NbLigne)) then
    begin
      Chaine := FListe[PosCur];
      inc(PosCur);
      Resultat := True;
    end
    else
    begin
      Chaine := '';
      Resultat := False;
    end;
  TDI_LireFichierCR := Resultat;
end;

{---------------------------------------------------
 --- Nom       : FONCTION TDI_EcrireFichierCR
 --- Objet     : Ecrit une chaine de caractère
 --- Retour    : True si l'opération a réussie
 ---           : False si l'opération a échouée
 ---------------------------------------------------}


function TDI_EcrireFichierCR(var Fichier: INFOFICHIER; var Chaine: string): Boolean;
var
  Resultat: Boolean;
begin
  with Fichier do
    if ((ModeAcces = MODE_ECRITURE) or (ModeAcces = MODE_AJOUT)) then
    begin
      PosCur := FListe.Add(Chaine);
      Resultat := True;
    end
    else
      Resultat := False;
  TDI_EcrireFichierCR := Resultat;
end;

{------------------------------------------------------
 --- Nom       : FONCTION TDI_FinFichierCR
 --- Objet     : Teste si on est à la fin du fichier
 --- Retour    : True si l'opération a réussie
 ---           : False si l'opération a échouée
 ------------------------------------------------------}

function TDI_FinFichierCR(var Fichier: INFOFICHIER): Boolean;
var
  Resultat: Boolean;
begin
  Resultat := False;
  with Fichier do
    if (PosCur = NbLigne) then
      Resultat := True;
  TDI_FinFichierCR := Resultat;
end;

function TDI_ObtenirDateFormat(): string;
var
  DateActuelle: TDateTime;
  DateChaine: string;
  Annee, Mois, Jour: Word;
begin
  DateActuelle := Now;
  DecodeDate(DateActuelle, Annee, Mois, Jour);
  DateChaine := Format('%2.2d', [Jour]) + '/' + Format('%2.2d', [Mois]) + '/' + IntToStr(Annee);
  Tdi_ObtenirDateFormat := DateChaine;
end;

function TDI_CopierFichier(CheminSource, CheminDestination: string): boolean;
var
  Buffer: Pointer;
  Resultat: Boolean;
  NbOctetCopie: Longint;
  FichierSource, FichierDestination: Integer;
const
  BufferTaille: longint = 8192;
begin
  Resultat := False;
  GetMem(Buffer, BufferTaille);
  try
    FichierSource := FileOpen(CheminSource, fmShareDenyWrite);
    if FichierSource < 0 then
      PgiInfo('erreur d''ouverture fichier source', CheminSource);
//      raise EFOpenError.CreateFmt(SFOpenError, [CheminSource]);
    try
      FichierDestination := FileCreate(CheminDestination);
      if FichierDestination < 0 then
      PgiInfo('erreur de création fichier destination', CheminDestination);
//        raise EFCreateError.CreateFmt(SFCreateError, [CheminDestination]);
      try
        repeat
          NbOctetCopie := FileRead(FichierSource, Buffer^, BufferTaille);
          if NbOctetCopie > 0 then
            if (FileWrite(FichierDestination, Buffer^, NbOctetCopie) <> -1) then
              Resultat := True
            else
              Resultat := False;
        until NbOctetCopie < BufferTaille;
      finally
        FileClose(FichierDestination);
      end;
    finally
      FileClose(FichierSource);
    end;
  finally
    FreeMem(Buffer, BufferTaille);
  end;
  Tdi_CopierFichier := Resultat;
end;

function TDI_OuvrirFichier(var Fichier: TextFile; PChemin: string; PNomFichier: string; PModeAcces: Integer): Boolean;
var
  ChaineErreur, CheminSource: string;
  Resultat: boolean;
begin
{$I-}
  Resultat := True;
  if (pChemin <> '') then
    CheminSource := PChemin + '\' + PNomFichier
  else
    CheminSource := PNomFichier;

  AssignFile(Fichier, CheminSource);
  filemode := 2;

  case PModeAcces of
    MODE_CREATE: Rewrite(Fichier);
    MODE_READWRITE: Reset(Fichier);
    MODE_APPEND: Append(Fichier);
  end;
{$I+}

  if (IOResult <> 0) then
  begin
    ChaineErreur := 'Impossible d''ouvrir le fichier ' + CheminSource + '.';
    MessageDlg(ChaineErreur, mtError, [mbOk], 0);
    Resultat := False;
  end;

  TDI_OuvrirFichier := Resultat;
end;

{-----------------------------------------------------------
 --- Nom       : FONCTION TDI_FermerFichier
 --- Objet     : Ferme un fichier
 --- Retour    : True si l'opération a réussie
 ---           : False si l'opération a échouée
 ------------------------------------------------------------}

function TDI_FermerFichier(var Fichier: TextFile): Boolean;
var
  Resultat: Boolean;
begin
{$I-}
  CloseFile(Fichier);
{$I+}
  if IOResult = 0 then
    Resultat := True
  else
    Resultat := False;
  Tdi_FermerFichier := Resultat;
end;

{------------------------------------------------------------
 --- Nom       : FONCTION TDI_LireFichier
 --- Objet     : Lit une chaine de caractère de caractère
 --- Retour    : True si l'opération a réussie
 ---           : False si l'opération a échouée
 ------------------------------------------------------------}

function TDI_LireFichier(var Fichier: TextFile; var Chaine: string): Boolean;
var
  Resultat: Boolean;
begin
{$I-}
  Read(Fichier, Chaine);
{$I+}
  if (IOResult <> 0) then
  begin
    MessageDlg('Erreur de lecture.', mtError, [mbOk], 0);
    Resultat := FALSE;
  end
  else
    Resultat := TRUE;
  TDI_LireFichier := Resultat;
end;

{------------------------------------------------------------
 --- Nom       : FONCTION EcrireFichier
 --- Objet     : Ecrit une chaine de caractère de caractère
 --- Paramètre : Fichier : Variable du flux
 ---             Chaine : Contenu de la chaine à écrire
 --- Retour    : True si l'opération a réussie
 ---           : False si l'opération a échouée
 ------------------------------------------------------------}

function TDI_EcrireFichier(var Fichier: TextFile; var Chaine: string): Boolean;
var
  Resultat: Boolean;
begin
{$I-}
  Write(Fichier, Chaine);
{$I+}
  if (IOResult <> 0) then
  begin
    MessageDlg('Erreur d''écriture.', mtError, [mbOk], 0);
    Resultat := FALSE;
  end
  else
    Resultat := TRUE;
  TDI_EcrireFichier := Resultat;
end;



function InitInfoWS(P_SiretEmetteur: string; var Transmission: TTRANSMISSION; Social : boolean = false): boolean;    //@@
var
  Q, Q1: TQuery;
  sSql: string;
begin
  Transmission.IdWS := '';
  Transmission.MpWS := '';
  Transmission.urlWS := '';
  Result := FALSE;
// d @@
  if (not Social) then
  begin
   // Selon le contenu de NJDC_COMPTES
   sSql := 'SELECT  * FROM NTDI_EMETTEUR, NJDC_COMPTES ';
   sSql := sSql + ' WHERE NEM_SIRET="' + P_SiretEmetteur + '" ';
   sSql := sSql + 'AND NEM_SIRET = NJC_INTITULE';
  end
  else
  begin
     sSql := 'SELECT  * FROM  NJDC_COMPTES ';
     sSql := sSql + ' WHERE NJC_INTITULE="' + P_SiretEmetteur + '" ';
  end;
// f @@  
  Q := OpenSQL(sSql, TRUE);
  if (not Q.Eof) and (Q.Recordcount = 1) then
  begin
    Transmission.IdWS := Q.FindField('NJC_NOMCONNEXION').AsString;
    Transmission.MpWS := Q.FindField('NJC_PASSWORD').AsString;
    Result := TRUE;
  end
  else
  	//05/03/2007 l'absence d'enregistrement n'est pas nécessairement une erreur
    if Q.RecordCount <> 0 then
    begin
   // plusieurs comptes sont définis dans NJDC_COMPTES pour un même émetteur
  // la tablette TDC dans choixEXT associe un nom de connexion au couple
   // SiretEmetteur+CodeUtilisateur
      sSql := 'SELECT YX_LIBELLE, NJC_NOMCONNEXION, NJC_PASSWORD FROM ##DP##.CHOIXEXT, NJDC_COMPTES';
      sSql := sSql + ' WHERE YX_TYPE = "TDC" AND YX_CODE ="' + P_SiretEmetteur + V_PGI.User + '"';
      sSql := sSql + ' AND NJC_NOMCONNEXION = YX_LIBELLE';

      Q1 := OpenSQL(sSql, TRUE);
      if not Q1.Eof then
      begin
				Result := TRUE;
        Transmission.IdWS := Q1.FindField('NJC_NOMCONNEXION').AsString;
        Transmission.MpWS := Q1.FindField('NJC_PASSWORD').AsString;
      end
      else
        PgiInfo('n''est pas spécifié : vérifier le paramétrage du site', 'Emetteur:' + P_SiretEmetteur + ', le compte JDC à utiliser pour l''utilisateur ' + V_Pgi.User);
      Ferme(Q1);
    end;
  Ferme(Q);

  if (Result = TRUE) AND  ((Transmission.IdWS = '') or (Transmission.MpWS = '')) then
  begin
    Result := FALSE;
    PgiInfo(' Connexion au Service Web JDC impossible', 'Comptes JDC invalides ou absents pour: '+ P_SiretEmetteur);
  end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : GCattenot
Créé le ...... : 21/05/2007
Modifié le ... :   /  /
Description .. : Avec TDI, pour un même site, on peut avoir plusieurs
Suite ........ : émetteurs. Cette liste tient compte des gproupes de
Suite ........ : confidentialité.
Suite ........ :
Suite ........ : A simplifier pour la paie ?
Mots clefs ... :
*****************************************************************}
function ChargerListeEmetteurs(ComboEmetteur: THVALCOMBOBOX; Item1: string): string;
var
  sSql, sSiret, sNom, Requete, sSqlEmetteur: string;
  sGrpConf, sGrp, YX_CODE: string;
  Q, Q1: TQuery;
  i, NbEmetteurs: integer;
begin
  result := '';
  sSqlEmetteur := ' AND NTM_SIRETEMETTEUR IN ('; //requête si <Tous>
    // récuperer les groupes de confidentialité de l'utilisateurs
  Q := OpenSQL('SELECT UCO_GROUPECONF FROM ##DP##.USERCONF WHERE UCO_USER="' + V_PGI.User + '"', True);
  try
    while not Q.Eof do
    begin
      sGrpConf := sGrpConf + Q.FindField('UCO_GROUPECONF').AsString + ';';
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
  // récupérer la tablette TDG qui donne les associations groupe+Siret Emetteur
  Q := OpenSQL('SELECT * FROM ##DP##.CHOIXEXT WHERE YX_TYPE = "TDG"', True); // liste des groupes de l'émetteur
  if not Q.EOF then
  begin
   // la confidentialité est en place
    if item1 <> '' then
    begin
      ComboEmetteur.Items.Add(Item1);
    end;

    while not Q.EOF do
    begin
      YX_CODE := Q.FindField('YX_CODE').AsString;
      sGrp := copy(YX_CODE, 1, 3) + ';';
      if pos(sGrp, SGrpConf) <> 0 then
      begin
     // ce groupe fait partie de ceux auxquels appartient l'utilisateur
        sSiret := copy(YX_CODE, 4, 14);
        sSql := 'SELECT ANN_NOM1 FROM NTDI_EMETTEUR, ANNUAIRE';
        sSql := sSql + ' WHERE NEM_SIRET ="' + sSiret + '" AND NEM_GUIDPER=ANN_GUIDPER ';

        Q1 := OpenSql(sSQl, true);
        if (not Q1.EOF) and (pos(sSiret, Result) = 0) then
        begin
          sNom := Q1.FindField('ANN_NOM1').AsString;
          Requete := ' AND NTM_SIRETEMETTEUR = "' + sSiret + '"';
          ComboEmetteur.Items.Add(sNom);
          ComboEmetteur.Values.Add(Requete);
          Result := Result + sSiret + ';';
          sSqlEmetteur := sSqlEmetteur + '"' + sSiret + '",'; // IN (....)
        end;
        Ferme(Q1);
      end;
      Q.Next;
    end;
    if item1 <> '' then
    begin
      sSqlEmetteur := sSqlEmetteur + ')'; // IN (... ",)
      // Enlever la dernière virgule.
      i := pos(',)', sSqlEmetteur);
      if i > 0 then
        sSqlEmetteur := copy(sSqlEmetteur, 0, i - 1) + ')'
      else
        sSqlEmetteur := '';
      ComboEmetteur.Values.Insert(0, sSqlEmetteur);
    end;
    if Result = '' then
      PgiInfo('Votre compte utilisateur ne vous permet pas d''effectuer des prises en compte', 'Vérifier votre groupe de travail avec ceux des fiches émetteurs');
  end
  else
  begin
   // pas de gestion de la confidentialité
    // Charger le combo à partir de la table EMETTEURS

    sSql := 'SELECT NEM_SIRET, ANN_NOM1 FROM NTDI_EMETTEUR, ANNUAIRE';
    sSql := sSql + ' WHERE NEM_GUIDPER=ANN_GUIDPER ';

    //10/10/2006
    //ComboEmetteur.Items.Add('<Tous>');
    //ComboEmetteur.Values.Add(' AND NTM_SIRETEMETTEUR <> ""');
    if item1 <> '' then
    begin
      ComboEmetteur.Items.Add(Item1);
      ComboEmetteur.Values.Add(' AND NTM_SIRETEMETTEUR <> ""')
    end;

    Q := OpenSql(sSQl, true);
    while not Q.EOF do
    begin
      sSiret := Q.FindField('NEM_SIRET').AsString;
      sNom := Q.FindField('ANN_NOM1').AsString;
      Requete := ' AND NTM_SIRETEMETTEUR = "' + sSiret + '"';
      ComboEmetteur.Items.Add(sNom);
      ComboEmetteur.Values.Add(Requete);

      Result := Result + sSiret + ';'; //cv4 le 04/10/06
      Q.Next;
    end;
    Ferme(Q);
  end;
  ComboEmetteur.ItemIndex := 0;
  ComboEmetteur.Text := ComboEmetteur.items[0];
end;

function TDI_ObtenirDateHeure(): string;
var
  DateActuelle: TDateTime;
  DateChaine, HeureChaine: string;
  Annee, Mois, Jour, Heure, Minute, Seconde, MilSeconde: Word;
begin
  DateActuelle := Now;
  DecodeDate(DateActuelle, Annee, Mois, Jour);
  DateChaine := Format('%4.4d', [Annee]) + Format('%2.2d', [Mois]) + Format('%2.2d', [Jour]);

  DecodeTime(DateActuelle, Heure, Minute, Seconde, MilSeconde);
  HeureChaine := Format('%2.2d', [Heure]) + Format('%2.2d', [Minute]);
  Result := DateChaine + HeureChaine;
end;
// @@  pour la PAIE : récupéré de TDI_LibFonction
function EcrireTabletteConfig(Tablette, Code, Libelle, Abrege, Libre: string): boolean;
var
  sSql: string;
begin
{
 pas possiblr d'utiliser une tob en forçant lecture de la db0
}
  if Libelle = '' then Libelle := ' ';
  if Abrege = '' then Abrege := ' ';

  if ExisteSql('SELECT 1 FROM ##DP##.CHOIXEXT WHERE YX_TYPE="' + Tablette + '" AND YX_CODE="' + Code + '"') then
  begin
    sSQl := 'UPDATE ##DP##.CHOIXEXT';
    sSql := sSql + ' SET YX_TYPE="' + Tablette + '",';
    sSql := sSql + ' YX_CODE="' + Code + '",';
    sSql := sSql + ' YX_LIBELLE="' + Libelle + '",';
    sSql := sSql + ' YX_ABREGE="' + Abrege + '",';
    sSql := sSql + ' YX_LIBRE="' + Libre + '"';
    sSql := sSql + ' WHERE YX_TYPE="TDI" AND YX_CODE="' + Code + '"';
  end
  else
  begin
    sSql := 'INSERT INTO  ##DP##.CHOIXEXT';
    sSql := sSql + '(YX_TYPE, YX_CODE, YX_LIBELLE, YX_ABREGE, YX_LIBRE) VALUES (';
    sSql := sSql + '"' + Tablette + '","' + code + '","' + Libelle + '","' + Abrege + '","' + Libre + '")';
  end;
  ExecuteSQL(sSql);

  Result := TRUE;
end;
end.

