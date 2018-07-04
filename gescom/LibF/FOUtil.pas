{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Fonctions utilitaires pour le Front Office
Mots clefs ... : FO
*****************************************************************}
unit FOUtil;

interface
uses
  Windows, Forms, Graphics, Classes, ExtCtrls, Controls, comctrls, registry,
  SysUtils, ShellAPI, Hent1, ParamSoc, LicUtil, Math, Buttons, stdctrls, FileCtrl,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  FE_Main, dbtables, MajTable,
  {$ENDIF}
  UTob, Hctrls, ED_Tools, HStatus, HMsgBox, HTB97, HMenu97, HDimension, Hqry,
  {$IFDEF FOS5}
  MC_Erreur,
  {$ENDIF}
  {$IFDEF CHR}
  EntCHR,
  {$ENDIF}
  FODefi;

///////////////////////////////////////////////////////////////////////////////////////
//  déclarations des composants.
///////////////////////////////////////////////////////////////////////////////////////
type TFOProgressForm = class(TComponent)
  private
    FTimer: TTimer;
    procedure IsTimer(Sender: TObject);
  public
    constructor Create(AOwner: TWinControl; Caption, Titre: string); reintroduce; overload; virtual;
    destructor Destroy; override;
  end;

  ///////////////////////////////////////////////////////////////////////////////////////
  // Liste des options de connexion à une caisse
  ///////////////////////////////////////////////////////////////////////////////////////
type TTypeOptConnCaisseFO = (cfoChaqueFois, cfoToutesLignes, cfoVerifVerrou);

  ///////////////////////////////////////////////////////////////////////////////////////
  // Liste des compteurs d'une journée de vente
  ///////////////////////////////////////////////////////////////////////////////////////
type TTypeCptJourCaisseFO = (jfoNbTicAbandon, jfoNbTicAnnul, jfoNbTicMisatt, jfoNbTicattRepri,
    jfoNbTicattSup, jfoNbLigAnnul, jfoNbModifMdp, jfoNbOuvTiroir, jfoNbRatcli);

  ///////////////////////////////////////////////////////////////////////////////////////
  //  déclarations des fonctions et procédures.
  ///////////////////////////////////////////////////////////////////////////////////////
function FOCadreGauche(Texte: string; Taille: Integer): string;
function FOCadreDroite(Texte: string; Taille: Integer): string;
function FODecaleGauche(Texte: string; NbCar, Taille: Integer): string;
function FODecaleDroite(Texte: string; NbCar, Taille: Integer): string;
function FOGetEnvironnement(const NomVar: string): string;
function FOLanceProg(const Cmd: string; const ShowCmd: Integer = SW_SHOWNORMAL): Boolean;
function FOIsDirectory(FName: string): Boolean;
procedure FOLanceClavierVisuel;
function FOTimeToFrench(Heure: string): string;
function FOFrenchToTime(Heure: string): string;
function FODonneAn(Date: TDateTime; Taille: Integer = 4): string;
function FOEstVrai(NomTob: TOB; NomChamp: string): Boolean;
function FOEstFaux(NomTob: TOB; NomChamp: string): Boolean;
{$IFDEF FOS5}
function FOAvantConnexion(var SortieHalley: boolean): boolean;
procedure FOApresConnexion;
procedure FOAvantDeConnexion;
procedure FOApresDeConnexion;
{$ENDIF}
function FOCaisseIsMonoPoste: boolean;
function FOEtatCaisse(Caisse: string; Recharge: boolean = True): string;
function FOEtatJournee(Caisse: string; NumZ: Integer): string;
function FOCaisseCourante: string;
function FOCaisseDisponible: Boolean;
function FOCaisseEtab(Etab: string): string;
function FOVerifieVHGCCaisse: Boolean;
procedure FOChargeVHGCCaisse(Caisse: string);
function FOGetDeviseCaisse(Caisse: string; Symbole: boolean): string;
function FOGetDeviseEtab(Etablissement: string; Symbole: boolean): string;
{$IFDEF FOS5}
procedure FOChargeMaterielCaisse(Caisse: string);
procedure FOLibereMaterielCaisse;
{$ENDIF}
function FOGetNumZCaisse(Caisse: string; Borne: string = 'MAX'): Integer;
procedure FOGetInfoNumZ(Caisse: string; NumZ: Integer; var DateOuv: TDateTime; var ExceptionTaxe: string);
function FOGetDateOuv(Caisse: string; NumZ: Integer): TDateTime;
function FOSelectNumZDate(Prefixe: string; DeNumZ, ANumZ: integer; Caisse: string = ''): string;
procedure FOBorneNumZDate(DeNumZ, ANumZ: integer; var DeDateOuv, ADateOuv: TDateTime; Caisse: string = '');
procedure FOReChargeVHGCCaisse;
procedure FOForceVHGCCaisse;
function FOLockCaisse(Caisse: string): Boolean;
function FOLibereCaisse(Caisse: string): Boolean;
function FOLibereVHGCCaisse: Boolean;
function FODemarqueCaisse(Caisse: string): Boolean;
function FOMonoCaisse(Ouverte: Boolean = False; Libre: Boolean = False): Boolean;
function FOOptionConnexionCaisse(Option: TTypeOptConnCaisseFO): Boolean;
function FOGetParamCaisse(NomChamp: string; TypeChamp: string = ''): variant;
procedure FOPutParamCaisse(NomChamp: string; Valeur: Variant);
function FOLogoCaisse: string;
function FONbExemplaire(TypeDoc: string = 'BON'): Integer;
function FOGetNatureTicket (EnAttente, PourSql: boolean): string;
procedure FOPositionneCaisseUser (cbCaisse, cbEtab: THValComboBox; ForceEtabDefaut: boolean = False);
function FOIncrJoursCaisse(TypeCpt: TTypeCptJourCaisseFO; MsgEvent: string = ''; Increment: Integer = 1): Boolean;
function FODonneArrondiRemise: string;
function FOArrondiRemise(MntRemise, MntBrut: Double): Double;
function FOVerifMaxRemise(TauxRemise: double; TypeRemise: string; TypeUniquement: boolean; var MaxRem: integer; var CodeDemValide: string): integer;
function FOSaisiePwd(NomChamp: array of string): Boolean;
function FOCompareMontant(Montant1, Montant2: double; Operateur: string = '='): boolean;
procedure FOChgStatus;
function FOMsgBoxExec(MsgBox: THMsgBox; NoMsg: integer; Av, Ap: string; DonneTitre: boolean = False): string;
function FOComboValueExist(Combo: THValComboBox; Value: string): Boolean;
function FOComboDeleteValue(Combo: THValComboBox; Value: string): Boolean;
function FOTabletteValueExist(Tablette: string; Value: string; Plus: string = ''): Boolean;
function FOFabriqueSQLIN(Liste, NomChamp: string; WithAnd, NotIn, ChampNumeric: Boolean): string;
function FOFabriqueListeMDP(NomListe, NomChamp: string): string;
function FOFabriqueListeTypArt(NaturePieceG, NomChamp: string): string;
function FOControlChamp(NomChamp: string): string;
function FOTestControlChamp(NomChamp, Flag: string): Boolean;
function FOChampValeurVide(NomChamp: string; TypeChamp: string = ''): variant;
function FOTestChampVide(NomChamp: string; ValeurChamp: Variant): Boolean;
procedure FOSimuleClavier(Touche: Word; Shift: Boolean = False; Attente: Integer = 0);
function FOExisteClavierEcran: Boolean;
procedure FORecopieTob(TOBOrg: TOB; var TOBDest: TOB);
procedure FODepileTOBGrid(GS: THGrid; TOBG: TOB; ACols: array of Integer);
procedure FOMergeTobFille(TOBPrinc, TOBSecond: TOB; NomCle: string; NouveauChamp: array of string);
procedure FOMergeTob(TOBPrinc, TOBSecond: TOB; OverWrite: Boolean = False);
procedure FOPutValueDetail(TOBPrinc: TOB; ChampMaj: string; VV: Variant);
procedure FOMajChampSupValeur(TOBL: TOB; NomChamp: string; Value: Variant);
function FOTesteChampSupValeur(TOBL: TOB; NomChamp: string; Value: variant): boolean;
function FOChampSupValeurNonVide(TOBL: TOB; NomChamp: string): boolean;
function FOGetChampSupValeur(TOBL: TOB; NomChamp: string; Value: variant): variant;
function FOExisteTPE: Boolean;
function FODonneParamTPE(Caisse: string; var Tpe, Port, Param: string): Boolean;
function FOExisteAfficheur: Boolean;
function FODonneParamAfficheur(var Afficheur, Port, Param: string): Boolean;
function FOExisteTiroir: Boolean;
function FODonneParamTiroir(var Tiroir, Port, Param: string): Boolean;
function FOVerifOuvreTiroir(TOBEches: TOB): Boolean;
procedure FOOuvreTiroir(FromUser: Boolean = False; FinJour: Boolean = False);
function FOExisteLP: Boolean;
function FOExisteLPCheque: Boolean;
function FODonneTypeLP: string;
function FODonneCodeEtat(TypeEtat: TTypeEtatFO; ValeurDefaut: Boolean; var Nature, CodeEtat, Titre: string; CodeOk : boolean=false): Boolean;
function FOExisteCodeEtat(TypeEtat: TTypeEtatFO): Boolean;
function FODonneParamLP(var Imprimante, Port, Param: string): Boolean;
procedure FOLanceImprimeLP(TypeEtat: TTypeEtatFO; Where: string; Apercu: Boolean; UneTOB: TOB; Code : string='');
{$IFDEF FOS5}
function FOImprimeLP(TypeEtat, NatEtat, CodeEtat, Where, Imprimante, Port, Params: string; Initialise, PreView, SQLBased: Boolean; UneTOB: TOB; var Err:
  TMC_Err; IsImpControl: Boolean = FALSE; IsTestMateriel: Boolean = FALSE): Boolean; //XMG 22/06/01
procedure FOMCErr(Erreur: TMC_Err; Titre: string);
{$ENDIF}
procedure FODispatchParamModele(Num: integer; Action: TActionFiche; Lequel, TT, Range: string);
function FOMakeWhereTicketX(Prefixe: string; Caisse: string = ''): string;
function FOMakeWhereTicketZ(Prefixe, Caisse: string; DeNumZ, ANumZ: integer): string;
function FOMakeWhereRecapVendeurs(Prefixe, Caisse: string; DeNumZ, ANumZ: Integer): string;
procedure FOSetFontSize(Composant: TComponent; Action: string; Taille: Integer);
procedure FOSetFontColor(FF: TForm; NomComp: string; Color: TColor);
procedure FoShowEtiquette(F: Forms.TForm; Visu: boolean);
function FOIsInteger(st: string): string;
function FOReverseSt(st: string): string;
function FOCutspaceSt(st: string): string;
function FOExtract(var Stg: string; Index, Count: Integer): string;
function FOStrCmp(Chaine, Liste: string; GereTous: Boolean = False): Boolean;
function FOStrFMontant(Montant: Extended; Taille: Integer; Devise: string): string;
function FoExtractWhere(PageCritere: TPageControl): string;
function FOExisteSQL(sSql: string): Boolean;
function FOChargeDevise(pD_: TOB; Code: string): boolean;
function FODonneDeviseNbDec(Code: string): Integer;
function FODonneDeviseSymbole(Code: string): string;
function FOChargeDeviseEuro: string;
function FODonneCodeEuro: string;
function FOTestCodeEuro(CodeDev: string): Boolean;
function FODonneMaskDevise(CodeDev: string): string;
function FODonneRendu: string;
function FODonneEcartChange: string;
function FODonneClientDefaut: string;
function FOTiersEstFictif(TOBTiers: TOB): boolean;
function FOGereDemarque(Obligatoire: Boolean): Boolean;
procedure FOChargeDemarque;
procedure FOChargeArtFiTiers; // JTR Lien Opération Caisse avec Tiers
function FOGetPTobDemarque: TOB;
function FOGetPTobArtFiTiers: TOB; // JTR Lien Opération Caisse avec Tiers
function FODecodeLibreDemarque(TOBDem: TOB; Option: string): string;
function FOIdentArticleGen(CodeArt: string): string;
function FOAfficheImageArt: Boolean;
procedure FOSaveInRegistry(LocalKey, NomChamp: string; Valeur: Variant);
function FOGetFromRegistry(LocalKey, NomChamp: string; ValeurDefaut: Variant): Variant;
procedure FODeleteFromRegistry(LocalKey, NomChamp: string);
function FOChangeBase(Valeur: string; Base: Byte; Taille: Integer): string;
function FOAlphaCodeNumeric(const Code: string): string;
procedure FOChargeConfidentialite(NoMenu: integer);
function FOJaiLeDroit(CodeEvenement: integer; Parle: boolean = True; ChgUser: boolean = True; ErrorMsg: string = ''; UserGrp: Integer = 0): boolean;

  ///////////////////////////////////////////////////////////////////////////////////////
  //  déclarations des types, fonctions et procédures issus de S1 pour le pavé tactile
  ///////////////////////////////////////////////////////////////////////////////////////
type TPGIErrS1 = record
   code, col, Row, SysError: integer;
   Libelle, WavFile, Champ, RefreshQ: string;
   Avertir: boolean;
   end;
function KeytoChar(Key: Word; Shift: TShiftState): Char;
procedure DisableControlsFiche(F: TCustomForm; FCaisse: TWinControl; Disable, OkEnableCaisse: Boolean; OkCaisse: Boolean = FALSE);
function AbsoluteGoForward(Ctrl1, Ctrl2: TWinControl): Boolean;
function MsgErrDefautS1(Code: integer): string;
function CountTokenPipe(LargeSt, Sep: string): Integer;
{$IFNDEF EAGLCLIENT}
function OpenSQLDB(SQL: string; DB: TDatabase): TQuery;
{$ENDIF}
function FOVerifieVendeurPave (Etablis: string; Caisse: array of string; RemplaceAuto: boolean): boolean;
function IsArtFiPce(TobPiece : TOB) : boolean;

implementation
uses
  {$IFDEF FOS5}
  MC_Admin, LP_View, TR_BASE, LP_impri, EntFO,
  {$ENDIF}
  {$IFNDEF EAGLCLIENT}
  EdtEtat, Edtdoc, LP_Param, LP_Base,
  {$ENDIF}
  MC_Lib, Ent1, EntGC, UtilGC, TarifUtil;

const Liste_delimiters: string = ',.;:-_/ '; // liste des separateurs

type TFOFontControl = class(TControl)
  public
    property Font;
  end;

  {==============================================================================================}
  {================================= FORMATAGE D'UN CHAMP =======================================}
  {==============================================================================================}
  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 23/07/2001
  Modifié le ... : 23/07/2001
  Description .. : Cadrage à gauche d'un texte
  Mots clefs ... : FO
  *****************************************************************}

function FOCadreGauche(Texte: string; Taille: Integer): string;
begin
  Result := Texte;
  if Taille > Length(Texte) then Result := Result + StringOfChar(' ', Taille - Length(Texte));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Cadrage à droite d'un texte
Mots clefs ... : FO
*****************************************************************}

function FOCadreDroite(Texte: string; Taille: Integer): string;
begin
  Result := '';
  if Taille > Length(Texte) then Result := StringOfChar(' ', Taille - Length(Texte));
  Result := Result + Texte;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 10/08/2001
Description .. : Décalage de n caractères d'un texte de la droite vers la
Suite ........ : gauche dans une 'fenêtre' de taille fixe.
Mots clefs ... : FO
*****************************************************************}

function FODecaleGauche(Texte: string; NbCar, Taille: Integer): string;
begin
  Result := Copy(Texte, NbCar + 1, Length(Texte) - NbCar);
  if Taille > Length(Texte) then Result := Result + StringOfChar(' ', Taille - Length(Texte));
  Result := Result + Copy(Texte, 1, NbCar);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 10/08/2001
Description .. : Décalage de n caractères d'un texte de la gauche vers la
Suite ........ : droite dans une 'fenêtre' de taille fixe.
Mots clefs ... : FO
*****************************************************************}

function FODecaleDroite(Texte: string; NbCar, Taille: Integer): string;
begin
  if Taille > Length(Texte) then Texte := Texte + StringOfChar(' ', Taille - Length(Texte));
  Result := Copy(Texte, Length(Texte) - NbCar + 1, NbCar);
  Result := Result + Copy(Texte, 1, Length(Texte) - NbCar);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOEstSeparateur : Vérifie si un caractère est un séparateur
///////////////////////////////////////////////////////////////////////////////////////

function FOEstSeparateur(const Texte: string; const Pos: integer): Boolean;
begin
  Result := IsDelimiter(Liste_delimiters, Texte, Pos);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOSupDoubleSep : Supprime les séparateurs redondants dans un texte.
///////////////////////////////////////////////////////////////////////////////////////

function FOSupDoubleSep(const Texte: string): string;
var i, j: Integer;
  SupEspaceSuivant: Boolean;
begin
  Result := Texte;
  SupEspaceSuivant := False;
  for i := 1 to Length(Result) do
  begin
    if (SupEspaceSuivant) and (FOEstSeparateur(Result, i)) then
    begin
      j := i;
      while (j <= Length(Result)) and (FOEstSeparateur(Result, j)) do inc(j);
      Delete(Result, i, (j - i));
    end;
    SupEspaceSuivant := (FOEstSeparateur(Result, i));
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOSupDoubleEspace : Supprime les espaces redondants dans un texte.
///////////////////////////////////////////////////////////////////////////////////////

function FOSupDoubleEspace(const Texte: string): string;
var i, j: Integer;
  SupEspaceSuivant: Boolean;
begin
  Result := Texte;
  SupEspaceSuivant := False;
  for i := 1 to Length(Result) do
  begin
    if (SupEspaceSuivant) and (Result[i] = ' ') then
    begin
      j := i;
      while (j <= Length(Result)) and (Result[j] = ' ') do inc(j);
      Delete(Result, i, (j - i));
    end;
    SupEspaceSuivant := (Result[i] = ' ');
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FODebutMotMajuscule : Passe en majuscule le début de chaque mot d'une taille minimum.
///////////////////////////////////////////////////////////////////////////////////////

function FODebutMotMajuscule(const Texte: string; const TailleMiniMot: integer; const MajMiniMot: boolean): string;
var i, j: Integer;
begin
  Result := LowerCase(Texte);
  i := 1;
  while i <= Length(Result) do
  begin
    if (Result[i] <> ' ') then
    begin
      j := i;
      while (j <= Length(Result)) and not (FOEstSeparateur(Result, j)) do inc(j);
      if ((j - i) > TailleMiniMot) then
      begin
        while (i < j) and not (Result[i] in ['a'..'z']) do inc(i);
        Result[i] := UpCase(Result[i]);
        i := j;
      end else
        if MajMiniMot then
      begin
        while (i < j) do
        begin
          if (Result[i] in ['a'..'z']) then Result[i] := UpCase(Result[i]);
          inc(i);
        end;
      end;
    end;
    inc(i);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOMetTexteEnForme : Met en forme un texte.
///////////////////////////////////////////////////////////////////////////////////////

function FOMetTexteEnForme(const Texte, Format: string; const TailleMiniMot: integer; const MajMiniMot, SupEsp: boolean): string;
begin
  Result := Texte;
  if SupEsp then result := FOSupDoubleSep(Result);
  if Format = 'MAJ' then Result := AnsiUpperCase(Result) else  // JTR - eQualité 10389
    if Format = 'MIN' then Result := AnsiLowerCase(Result) else // JTR - eQualité 10389
    if Format = 'PRE' then Result := FirstMajuscule(LowerCase(Result)) else
    if Format = 'MOT' then Result := FODebutMotMajuscule(Result, TailleMiniMot, MajMiniMot);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOAppliqueFormat : Applique à un champ le format défini dans les paramètres société.
///////////////////////////////////////////////////////////////////////////////////////

function FOAppliqueFormat(NomChamp, ValeurChamp: string): variant;
var chp, fmt: string;
  maj, sup: boolean;
  min: integer;
begin
  Result := ValeurChamp;

  if NomChamp = 'T_LIBELLE' then chp := 'SO_FMTTIERSLIBELLE' else
    if NomChamp = 'T_PRENOM' then chp := 'SO_FMTTIERSPRENOM' else
    if NomChamp = 'T_ABREGE' then chp := 'SO_FMTTIERSABREGE' else exit;

  // En attendant de mettre en place les paramètres société !!!!
  fmt := 'MOT'; // varAsType(GetParamSoc(chp), varString);
  sup := True; // varAsType(GetParamSoc('SO_FMTTIERSSUPESPACE'), varBoolean);
  min := 3; // varAsType(GetParamSoc('SO_FMTTIERSTAILLEMOT'), varInteger);
  maj := False; // varAsType(GetParamSoc('SO_FMTTIERSMAJMINIMOT'), varBoolean);

  Result := FOMetTexteEnForme(Result, fmt, min, maj, sup);
end;

{==============================================================================================}
{=================================== APPELS SYSTEME ===========================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne la valeur d'une variable d'environnement
Suite ........ : Windows
Mots clefs ... : FO
*****************************************************************}

function FOGetEnvironnement(const NomVar: string): string;
var Buffer: array[0..1028] of Char;
  Taille: Integer;
begin
  Result := '';
  Taille := GetEnvironmentVariable(PChar(NomVar), Buffer, SizeOf(Buffer));
  if Taille > 0 then SetString(Result, Buffer, Taille);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Lance un programme
Mots clefs ... : FO
*****************************************************************}

function FOLanceProg(const Cmd: string; const ShowCmd: Integer = SW_SHOWNORMAL): Boolean;
var WPtr: hWnd;
begin
  Result := False;
  if FileExists(Cmd) then
  begin
    WPtr := GetActiveWindow;
    Result := (ShellExecute(WPtr, nil, PChar(Cmd), nil, nil, ShowCmd) > 32);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 07/11/2001
Modifié le ... : 07/11/2001
Description .. : Vérifie si un nom correspond à un répertoire
Mots clefs ... : FO
*****************************************************************}

function FOIsDirectory(FName: string): Boolean;
begin
  Result := ((FName <> '') and ((FileGetAttr(FName) and faDirectory) <> 0));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 05/05/2004
Modifié le ... : 05/05/2004
Description .. : Crée un sémaphore pour indiquer qu'une instance du
Suite ........ : programme est en cours d'exécution
Mots clefs ... : FO
*****************************************************************}

{$IFDEF FOS5}
function FOS5isRunning : boolean;
begin
  if VH_FO.hSemaphoreFORun = 0 then
  begin
    SetLastError(0) ;
    VH_FO.hSemaphoreFORun := CreateSemaphore(nil, 0, 1, PChar('FrontOfficeCegid'));
    Result := (GetLastError = ERROR_ALREADY_EXISTS);
  end else
    Result := False;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 05/05/2004
Modifié le ... : 05/05/2004
Description .. : Supprime le sémaphore qui indique qu'une instance du
Suite ........ : programme est en cours d'exécution
Mots clefs ... : FO
*****************************************************************}

{$IFDEF FOS5}
procedure FOS5isStopping;
begin
  if VH_FO.hSemaphoreFORun <> 0 then
  begin
    CloseHandle(VH_FO.hSemaphoreFORun);
    VH_FO.hSemaphoreFORun := 0;
  end;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 07/11/2001
Modifié le ... : 07/11/2001
Description .. : Lance le clavier visuel de l'OS
Mots clefs ... : FO
*****************************************************************}

procedure FOLanceClavierVisuel;
var
  Cmd: string;
  WPtr: hWnd;
begin
  Cmd := 'OSK';
  WPtr := GetActiveWindow;
  ShellExecute(WPtr, nil, PChar(Cmd), nil, nil, SW_SHOWNORMAL);
end;

{==============================================================================================}
{======================================= DATES ================================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Extrait l'année d'une date
Mots clefs ... : FO
*****************************************************************}

function FODonneAn(Date: TDateTime; Taille: Integer = 4): string;
var Year, Month, Day: Word;
begin
  if Date = 0 then Date := NowH;
  DecodeDate(Date, Year, Month, Day);
  Result := IntToStr(Year);
  if Taille = 2 then Result := Copy(Result, Length(Result) - 1, 2);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 30/04/2004
Modifié le ... : 30/04/2004
Description .. : Convertit en format français (hh:mm:ss) une heure au format local
Mots clefs ... : FO
*****************************************************************}

function FOTimeToFrench(Heure: string): string;
var
  hh, mm, ss, ms: word;
  TT: TDateTime;
begin
  Result := '';
  if Trim(Heure) = '' then Exit;
  try
    TT := StrToTime(Heure);
  except
    on EConvertError do Exit;
  end;
  DecodeTime(TT, hh, mm, ss, ms);
  Result := Format('%2.2d:%2.2d:%2.2d', [hh, mm, ss]);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 30/04/2004
Modifié le ... : 30/04/2004
Description .. : Convertit en format local une heure au format français (hh:mm:ss)
Mots clefs ... : FO
*****************************************************************}

function FOFrenchToTime(Heure: string): string;
var
  hh, mm, ss, ms: word;
  TT: TDateTime;
begin
  Result := '';
  if Trim(Heure) = '' then Exit;
  hh := StrToInt(ReadTokenPipe(Heure, ':'));
  mm := StrToInt(ReadTokenPipe(Heure, ':'));
  ss := StrToInt(ReadTokenPipe(Heure, ':'));
  ms := 0;
  TT := EncodeTime(hh, mm, ss, ms);
  Result := TimeToStr(TT);
end;

{==============================================================================================}
{====================================== BOOLEENS ==============================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un champ de type boolean (X/-) d'une TOB est vrai
Mots clefs ... : FO
*****************************************************************}

function FOEstVrai(NomTob: TOB; NomChamp: string): Boolean;
begin
  if NomTob <> nil then Result := (NomTob.GetValue(NomChamp) = 'X')
  else Result := False;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un champ de type boolean (X/-) d'une TOB est
Suite ........ : faux
Mots clefs ... : FO
*****************************************************************}

function FOEstFaux(NomTob: TOB; NomChamp: string): Boolean;
begin
  if NomTob <> nil then Result := (NomTob.GetValue(NomChamp) = '-')
  else Result := False;
end;

{==============================================================================================}
{================================= GESTION DE LA CAISSE =======================================}
{==============================================================================================}
{*****************************************************************
 VH_GC.TOBPCaisse est une TOB qui contient le paramétrage de la caisse courante.

 Elle dispose des champs supplémentaires suivants :
  GPK_NUMZCAISSE   : numéro de la journée
  GPK_DATEOUV      : date de la journée
  GPK_CODEEURO     : code de la devise Euro
  OKDECF           : nombre de décimales de la devise fongible
  LASTNUMTIC       : numéro du dernier ticket de la caisse
  DEVISECAISSE     : devise de l'établissement
  REGIMETAXECAISSE : règime de taxe de la caisse
  QTEMAX           : quantité autorisée pour une ligne de ticket
  PRIXMAX          : prix maximum autorisé pour une ligne de ticket
  SO_GCFOLOGO      : nom du fichier contenant le logo de la caisse
  LASTVENDEUR      : code du vendeur du dernier ticket de la caisse

 Ses filles sont :
  'Types remise'    : liste des codes démarques et de leurs options
  'Confidentialité' : chargement en tableaux de booléens des évenements de confidentialité.

*****************************************************************}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : Indique si la caisse est mono-poste
Mots clefs ... : FO
*****************************************************************}

function FOCaisseIsMonoPoste: boolean;
begin
  Result := False;
  Result := GetSynRegKey(CAISSEMONOPOSTE, Result, True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : Traitement avant la phase de connexion et de séria
Mots clefs ... : FO
*****************************************************************}

{$IFDEF FOS5}
function FOAvantConnexion(var SortieHalley: boolean): boolean;
var
  Stg: string;
begin
  Result := True;
  if FOCaisseIsMonoPoste then
  begin
    if FOS5isRunning then
    begin
      Stg := TraduireMemoire('L''application ') + NomHalley
        + TraduireMemoire(' est déjà en cours d''exécution');
      PGIError(Stg);
      SortieHalley := True;
      Result := False;
      Application.Terminate;
    end else
      V_PGI.MultiUserLogin := True;
  end;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : Traitement après la phase de connexion
Mots clefs ... : FO
*****************************************************************}

{$IFDEF FOS5}
procedure FOApresConnexion;
var
  sSql: string;
  QQ: TQuery;
begin
  if FOCaisseIsMonoPoste then
  begin
    sSql := 'SELECT COUNT(*) FROM PARCAISSE WHERE GPK_FERME="-"';
    QQ := OpenSQL(sSql, True);
    if (QQ.EOF) or (QQ.Fields[0].AsInteger <> 1) then
    begin
      SaveSynRegKey(CAISSEMONOPOSTE, False, True) ;
      V_PGI.MultiUserLogin := False;
    end;
    Ferme(QQ);
  end;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 21/04/2004
Modifié le ... : 21/04/2004
Description .. : Traitement avant la phase de déconnexion
Mots clefs ... : FO
*****************************************************************}

{$IFDEF FOS5}
procedure FOAvantDeConnexion;
begin
  FOLibereMaterielCaisse;
  FOLibereVHGCCaisse;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : Traitement après la phase de déconnexion
Mots clefs ... : FO
*****************************************************************}

{$IFDEF FOS5}
procedure FOApresDeConnexion;
begin
  FOS5isStopping;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si une caisse de travail est sélectionnée
Mots clefs ... : FO
*****************************************************************}

function FOVerifieVHGCCaisse: Boolean;
begin
  Result := ((VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.NomTable = 'PARCAISSE') and
    (VH_GC.TOBPCaisse.FieldExists('GPK_CAISSE')) and (VH_GC.TOBPCaisse.GetValue('GPK_CAISSE') <> ''));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Mémorisation du paramètrage de la caisse choisie dans la
Suite ........ : variable globale PGI
Mots clefs ... : FO
*****************************************************************}

procedure FOChargeVHGCCaisse(Caisse: string);
var
  NumZ, NbDec, NumTic, iVal: integer;
  Stg, sVend: string;
  dVal: double;
  QQ: TQuery;
  TOBCais: TOB;
begin
  NumTic := 0;
  TOBCais := TOB.Create('PARCAISSE', nil, -1);
  TOBCais.InitValeurs;
  TOBCais.SelectDB('"' + Caisse + '"', nil);
  if VH_GC.TOBPCaisse <> nil then
  begin
    if VH_GC.TOBPCaisse.FieldExists('LASTNUMTIC') then
      NumTic := VH_GC.TOBPCaisse.GetValue('LASTNUMTIC');
    if VH_GC.TOBPCaisse.FieldExists('LASTVENDEUR') then
      sVend := VH_GC.TOBPCaisse.GetValue('LASTVENDEUR');
    for iVal := 0 to VH_GC.TOBPCaisse.Detail.Count - 1 do
    begin
      VH_GC.TOBPCaisse.Detail[0].ChangeParent(TOBCais, -1);
    end;
    VH_GC.TOBPCaisse.Free;
  end;
  VH_GC.TOBPCaisse := TOBCais;
  NumZ := FOGetNumZCaisse(Caisse);
  VH_GC.TOBPCaisse.AddChampSupValeur('GPK_NUMZCAISSE', NumZ);
  VH_GC.TOBPCaisse.AddChampSupValeur('GPK_DATEOUV', FOGetDateOuv(Caisse, NumZ));
  Stg := FOChargeDeviseEuro;
  VH_GC.TOBPCaisse.AddChampSupValeur('GPK_CODEEURO', Stg);
  if VH^.TenueEuro then
    NbDec := FODonneDeviseNbDec(V_PGI.DeviseFongible)
  else
    NbDec := FODonneDeviseNbDec(Stg);
  VH_GC.TOBPCaisse.AddChampSupValeur('OKDECF', NbDec);
  VH_GC.TOBPCaisse.AddChampSupValeur('LASTNUMTIC', NumTic);
  VH_GC.TOBPCaisse.AddChampSupValeur('LASTVENDEUR', sVend);
  // Recherche de la devise
  Stg := VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT');
  QQ := OpenSQL('SELECT ET_DEVISE FROM ETABLISS WHERE ET_ETABLISSEMENT="' + Stg + '"', True);
  Stg := '';
  if not QQ.EOF then
  begin
    Stg := QQ.FindField('ET_DEVISE').AsString;
    if not FOTabletteValueExist('TTDEVISE', Stg) then Stg := '';
  end;
  Ferme(QQ);
  if Stg = '' then Stg := V_PGI.DevisePivot;
  VH_GC.TOBPCaisse.AddChampSupValeur('DEVISECAISSE', Stg);
  // Recherche le symbole de la devise
  if Stg <> '' then
  begin
    QQ := OpenSQL('SELECT D_SYMBOLE FROM DEVISE WHERE D_DEVISE ="' + Stg + '"', True);
    if not QQ.EOF then
      Stg := QQ.FindField('D_SYMBOLE').AsString
      else
      Stg := '';
    Ferme(QQ);  
  end;
  VH_GC.TOBPCaisse.AddChampSupValeur('DEVISESYMBOLE', Stg);
  // Recheche du règime de taxe de la caisse
  Stg := VH_GC.TOBPCaisse.GetValue('GPK_TIERS');
  QQ := OpenSQL('SELECT T_REGIMETVA FROM TIERS WHERE T_TIERS="' + Stg + '"', True);
  if QQ.EOF then
    Stg := ''
  else
    Stg := QQ.FindField('T_REGIMETVA').AsString;
  Ferme(QQ);
  if Stg = '' then Stg := VH^.RegimeDefaut;
  VH_GC.TOBPCaisse.AddChampSupValeur('REGIMETAXECAISSE', Stg);
  // Force les variables VH^.EtablisDefaut VH_GC.GCDepotDefaut V_PGI.DateEntree
  FOForceVHGCCaisse;
  {$IFDEF FOS5}
  // Charge les caractèristiques des périphériques d'encaissement
  FOChargeMaterielCaisse(Caisse);
  // Charge les codes démarques
  FOChargeDemarque;
  FOChargeArtFiTiers; // JTR Lien Opération Caisse avec Tiers
  {$ENDIF}
  // Charge la quantité et le prix maximum autorisés pour une ligne de ticket
  iVal := 10000;
  iVal := FOGetFromRegistry(REGSAISIETIC, REGQTEMAXFFO, iVal);
  if iVal > 0 then
    dVal := iVal
  else
    dVal := 10000;
  VH_GC.TOBPCaisse.AddChampSupValeur('QTEMAX', dVal);
  iVal := 10000000;
  iVal := FOGetFromRegistry(REGSAISIETIC, REGPRIXMAXFFO, iVal);
  if iVal > 0 then
    dVal := iVal
  else
    dVal := 10000000;
  VH_GC.TOBPCaisse.AddChampSupValeur('PRIXMAX', dVal);
  // chargement des droits d'accès
  FOChargeConfidentialite(V_PGI.MenuCourant);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 02/03/2004
Modifié le ... : 02/03/2004
Description .. : Retourne la devise d'une caisse
Mots clefs ... :
*****************************************************************}

function FOGetDeviseEtab(Etablissement: string; Symbole: boolean): string;
begin
  Result := '';
  if Etablissement = '' then Exit;
  if Symbole then
  begin
    if VH_GC.TOBPCaisse.FieldExists('DEVISESYMBOLE') then
      Result := VH_GC.TOBPCaisse.GetString('DEVISESYMBOLE')
      else
      Result := '';
  end else
  begin
    if VH_GC.TOBPCaisse.FieldExists('DEVISECAISSE') then
      Result := VH_GC.TOBPCaisse.GetString('DEVISECAISSE')
      else
      Result := '';
  end;
  if Result = '' then
  begin
    if Symbole then
      Result := V_PGI.SymbolePivot
      else
      Result := V_PGI.DevisePivot;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 02/03/2004
Modifié le ... : 02/03/2004
Description .. : Retourne la devise d'une caisse
Mots clefs ... :
*****************************************************************}

function FOGetDeviseCaisse(Caisse: string; Symbole: boolean): string;
var
  sEtab: string;
begin
  Result := '';
  if Caisse = '' then Caisse := FOGetParamCaisse('GPK_CAISSE');
  if Caisse = FOGetParamCaisse('GPK_CAISSE') then
  begin
    if Symbole then
    begin
      if VH_GC.TOBPCaisse.FieldExists('DEVISESYMBOLE') then
        Result := VH_GC.TOBPCaisse.GetString('DEVISESYMBOLE')
        else
        Result := '';
    end else
    begin
      if VH_GC.TOBPCaisse.FieldExists('DEVISECAISSE') then
        Result := VH_GC.TOBPCaisse.GetString('DEVISECAISSE')
        else
        Result := '';
    end;
  end else
  begin
    sEtab := GetColonneSQL('PARCAISSE', 'GPK_ETABLISSEMENT', 'GPK_CAISSE="'+ Caisse +'"');
    Result := FOGetDeviseEtab(sEtab, Symbole);
  end;
  if Result = '' then
  begin
    if Symbole then
      Result := V_PGI.SymbolePivot
      else
      Result := V_PGI.DevisePivot;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Mémorisation du paramètrage de la caisse choisie dans la
Suite ........ : variable globale PGI
Mots clefs ... : FO
*****************************************************************}
{$IFDEF FOS5}

procedure FOChargeMaterielCaisse(Caisse: string);
var Materiel, Port, Param, Champs: string;
  Err: TMC_Err;
begin
  if V_MC = nil then Exit;
  V_MC.TableCaisse := 'PARCAISSE'; // Nom de la table où est stocké le paramétrage de la caisse
  V_MC.CodeCaisse := Caisse; // Code de la caisse
  //V_MC.Clear; // vide l'ancien paramétrage.
  // Chargement de l'afficheur
  if FOExisteAfficheur then
  begin
    FODonneParamAfficheur(Materiel, Port, Param);
    Champs := 'GPK_AFFTYPE;"";|GPK_AFFPORT;"";|GPK_AFFPARAM;"";|';
    if not V_MC.SetMateriel(mcAfficheur, Materiel, Port, Param, 'GCTYPEAFFICHEUR', Champs, Err) then
      FOMCErr(Err, Application.Title);
  end else
  begin
    if Trim(V_MC.GetDispositif(mcAfficheur)) <> '' then V_MC.DeleteMateriel(mcAfficheur, Err);
  end;
  // Chargement de l'imprimante de caisse
  if FOExisteLP then
  begin
    FODonneParamLP(Materiel, Port, Param);
    Champs := 'GPK_PRTTYPE;"";|GPK_PRTPORT;"";|GPK_PRTBAUDS;"";|GPK_PRTNBBIT;"";|'
      + 'GPK_PRTPARITE;"";|GPK_PRTSTOPBIT;"";|GPK_PRTCTRLFLUX;"";|GPK_REMPLIRCHQ;"-";|'
      + 'GPK_PRTMODEBIDI;"-";|';
    if not V_MC.SetMateriel(mcPrinter, Materiel, Port, Param, 'GCTYPEPRTCAISSE', Champs, Err) then
      FOMCErr(Err, Application.Title);
  end else
  begin
    if Trim(V_MC.GetDispositif(mcPrinter)) <> '' then V_MC.DeleteMateriel(mcPrinter, Err);
    Champs := 'GPK_AFFTYPE;"";|GPK_AFFPORT;"";|GPK_AFFPARAM;"";|';
    if not V_MC.SetMateriel(mcAfficheur, Materiel, Port, Param, 'GCTYPEAFFICHEUR', Champs, Err) then
      FOMCErr(Err, Application.Title);
  end;
  // Chargement du TPE
  if FOExisteTPE then
  begin
    FODonneParamTPE(Caisse, Materiel, Port, Param);
    Champs := 'GPK_TPETYPE;"";|GPK_TPEPORT;"";|GPK_TPEBAUDS;"";|';
    if not V_MC.SetMateriel(mcCarteB, Materiel, Port, Param, 'GCTYPETPE', Champs, Err) then
      FOMCErr(Err, Application.Title);
  end;
  // Chargement du tiroir
  if FOExisteTiroir then
  begin
    FODonneParamTiroir(Materiel, Port, Param);
    Champs := 'GPK_TIROIRTYPE;"";|GPK_TIROIRPORT;"";|GPK_TIROIRPARAM;"";|'
      + 'GPK_TIROIRPIN5;"-";|GPK_TIROIRESP;"-";|';
    if not V_MC.SetMateriel(mcTiroir, Materiel, Port, Param, 'GCTYPETIROIR', Champs, Err) then
      FOMCErr(Err, Application.Title);
  end else
  begin
    if Trim(V_MC.GetDispositif(mcTiroir)) <> '' then V_MC.DeleteMateriel(mcTiroir, Err);
  end;
  // Chargement du TPE
  if FOExisteTPE then
  begin
    FODonneParamTPE(Caisse, Materiel, Port, Param);
    Champs := 'GPK_TPETYPE;"";|GPK_TPEPORT;"";|GPK_TPEBAUDS;"";|';
    if not V_MC.SetMateriel(mcCarteB, Materiel, Port, Param, 'GCTYPETPE', Champs, Err) then
      FOMCErr(Err, Application.Title);
  end else
  begin
    if Trim(V_MC.GetDispositif(mcCarteB)) <> '' then V_MC.DeleteMateriel(mcCarteB, Err);
  end;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 21/04/2004
Modifié le ... : 21/04/2004
Description .. : Libération des périphériques de la caisse
Suite ........ : 
Suite ........ : Attention l'ordre de libération est important. Voir les
Suite ........ : préconisations d'IBM pour OPOS.
Mots clefs ... : FO
*****************************************************************}
{$IFDEF FOS5}

procedure FOLibereMaterielCaisse;
var 
  Err: TMC_Err;
begin
  if V_MC = nil then Exit;

  // Libération de l'afficheur
  if Trim(V_MC.GetDispositif(mcAfficheur)) <> '' then V_MC.DeleteMateriel(mcAfficheur, Err);
  // Libération de l'imprimante de caisse
  if Trim(V_MC.GetDispositif(mcPrinter)) <> '' then V_MC.DeleteMateriel(mcPrinter, Err);
  // Libération du tiroir
  if Trim(V_MC.GetDispositif(mcTiroir)) <> '' then V_MC.DeleteMateriel(mcTiroir, Err);
  // Libération du TPE
  if Trim(V_MC.GetDispositif(mcCarteB)) <> '' then V_MC.DeleteMateriel(mcCarteB, Err);

  V_MC.Clear;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Recharge le paramètrage de la caisse choisie dans la
Suite ........ : variable globale PGI
Mots clefs ... : FO
*****************************************************************}

procedure FOReChargeVHGCCaisse;
begin
  if VH_GC.TOBPCaisse <> nil then FOChargeVHGCCaisse(VH_GC.TOBPCaisse.GetValue('GPK_CAISSE'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Met à jour la variable globale PGI à partir du paramètrage de
Suite ........ : la caisse choisie
Mots clefs ... : FO
*****************************************************************}

procedure FOForceVHGCCaisse;
var sBuf: string;
begin
  if VH_GC.TOBPCaisse <> nil then
  begin
    sBuf := VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT');
    if sBuf <> '' then VH^.EtablisDefaut := sBuf;
    sBuf := VH_GC.TOBPCaisse.GetValue('GPK_DEPOT');
    if sBuf <> '' then VH_GC.GCDepotDefaut := sBuf;
    V_PGI.DateEntree := VH_GC.TOBPCaisse.GetValue('GPK_DATEOUV');
  end;
  //if V_PGI.DevisePivot = CODEEURO then V_PGI.SymbolePivot := SIGLEEURO ;
  //if V_PGI.DeviseFongible = CODEEURO then VH^.SymboleFongible := SIGLEEURO ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 13/03/2003
Description .. : Retourne l'état de la journée en cours d'une caisse
Suite ........ : (Ouverte, Fermée, ...)
Mots clefs ... : FO
*****************************************************************}

function FOEtatCaisse(Caisse: string; Recharge: boolean = True): string;
var QQ: TQuery;
begin
  Result := '';
  // Lecture systèmatique des caractéristiques de la caisse
  if Recharge then FOChargeVHGCCaisse(Caisse);
  QQ := OpenSQL('SELECT GJC_ETAT FROM JOURSCAISSE WHERE GJC_CAISSE="' + Caisse + '"'
    + ' AND GJC_NUMZCAISSE = (SELECT MAX(GJC_NUMZCAISSE) FROM JOURSCAISSE'
    + ' WHERE GJC_CAISSE="' + Caisse + '")', TRUE);
  if not QQ.EOF then Result := QQ.FindField('GJC_ETAT').AsString;
  Ferme(QQ);
  if Result = '' then Result := ETATJOURCAISSE[1]; // NON Ouverte
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001                                                              m
Modifié le ... : 23/07/2001
Description .. : Retourne l'état d'une journée d'une caisse (Ouverte,
Suite ........ : Fermée, ...)
Mots clefs ... : FO
*****************************************************************}

function FOEtatJournee(Caisse: string; NumZ: Integer): string;
var FOEtat: TQuery;
begin
  Result := '';
  FOEtat := OpenSQL('SELECT GJC_ETAT FROM JOURSCAISSE WHERE GJC_CAISSE="' + Caisse + '" AND GJC_NUMZCAISSE=' + IntToStr(NumZ), TRUE);
  if not FOEtat.EOF then Result := FOEtat.FindField('GJC_ETAT').AsString;
  Ferme(FOEtat);
  if Result = '' then Result := ETATJOURCAISSE[1]; // NON Ouverte
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne le code de la caisse de travail
Mots clefs ... : FO
*****************************************************************}

function FOCaisseCourante: string;
begin
  Result := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.NomTable = 'PARCAISSE')
    and (VH_GC.TOBPCaisse.FieldExists('GPK_CAISSE')) then
  begin
    Result := VH_GC.TOBPCaisse.GetValue('GPK_CAISSE');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si la caisse courante est disponible
Mots clefs ... : FO
*****************************************************************}

function FOCaisseDisponible: Boolean;
begin
  Result := False;
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_FERME')) then
  begin
    Result := (VH_GC.TOBPCaisse.GetValue('GPK_FERME') = '-');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Mémorisation du paramètrage de la caisse choisie dans la
Suite ........ : variable globale PGI
Mots clefs ... : FO
*****************************************************************}

function FOCaisseEtab(Etab: string): string;
var QQ: TQuery;
  Sql: string;
begin
  Result := '';
  Sql := 'Select GPK_CAISSE from PARCAISSE where GPK_ETABLISSEMENT="' + Etab + '"';
  QQ := OpenSQL(Sql, True);
  if not QQ.EOF then Result := QQ.Fields[0].AsString;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne numéro d'ouverture de caisse attribué selon la
Suite ........ : borne choisie (MIN ou MAX)
Mots clefs ... : FO
*****************************************************************}

function FOGetNumZCaisse(Caisse: string; Borne: string = 'MAX'): Integer;
var FONumZ: TQuery;
begin
  Result := 0;
  FONumZ := OpenSQL('SELECT ' + Borne + '(GJC_NUMZCAISSE) FROM JOURSCAISSE WHERE GJC_CAISSE = "' + Caisse + '"', True);
  if not FONumZ.EOF then Result := FONumZ.Fields[0].AsInteger;
  Ferme(FONumZ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/01/2004
Description .. : Retourne la date d'ouverture d'une journée d'une caisse
Mots clefs ... : FO
*****************************************************************}

procedure FOGetInfoNumZ(Caisse: string; NumZ: Integer; var DateOuv: TDateTime; var ExceptionTaxe: string);
var
  QQ: TQuery;
  Sql: string;
begin
  DateOuv := Date;
  ExceptionTaxe := '';
  Sql := 'SELECT GJC_DATEOUV,GJC_EXCEPTIONTAXE FROM JOURSCAISSE'
    + ' WHERE GJC_CAISSE="' + Caisse + '" AND GJC_NUMZCAISSE=' + IntToStr(NumZ);
  QQ := OpenSQL(Sql, True);
  if not QQ.EOF then
  begin
    DateOuv := QQ.FindField('GJC_DATEOUV').AsDateTime;
    ExceptionTaxe := QQ.FindField('GJC_EXCEPTIONTAXE').AsString;
  end;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne la date d'ouverture d'une journée d'une caisse
Mots clefs ... : FO
*****************************************************************}

function FOGetDateOuv(Caisse: string; NumZ: Integer): TDateTime;
var QQ: TQuery;
  Sql: string;
begin
  Result := Date;
  Sql := 'select GJC_DATEOUV from JOURSCAISSE where GJC_CAISSE="' + Caisse + '" and GJC_NUMZCAISSE=' + IntToStr(NumZ);
  QQ := OpenSQL(Sql, True);
  if not QQ.EOF then Result := QQ.Fields[0].AsDateTime;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 21/11/2001
Modifié le ... : 21/11/2001
Description .. : Constitue une sélection sur la date des pièces à partir d'une
Suite ........ : borne de n° de clôture.
Mots clefs ... : FO
*****************************************************************}

function FOSelectNumZDate(Prefixe: string; DeNumZ, ANumZ: integer; Caisse: string = ''): string;
var DeDateOuv, ADateOuv: TDateTime;
begin
  FOBorneNumZDate(DeNumZ, ANumZ, DeDateOuv, ADateOuv, Caisse);
  Result := Prefixe + '_DATEPIECE>="' + UsDateTime(DeDateOuv) + '"'
    + ' AND ' + Prefixe + '_DATEPIECE<="' + UsDateTime(ADateOuv) + '"';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 21/11/2001
Modifié le ... : 21/11/2001
Description .. : Retourne les dates d'ouverture à partir d'une borne de n° de
Suite ........ : clôture.
Mots clefs ... : FO
*****************************************************************}

procedure FOBorneNumZDate(DeNumZ, ANumZ: integer; var DeDateOuv, ADateOuv: TDateTime; Caisse: string = '');
var sSql: string;
  QQ: TQuery;
begin
  DeDateOuv := EncodeDate(1900, 1, 1);
  ADateOuv := EncodeDate(1900, 1, 1);
  if Caisse = '' then Caisse := FOCaisseCourante;
  sSql := 'SELECT GJC_DATEOUV FROM JOURSCAISSE WHERE GJC_CAISSE="' + Caisse + '"'
    + ' AND GJC_NUMZCAISSE IN (' + IntToStr(DeNumZ) + ',' + IntToStr(ANumZ) + ')'
    + ' ORDER BY GJC_NUMZCAISSE';
  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then
  begin
    DeDateOuv := QQ.FindField('GJC_DATEOUV').AsDateTime;
    QQ.Next;
    if not QQ.EOF then
      ADateOuv := QQ.FindField('GJC_DATEOUV').AsDateTime
    else
      ADateOuv := DeDateOuv;
  end;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Verrouille une caisse (marque comme étant en cours
Suite ........ : d'utilisation).
Mots clefs ... : FO
*****************************************************************}

function FOLockCaisse(Caisse: string): boolean;
var Stg: string;
begin
  //Stg := 'update PARCAISSE set GPK_INUSE="X", GPK_DATECONNEXION="'+UsTime(NowH)+'" where GPK_CAISSE="'+Caisse+'"' ;
  Stg := 'UPDATE PARCAISSE SET GPK_INUSE="X",GPK_DATECONNEXION="' + UsTime(NowH) + '"'
    + ' WHERE GPK_CAISSE="' + Caisse + '" AND GPK_INUSE="-"';
  Result := (ExecuteSQL(Stg) <> 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 25/07/2002
Description .. : Libère une caisse (supprime la marque comme étant en
Suite ........ : cours d'utilisation).
Mots clefs ... : FO
*****************************************************************}

function FOLibereCaisse(Caisse: string): Boolean;
begin
  if Caisse = FOCaisseCourante then
  begin
    Result := FOLibereVHGCCaisse;
    GCTripoteStatus;
  end else
  begin
    Result := FODemarqueCaisse(Caisse);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 12/03/2003
Description .. : Libère la caisse courante (supprime la marque comme étant
Suite ........ : en cours d'utilisation).
Mots clefs ... : FO
*****************************************************************}

function FOLibereVHGCCaisse: boolean;
begin
  Result := True;
  if VH_GC.TOBPCaisse <> nil then
  begin
    if FOVerifieVHGCCaisse then
      Result := FODemarqueCaisse(VH_GC.TOBPCaisse.GetValue('GPK_CAISSE'));
    VH_GC.TOBPCaisse.Free;
    VH_GC.TOBPCaisse := TOB.Create('PARCAISSE', nil, -1);
    VH_GC.TOBPCaisse.InitValeurs;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 25/07/2002
Description .. : Démarque une caisse (supprime la marque comme étant en
Suite ........ : cours d'utilisation).
Mots clefs ... : FO
*****************************************************************}

function FODemarqueCaisse(Caisse: string): Boolean;
var Stg: string;
begin
  Stg := 'update PARCAISSE set GPK_INUSE="-" where GPK_CAISSE="' + Caisse + '"';
  Result := (ExecuteSQL(Stg) <> 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Indique si une seule caisse est définie dans la table des
Suite ........ : caisses
Mots clefs ... : FO
*****************************************************************}

function FOMonoCaisse(Ouverte: Boolean = False; Libre: Boolean = False): Boolean;
var QQ: TQuery;
  sSelect: string;
begin
  sSelect := 'SELECT COUNT(*) FROM PARCAISSE';
  if (Ouverte) or (Libre) then sSelect := sSelect + ' WHERE ';
  if Ouverte then sSelect := sSelect + 'GPK_FERME="-"';
  if (Ouverte) and (Libre) then sSelect := sSelect + ' AND ';
  if Libre then sSelect := sSelect + 'GPK_INUSE="-"';
  QQ := OpenSQL(sSelect, True);
  Result := ((QQ.EOF) or (QQ.Fields[0].AsInteger = 1));
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/09/2002
Modifié le ... : 27/09/2002
Description .. : Retoune une option de la connexion à la caisse :
Suite ........ :   - connexion systèmatiquement lorsqu'on lance une ligne de
Suite ........ : menu.
Suite ........ :   - connexion pour toutes les lignes de menu
Suite ........ :   - vérification si la caisse choisie est déjà utilisée
Mots clefs ... : FO
*****************************************************************}

function FOOptionConnexionCaisse(Option: TTypeOptConnCaisseFO): Boolean;
var Stg: string;
  bDef: Boolean;
begin
  Stg := '';
  Result := False;
  bDef := False;
  case Option of
    cfoChaqueFois:
      begin
        Stg := REGCHAQUEFOIS;
        bDef := False;
      end;
    cfoToutesLignes:
      begin
        Stg := REGTOUTESLIGNES;
        bDef := False;
      end;
    cfoVerifVerrou:
      begin
        Stg := REGVERIFVERROU;
        bDef := True;
      end;
  end;
  if Stg <> '' then Result := FOGetFromRegistry(REGCONNEXCAIS, Stg, bDef);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 21/01/2002
Description .. : Retourne un paramètre de la caisse courante
Mots clefs ... : FO
*****************************************************************}

function FOGetParamCaisse(NomChamp: string; TypeChamp: string = ''): variant;
begin
  Result := FOChampValeurVide(NomChamp, TypeChamp);
  if NomChamp = '' then Exit;
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists(NomChamp)) then
    Result := VH_GC.TOBPCaisse.GetValue(NomChamp);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 21/01/2002
Description .. : Force la valeur d'un paramètre de la caisse courante
Mots clefs ... : FO
*****************************************************************}

procedure FOPutParamCaisse(NomChamp: string; Valeur: Variant);
begin
  if NomChamp = '' then Exit;
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists(NomChamp)) then
  begin
    VH_GC.TOBPCaisse.PutValue(NomChamp, Valeur);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne le nom du fichier contenant le logo de la caisse
Suite ........ : courante
Mots clefs ... : FO
*****************************************************************}

function FOLogoCaisse: string;
begin
  Result := '';
  if VH_GC.TOBPCaisse <> nil then
  begin
    if not (VH_GC.TOBPCaisse.FieldExists('SO_GCFOLOGO')) then VH_GC.TOBPCaisse.AddChampSupValeur('SO_GCFOLOGO', Trim(GetParamSoc('SO_GCFOLOGO')));
    Result := VH_GC.TOBPCaisse.GetValue('SO_GCFOLOGO');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... :   /  /
Description .. : Retourne le nombre d'exemplaires d'un ticket ou d'un bon
Suite ........ : (avoir, arrhes, transfert, ...) à imprimer
Mots clefs ... : FO
*****************************************************************}

function FONbExemplaire(TypeDoc: string = 'BON'): Integer;
var NomChamp: string;
begin
  Result := 0;
  if TypeDoc = 'TIC' then NomChamp := 'GPK_NBEXEMPTIC' else NomChamp := 'GPK_NBEXEMPBON';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists(NomChamp)) then
  begin
    Result := VH_GC.TOBPCaisse.GetValue(NomChamp);
  end;
  if (Result < 1) or (Result > 9) then Result := 1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/07/2003
Modifié le ... : 03/07/2003
Description .. : retourne la naturte de pièce des tickets
Mots clefs ... : FO
*****************************************************************}

function FOGetNatureTicket (EnAttente, PourSql: boolean): string;
begin
  {$IFDEF CHR}
  Result := VH_CHR.PieceNotehotel;
  if PourSql then Result := '"'+ Result +'"';
  {$ELSE}
  if EnAttente then
    Result := 'FFA'
  else
    Result := 'FFO';
  if PourSql then Result := '"'+ Result +'"';
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 21/07/2003
Modifié le ... : 21/07/2003
Description .. : Filtre les caisses de l'établissement défini pour l'utilisateur
Mots clefs ... : 
*****************************************************************}

procedure FOPositionneCaisseUser (cbCaisse, cbEtab: THValComboBox; ForceEtabDefaut: boolean = False);
var
  Etab, Stg, Code : String ;
begin
  Etab := EtabForce;
  if (Etab = '') and (ForceEtabDefaut) then Etab := VH^.EtablisDefaut;
  if Etab = '' then Exit ;
  if (cbCaisse <> nil) and (cbCaisse.DataType = 'GCPCAISSE') then
  begin
    Code := cbCaisse.Value;
    Stg := cbCaisse.Plus;
    if Stg <> '' then Stg := Stg + ' AND ';
    //Stg := Stg + ' AND GPK_ETABLISSEMENT="'+ Etab +'"';
    Stg := Stg + 'GPK_ETABLISSEMENT="'+ Etab +'"';
    cbCaisse.Plus := Stg;
    cbCaisse.Value := Code;
  end;
  if cbEtab <> nil then
  begin
    if cbEtab.Visible then
      PositionneEtabUser(cbEtab)
    else
      cbEtab.Value := Etab;

    if (cbEtab.Value = '') and (ForceEtabDefaut) then cbEtab.Value := Etab;
  end;
end;

{==============================================================================================}
{============================ JOURNÉES DE VENTES DES CAISSES ==================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2002
Modifié le ... : 28/10/2002
Description .. : Incrémente les compteurs de la journée de vente courante
Suite ........ : pour la caisse courante.
Mots clefs ... : FO
*****************************************************************}

function FOIncrJoursCaisse(TypeCpt: TTypeCptJourCaisseFO; MsgEvent: string = ''; Increment: Integer = 1): Boolean;
var sNomChamp, sSql, sCaisse, sLib: string;
  iNumZ, iNumEvt: Integer;
  TOBJnal: TOB;
  QQ: TQuery;
begin
  Result := False;
  if Increment = 0 then Exit;
  sCaisse := FOGetParamCaisse('GPK_CAISSE');
  if sCaisse = '' then Exit;
  iNumZ := FOGetParamCaisse('GPK_NUMZCAISSE');
  if iNumZ = 0 then Exit;
  // Choix du compteur
  case TypeCpt of
    jfoNbTicAbandon:
      begin
        sNomChamp := 'GJC_NBTICABANDON';
        sLib := 'Abandon du ticket';
      end;
    jfoNbTicAnnul:
      begin
        sNomChamp := 'GJC_NBTICANNUL';
        sLib := 'Annulation du ticket';
      end;
    jfoNbTicMisatt:
      begin
        sNomChamp := 'GJC_NBTICMISATT';
        sLib := 'Mise en attente du ticket';
      end;
    jfoNbTicattRepri:
      begin
        sNomChamp := 'GJC_NBTICATTREPRI';
        sLib := 'Ticket en attente repris';
      end;
    jfoNbTicattSup:
      begin
        sNomChamp := 'GJC_NBTICATTSUP';
        sLib := 'Tickets en attente détruits';
      end;
    jfoNbLigAnnul:
      begin
        sNomChamp := 'GJC_NBLIGANNUL';
        sLib := 'Ligne annulée';
      end;
    jfoNbModifMdp:
      begin
        sNomChamp := 'GJC_NBMODIFMDP';
        sLib := 'Modification du règlement';
      end;
    jfoNbOuvTiroir:
      begin
        sNomChamp := 'GJC_NBOUVTIROIR';
        sLib := 'Ouverture du tiroir';
      end;
    jfoNbRatcli:
      begin
        sNomChamp := 'GJC_NBRATCLI';
        sLib := 'Rattachement à un client';
      end;
  else Exit;
  end;
  // Mise à jour du champ
  sSql := 'UPDATE JOURSCAISSE SET ' + sNomChamp + '=' + sNomChamp + '+' + IntToStr(Increment)
    + ' WHERE GJC_CAISSE="' + sCaisse + '" AND GJC_NUMZCAISSE="' + IntToStr(iNumZ) + '"';
  Result := (ExecuteSQL(sSql) <> 0);
  // Journal des événements
  if V_PGI.LogUser then
  begin
    TOBJnal := TOB.Create('JNALEVENT', nil, -1);
    TOBJnal.InitValeurs;
    TOBJnal.PutValue('GEV_TYPEEVENT', 'FFO');
    TOBJnal.PutValue('GEV_DATEEVENT', Date);
    TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
    TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
    TOBJnal.PutValue('GEV_LIBELLE', sLib);
    if MsgEvent <> '' then TOBJnal.PutValue('GEV_BLOCNOTE', MsgEvent);
    // Recheche du n° de l'événement
    QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True);
    if QQ.EOF then iNumEvt := 0 else iNumEvt := QQ.Fields[0].AsInteger;
    Ferme(QQ);
    Inc(iNumEvt);
    TOBJnal.PutValue('GEV_NUMEVENT', iNumEvt);
    TOBJnal.InsertDB(nil);
    TOBJnal.Free;
  end;
end;

{==============================================================================================}
{====================================== REMISES ===============================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne le code de l'arrondi à appliquer sur le montant de
Suite ........ : la remise
Mots clefs ... : FO
*****************************************************************}

function FODonneArrondiRemise: string;
begin
  Result := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_REMARRONDI')) then
  begin
    Result := VH_GC.TOBPCaisse.GetValue('GPK_REMARRONDI');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Arrondi le montant de la remise en fonction du paramétrage
Suite ........ : de la caisse
Mots clefs ... : FO
*****************************************************************}

function FOArrondiRemise(MntRemise, MntBrut: Double): Double;
var CodeArr: string;
begin
  Result := MntRemise;
  CodeArr := FODonneArrondiRemise;
  if CodeArr <> '' then
  begin
    Result := ArrondirPrix(CodeArr, MntRemise);
    // on ne peut excéder le montant brut par application de l'arrondi
    if Abs(Result) > Abs(MntBrut) then result := MntBrut;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 06/01/2003
Description .. : Vérifie si une remise dépasse le maximun autorisé dans le
Suite ........ : paramétrage de la caisse
Suite ........ :
Suite ........ : Valeur retournée :
Suite ........ :  0 = ok
Suite ........ :  1 = maximum autorisé pour la caisse dépassé
Suite ........ :  2 = maximum autorisé pour le type de remise dépassé
Suite ........ :  3 = Majoration du prix interdite
Mots clefs ... : FO
*****************************************************************}

function FOVerifMaxRemise(TauxRemise: double; TypeRemise: string; TypeUniquement: boolean; var MaxRem: integer; var CodeDemValide: string): integer;
const Mess1: string = 'Vous ne pouvez saisir une remise supérieure à ';
  Mess2: string = ' %';
  Mess3: string = ' pour le motif de démarque ';
  Mess4: string = 'Vous ne pouvez pas majorer le montant.';
var RemPourMax, RemPourMax2, RemPourMax3: integer;
  TOBDem, TOBL: TOB;
  Stg, Msg: string;
begin
  Result := 0;
  // Majoration du prix
  if TauxRemise < 0 then
  begin
    if (VH_GC.TOBPCaisse.GetValue('GPK_MAJORPRX') = '-') then
    begin
      PGIBox(Mess4);
      Result := 3;
    end else
      if not FOJaiLeDroit(34, True, True, Mess4) then
    begin
      Result := 3;
    end;
    Exit;
  end;
  // Vérification du maximum autorisé pour le type de remise
  if (TypeRemise <> '') and (FOGetParamCaisse('GPK_DEMSAISIE') = 'X') then
  begin
    TOBDem := FOGetPTobDemarque;
    if TOBDem <> nil then
    begin
      TOBL := TOBDem.FindFirst(['GTR_TYPEREMISE'], [TypeRemise], False);
      Stg := FODecodeLibreDemarque(TOBL, '3');
      if Stg <> '' then
      begin
        MaxRem := StrToInt(Stg);
        if (MaxRem > 0) and (TauxRemise > MaxRem) and (TypeRemise <> CodeDemValide) then
        begin
          Msg := TraduireMemoire(Mess1) + IntToStr(MaxRem) + Mess2
            + TraduireMemoire(Mess3) + TypeRemise;
          if FOJaiLeDroit(33, True, True, Msg) then
          begin
            if CodeDemValide = '' then CodeDemValide := TypeRemise;
          end else
          begin
            Result := 2;
            Exit;
          end;
        end;
      end;
    end;
  end;
  if TypeUniquement then Exit;
  // Vérification du maximun autorisé pour la caisse
  if not VH_GC.TOBPCaisse.FieldExists('GPK_REMPOURMAX') then
  begin
    Result := 0;
    Exit;
  end;
  Result := 1;
  MaxRem := 0;
  RemPourMax := VarAsType(VH_GC.TOBPCaisse.GetValue('GPK_REMPOURMAX'), VarInteger);
  // Pourcentage niveau 1
  if TauxRemise <= RemPourMax then
  begin
    Result := 0;
    MaxRem := RemPourMax;
    Exit;
  end;
  MaxRem := RemPourMax;
  // Pourcentage niveau 2
  RemPourMax2 := VarAsType(VH_GC.TOBPCaisse.GetValue('GPK_REMPOURMAX2'), VarInteger);
  if TauxRemise <= RemPourMax2 then
  begin
    //if FOSaisiePwd (['GPK_PWDREM2', 'GPK_PWDREM3']) then
    Msg := TraduireMemoire(Mess1) + IntToStr(MaxRem) + Mess2;
    if FOJaiLeDroit(30, True, True, Msg) then
    begin
      Result := 0;
      if RemPourMax2 > MaxRem then MaxRem := RemPourMax2;
    end;
    Exit;
  end;
  if RemPourMax2 > MaxRem then MaxRem := RemPourMax2;
  // Pourcentage niveau 3
  RemPourMax3 := VarAsType(VH_GC.TOBPCaisse.GetValue('GPK_REMPOURMAX3'), VarInteger);
  if TauxRemise <= RemPourMax3 then
  begin
    //if FOSaisiePwd (['GPK_PWDREM3']) then
    ///Msg := TraduireMemoire (Mess1) + IntToStr (MaxRem) + Mess2 ;
    ///if (FOJaiLeDroit (31, True, True, Msg)) and (FOJaiLeDroit (30, True, True, Msg)) then
    Msg := TraduireMemoire(Mess1) + IntToStr(RemPourMax2) + Mess2;
    if FOJaiLeDroit(31, True, True, Msg) then
    begin
      Msg := TraduireMemoire(Mess1) + IntToStr(RemPourMax) + Mess2;
      if FOJaiLeDroit(30, True, True, Msg) then
      begin
        Result := 0;
        if RemPourMax3 > MaxRem then MaxRem := RemPourMax3;
      end;
    end;
    Exit;
  end;
  if RemPourMax3 > MaxRem then MaxRem := RemPourMax3;
  // le pourcentage de remise ne peut dépasser les 100%
  if TauxRemise <= 100 then
  begin
    ///Msg := TraduireMemoire (Mess1) + IntToStr (RemPourMax3) + Mess2 ;
    ///if (FOJaiLeDroit (32, True, True, Msg)) and (FOJaiLeDroit (31, True, True, Msg)) and
    ///   (FOJaiLeDroit (30, True, True, Msg)) then
    Msg := TraduireMemoire(Mess1) + IntToStr(RemPourMax3) + Mess2;
    if FOJaiLeDroit(32, True, True, Msg) then
    begin
      Msg := TraduireMemoire(Mess1) + IntToStr(RemPourMax2) + Mess2;
      if FOJaiLeDroit(31, True, True, Msg) then
      begin
        Msg := TraduireMemoire(Mess1) + IntToStr(RemPourMax) + Mess2;
        if FOJaiLeDroit(30, True, True, Msg) then

        begin
          Result := 0;
          MaxRem := 100;
        end;
      end;
    end;
  end else
  begin
    Msg := TraduireMemoire(Mess1) + IntToStr(100) + Mess2;
    PGIError(Msg);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Demande la saisie d'un mot de passe par rapport à une liste
Suite ........ : de mots de passe possibles.
Mots clefs ... : FO
*****************************************************************}

function FOSaisiePwd(NomChamp: array of string): Boolean;
var PwdLst, PwdSaisi: string;
  Ind: Integer;
begin
  PwdLst := '';
  for Ind := Low(NomChamp) to High(NomChamp) do
  begin
    if PwdLst <> '' then PwdLst := PwdLst + ';';
    PwdLst := PwdLst + DeCryptageSt(VH_GC.TOBPCaisse.GetValue(NomChamp[Ind]));
  end;
  PwdSaisi := AGLLanceFiche('MFO', 'MOTPASSE', '', '', PwdLst);
  Result := (FOStrCmp(PwdSaisi, PwdLst));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/02/2003
Modifié le ... : 20/02/2003
Description .. : Compare 2 montants de manière à éviter les pouième (e-18)
Mots clefs ... : FO
*****************************************************************}

function FOCompareMontant(Montant1, Montant2: double; Operateur: string = '='): boolean;
var Mnt1, Mnt2: double;
begin
  Mnt1 := Arrondi((Montant1 * 1000000), 0);
  Mnt2 := Arrondi((Montant2 * 1000000), 0);
  if Operateur = '<' then
    Result := (Mnt1 < Mnt2)
  else if Operateur = '<=' then
    Result := (Mnt1 <= Mnt2)
  else if Operateur = '>' then
    Result := (Mnt1 > Mnt2)
  else if Operateur = '>=' then
    Result := (Mnt1 >= Mnt2)
  else
    Result := (Mnt1 = Mnt2);
end;

{==============================================================================================}
{============================= BARRE DE STATUS et BOÎTES DE MESSAGE ===========================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Ajoute le code et le libellé de la caisse courante dans la
Suite ........ : barre de status du menu principal.
Mots clefs ... : FO
*****************************************************************}

procedure FOChgStatus;
var Stg: string;
begin
  if FOCaisseCourante <> '' then
  begin
    Stg := GetDefStatus + '               ' + TraduireMemoire('Caisse : ') + FOCaisseCourante
      + ' - ' + VH_GC.TOBPCaisse.GetValue('GPK_LIBELLE');
    ChgStatus(Stg);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/03/2003
Modifié le ... : 11/03/2003
Description .. : Retourne le message ou le titre affiché par MsgBox.Execute
Mots clefs ... : FO
*****************************************************************}

function FOMsgBoxExec(MsgBox: THMsgBox; NoMsg: integer; Av, Ap: string; DonneTitre: boolean = False): string;
var Stg, Msg, Titre: string;
begin
  Stg := MsgBox.Mess[NoMsg];
  ReadTokenSt(Stg);
  Titre := ReadTokenSt(Stg);
  if (Uppercase(Titre) = '?CAPTION?') or (Trim(Titre) = '') then
  begin
    if Av = '' then
      Titre := 'Attention'
    else
      Titre := Av;
    Av := '';
  end;
  Titre := FindEtReplace(Titre, '%%', Av, True);
  Titre := FindEtReplace(Titre, '$$', Ap, True);
  Msg := ReadTokenSt(Stg);
  if Pos('%%', Msg) = 0 then Msg := Av + ' ' + Msg;
  if Pos('$$', Msg) = 0 then Msg := Msg + ' ' + Ap;
  Msg := FindEtReplace(Msg, '%%', Av, True);
  Msg := FindEtReplace(Msg, '$$', Ap, True);
  if DonneTitre then
    Result := Titre
  else
    Result := Msg;
end;

{==============================================================================================}
{======================================= COMBO ================================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un code existe dans la liste des valeurs d'un
Suite ........ : combo.
Mots clefs ... : FO
*****************************************************************}

function FOComboValueExist(Combo: THValComboBox; Value: string): Boolean;
begin
  Result := (Combo.Values.IndexOf(Value) >= 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Supprime un code dans la liste des valeus d'un combo.
Mots clefs ... : FO
*****************************************************************}

function FOComboDeleteValue(Combo: THValComboBox; Value: string): Boolean;
var Ind: Integer;
begin
  Result := False;
  Ind := Combo.Values.IndexOf(Value);
  if Ind >= 0 then
  begin
    Combo.Values.Delete(Ind);
    Combo.Items.Delete(Ind);
    Result := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un code existe dans une tablette et retourne un
Suite ........ : libellé vide (ou 'Error' en mode SAV) si le code n'existe pas
Mots clefs ... : FO
*****************************************************************}

function FOTabletteValueExist(Tablette: string; Value: string; Plus: string = ''): Boolean;
var Libelle: string;
begin
  Result := False;
  if Value <> '' then
  begin
    if V_PGI.SAV then Libelle := 'Error' else Libelle := '';
    if RechDom(Tablette, Value, False, Plus) <> Libelle then Result := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Fabrique une clause IN pour un SELECT à partir des
Suite ........ : codes sélectionnés d'un multi-combo
Mots clefs ... : FO
*****************************************************************}

function FOFabriqueSQLIN(Liste, NomChamp: string; WithAnd, NotIn, ChampNumeric: Boolean): string;
var Stg, Operateur, Separateur: string;
begin
  Result := '';
  if (Liste = '') or (NomChamp = '') then Exit;
  Stg := Liste;
  Stg := Trim(Stg);
  // Cas du <<Tous>>
  if (Pos('<<', Stg) = 1) and (Pos('>>', Stg) = Length(Stg) - 1) then Stg := '';
  if Stg <> '' then
  begin
    if Stg[Length(Stg)] = ';' then Delete(Stg, Length(Stg), 1);
    if WithAnd then Result := ' and ';
    //   Result := Result + NomChamp + ' in ("'+FindEtReplace(Stg, ';', '","', True)+'") ' ;
    if ChampNumeric then Separateur := '' else Separateur := '"';
    Stg := Separateur + FindEtReplace(Stg, ';', Separateur + ',' + Separateur, True) + Separateur;
    if NotIn then Operateur := 'not in' else Operateur := 'in';
    Result := Result + NomChamp + ' ' + Operateur + ' (' + Stg + ') '
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Fabrique la liste des modes de paiement à prendre en
Suite ........ : compte dans le fond de caisse ou le contrôle de caisse
Suite ........ : selon le paramétrage de la caisse
Mots clefs ... : FO
*****************************************************************}

function FOFabriqueListeMDP(NomListe, NomChamp: string): string;
begin
  Result := '';
  if (NomListe = '') or (NomChamp = '') then Exit;
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists(NomListe)) then
  begin
    Result := FOFabriqueSQLIN(VH_GC.TOBPCaisse.GetValue(NomListe), NomChamp, True, False, False);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/11/2001
Modifié le ... : 27/11/2001
Description .. : Fabrique la liste des types d'articles autorisés pour une
Suite ........ : nature de pièce.
Mots clefs ... : FO
*****************************************************************}

function FOFabriqueListeTypArt(NaturePieceG, NomChamp: string): string;
var Stg: string;
begin
  Result := '';
  if (NaturePieceG = '') or (NomChamp = '') then Exit;
  Stg := GetInfoParPiece(NaturePieceG, 'GPP_TYPEARTICLE');
  Result := FOFabriqueSQLIN(Stg, NomChamp, False, False, False);
end;

{==============================================================================================}
{================================== CHAMPS ET TABLES ==========================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne l'indicateur 'Controle' d'un champ
Mots clefs ... : FO
*****************************************************************}

function FOControlChamp(NomChamp: string): string;
var PrfxTable: string;
  NoTable, Ind: Integer;
begin
  Result := '';
  if NomChamp = '' then Exit;
  PrfxTable := ExtractPrefixe(NomChamp);
  NoTable := PrefixeToNum(PrfxTable);
  if NoTable <= 0 then Exit;
  for Ind := Low(V_PGI.DEChamps[NoTable]) to High(V_PGI.DEChamps[NoTable]) do
    if V_PGI.DEChamps[NoTable, Ind].Nom = NomChamp then
    begin
      Result := V_PGI.DEChamps[NoTable, Ind].Control;
      Break;
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si une lettre code est positionnée dans l'indicateur
Suite ........ : 'Controle' d'un champ
Mots clefs ... : FO
*****************************************************************}

function FOTestControlChamp(NomChamp, Flag: string): Boolean;
begin
  Result := (Pos(Flag, FOControlChamp(NomChamp)) > 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne la valeur vide d'un champ en fonction de son type
Mots clefs ... : FO
*****************************************************************}

function FOChampValeurVide(NomChamp: string; TypeChamp: string = ''): variant;
begin
  if TypeChamp = '' then TypeChamp := ChampToType(NomChamp); // retourne '??' si champ inconnu

  if (TypeChamp = 'INTEGER') or (TypeChamp = 'SMALLINT') or (TypeChamp = 'DOUBLE') or
    (TypeChamp = 'RATE') or (TypeChamp = 'EXTENDED') then
    Result := 0
  else if (TypeChamp = 'DATA') or (TypeChamp = 'BLOB') then
    Result := #0
  else if TypeChamp = 'DATE' then
    Result := IDate1900
  else if TypeChamp = 'BOOLEAN' then
    Result := '-'
  else
    Result := '';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 30/06/2004
Description .. : Vérifie si la valeur d'un champ est vide
Mots clefs ... : FO
*****************************************************************}

function FOTestChampVide(NomChamp: string; ValeurChamp: Variant): Boolean;
var
  sTypeChamp: string;
begin
//  if (VarIsEmpty(ValeurChamp)) or (VarIsNull(ValeurChamp)) then // JTR - Enlever VarIsNull
  if VarIsEmpty(ValeurChamp) then
    Result := True
  else
  begin
    sTypeChamp := ChampToType(NomChamp);
    Result := (ValeurChamp = FOChampValeurVide(NomChamp, sTypeChamp));
    if not (Result) and ((sTypeChamp = 'DATA') or (sTypeChamp = 'BLOB')) then
      Result := (ValeurChamp = '');
  end;
end;

{==============================================================================================}
{====================================== CLAVIER ===============================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Simule la frappe d'une touche du clavier
Mots clefs ... : FO
*****************************************************************}

procedure FOSimuleClavier(Touche: Word; Shift: Boolean = False; Attente: Integer = 0);
var Key: Word;
begin
  if Attente > 0 then Delay(Attente);
  if Shift then
  begin
    Key := MapVirtualKey(VK_SHIFT, 0);
    Keybd_Event(VK_SHIFT, Key, 0, 0);
  end;
  Key := MapVirtualKey(Touche, 0);
  Keybd_Event(Touche, Key, 0, 0);
  Key := MapVirtualKey(Touche, 0);
  Keybd_Event(Touche, Key, KEYEVENTF_KEYUP, 0);
  if Shift then
  begin
    Key := MapVirtualKey(VK_SHIFT, 0);
    Keybd_Event(VK_SHIFT, Key, KEYEVENTF_KEYUP, 0);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un pavé tactile est défini dans le paramétrage de
Suite ........ : lacaisse courante
Mots clefs ... : FO
*****************************************************************}

function FOExisteClavierEcran: Boolean;
var Clavier: string;
begin
  Clavier := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_CLAVIERECRAN')) then Clavier := VH_GC.TOBPCaisse.GetValue('GPK_CLAVIERECRAN');
  Result := (Clavier = 'X');
end;

{==============================================================================================}
{================================= MANIPULATION DES TOBS ======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Recrée une TOB à partir d'une autre TOB
Mots clefs ... : FO
*****************************************************************}

procedure FORecopieTob(TOBOrg: TOB; var TOBDest: TOB);
var NomTable: string;
begin
  if TOBOrg = nil then Exit;
  NomTable := '';
  if TOBDest <> nil then
  begin
    NomTable := TOBDest.NomTable;
    TOBDest.Free;
    TOBDest := nil;
  end;
  TOBDest := TOB.Create(NomTable, nil, -1);
  TOBDest.Dupliquer(TOBOrg, True, True, True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Dépile la TOB associé à une grille
Mots clefs ... : FO
*****************************************************************}

procedure FODepileTOBGrid(GS: THGrid; TOBG: TOB; ACols: array of Integer);
///////////////////////////////////////////////////////////////////////////////////////
  function EstRempli(NoLig: Integer): Boolean;
  var iCol: Integer;
  begin
    Result := False;
    for iCol := Low(ACols) to High(ACols) do
    begin
      Result := ((Result) or (GS.Cells[ACols[iCol], NoLig] <> ''));
      if Result then Break;
    end;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
var Ind, NoCol, Limite, MaxL: Integer;
begin
  if GS.Row < GS.FixedRows then Exit;
  Limite := TOBG.Detail.Count + 1;
  MaxL := Max(Limite, GS.Row);
  for Ind := MaxL downto GS.FixedRows + 1 do
  begin
    if not EstRempli(Ind) then Limite := Ind else Break;
  end;
  for Ind := TOBG.Detail.Count - 1 downto Limite - 1 do
  begin
    TOBG.Detail[Ind].Free;
    for NoCol := GS.FixedCols to GS.ColCount - 1 do GS.Cells[NoCol, Ind + 1] := '';
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Fusionne une TOB fille dans une autre TOB
Mots clefs ... : FO
*****************************************************************}

procedure FOMergeTobFille(TOBPrinc, TOBSecond: TOB; NomCle: string; NouveauChamp: array of string);
var TOBP, TOBS: TOB;
  Ind, Nbc: Integer;
  sNomChamp: string;
  ValeurCle: Variant;
begin
  if (TOBPrinc = nil) or (TOBSecond = nil) then Exit;
  for Ind := 0 to TOBSecond.Detail.Count - 1 do
  begin
    TOBS := TOBSecond.Detail[Ind];
    // recherche de la TOB principale équivalente
    ValeurCle := TOBS.GetValue(NomCle);
    TOBP := TOBPrinc.FindFirst([NomCle], [ValeurCle], False);
    if TOBP = nil then
    begin
      TOBP := TOB.Create(TOBPrinc.NomTable, TOBPrinc, -1);
      TOBP.AddChampSupValeur(NomCle, ValeurCle);
    end;
    // ajout des champs de la TOB secondaire inexistants dans la TOB principale
    for Nbc := Low(NouveauChamp) to High(NouveauChamp) do
    begin
      sNomChamp := NouveauChamp[Nbc];
      TOBP.AddChampSupValeur(sNomChamp, TOBS.GetValue(sNomChamp));
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Fusionne une TOB dans une autre TOB
Mots clefs ... : FO
*****************************************************************}

procedure FOMergeTob(TOBPrinc, TOBSecond: TOB; OverWrite: Boolean = False);
var Ind: Integer;
  sNomChamp: string;
begin
  if (TOBPrinc = nil) or (TOBSecond = nil) then Exit;
  for Ind := 0 to TOBSecond.NbChamps do
  begin
    sNomChamp := TOBSecond.GetNomChamp(Ind);
    if sNomChamp <> '' then
    begin
      if TOBPrinc.FieldExists(sNomChamp) then
      begin
        if OverWrite then TOBPrinc.PutValue(sNomChamp, TOBSecond.GetValue(sNomChamp));
      end else
      begin
        TOBPrinc.AddChampSupValeur(sNomChamp, TOBSecond.GetValue(sNomChamp));
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Modifie la valeur d'un champ des filles d'une TOB
Mots clefs ... : FO
*****************************************************************}

procedure FOPutValueDetail(TOBPrinc: TOB; ChampMaj: string; VV: Variant);
var Ind, NoChp: Integer;
  TOBL: TOB;
begin
  if TOBPrinc = nil then Exit;
  if TOBPrinc.Detail.Count <= 0 then Exit;
  NoChp := -1;
  for Ind := 0 to TOBPrinc.Detail.Count - 1 do
  begin
    TOBL := TOBPrinc.Detail[Ind];
    if NoChp <= 0 then NoChp := TOBL.GetNumChamp(ChampMaj);
    if NoChp > 0 then TOBL.PutValeur(NoChp, VV) else Break;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/06/2003
Modifié le ... : 19/06/2003
Description .. : Modifie la valeur d'un champ supplémentaire ou le crée si
Suite ........ : besoin
Mots clefs ... : FO
*****************************************************************}

procedure FOMajChampSupValeur(TOBL: TOB; NomChamp: string; Value: variant);
begin
  if (TOBL = nil) or (NomChamp = '') then Exit;

  if TOBL.FieldExists(NomChamp) then
    TOBL.PutValue(NomChamp, Value)
  else
    TOBL.AddChampSupValeur(NomChamp, Value);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/06/2003
Modifié le ... : 19/06/2003
Description .. : Retourne la valeur d'un champ supplémentaire
Mots clefs ... : FO
*****************************************************************}

function FOGetChampSupValeur(TOBL: TOB; NomChamp: string; Value: variant): variant;
begin
  Result := Value ;
  if (TOBL = nil) or (NomChamp = '') then Exit;

  if TOBL.FieldExists(NomChamp) then Result := TOBL.GetValue(NomChamp);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/06/2003
Modifié le ... : 19/06/2003
Description .. : Teste la valeur d'un champ supplémentaire
Mots clefs ... : FO
*****************************************************************}

function FOTesteChampSupValeur(TOBL: TOB; NomChamp: string; Value: variant): boolean;
begin
  Result := False ;
  if (TOBL = nil) or (NomChamp = '') then Exit;

  if (TOBL.FieldExists(NomChamp)) and (Value = TOBL.GetValue(NomChamp)) then
    Result := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/06/2003
Modifié le ... : 04/07/2003
Description .. : Teste si un champ supplémentaire existe et s'il a une valeur
Suite ........ : non vide.
Mots clefs ... : FO
*****************************************************************}

function FOChampSupValeurNonVide(TOBL: TOB; NomChamp: string): boolean;
var
  ValeurChamp: variant;
begin
  Result := False ;
  if (TOBL = nil) or (NomChamp = '') then Exit;

  if TOBL.FieldExists(NomChamp) then
  begin
    ValeurChamp := TOBL.GetValue(NomChamp);
//    if (not VarIsEmpty(ValeurChamp)) and (not VarIsNull(ValeurChamp)) and // JTR - Enlever VarIsNull
    if not VarIsEmpty(ValeurChamp) and (ValeurChamp <> #0) and (ValeurChamp <> '') then
    begin
      Result := True;
    end;
  end;
end;

{==============================================================================================}
{========================== FONCTIONS POUR LA GESTION DES TPE =================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un TPE est défini pour la caisse courante
Mots clefs ... : FO
*****************************************************************}

function FOExisteTPE: Boolean;
var Tpe: string;
begin
  Tpe := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_TPETYPE')) then Tpe := VH_GC.TOBPCaisse.GetValue('GPK_TPETYPE');
  Result := ((Tpe <> '') and (Tpe <> 'AUC'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne les paramètres de la liaison avec le TPE
Mots clefs ... : FO
*****************************************************************}

function FODonneParamTPE(Caisse: string; var Tpe, Port, Param: string): Boolean;
var TOBC: TOB;
begin
  Tpe := '';
  Port := '';
  Param := '';
  if Caisse <> FOCaisseCourante then
  begin
    TOBC := TOB.Create('PARCAISSE', nil, -1);
    TOBC.SelectDB('"' + Caisse + '"', nil);
  end else TOBC := VH_GC.TOBPCaisse;
  if (VH_GC.TOBPCaisse <> nil) and (TOBC.FieldExists('GPK_TPETYPE')) then
  begin
    Tpe := TOBC.GetValue('GPK_TPETYPE');
    Port := TOBC.GetValue('GPK_TPEPORT');
    Param := TOBC.GetValue('GPK_TPEBAUDS') + ';'
      + '003;' // parité paire
    + '003;' // 7 bits
    + '001;' // 1 bit d'arrêt
    + '001'; // aucun contrôle de flux
  end;
  Result := (Tpe <> '');
  if Caisse <> FOCaisseCourante then TOBC.Free;
end;

{==============================================================================================}
{========================== FONCTIONS POUR L'AFFICHEUR EXTERNE ================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un afficheur externe est défini pour la caisse
Suite ........ : courante
Mots clefs ... : FO
*****************************************************************}

function FOExisteAfficheur: Boolean;
var Afficheur: string;
begin
  Afficheur := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_AFFTYPE')) then Afficheur := VH_GC.TOBPCaisse.GetValue('GPK_AFFTYPE');
  Result := ((Afficheur <> '') and (Afficheur <> 'AUC'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne les paramètres de la liaison avec l'afficheur externe
Mots clefs ... : FO
*****************************************************************}

function FODonneParamAfficheur(var Afficheur, Port, Param: string): boolean;
begin
  Afficheur := '';
  Port := '';
  Param := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_AFFTYPE')) then
  begin
    Afficheur := VH_GC.TOBPCaisse.GetValue('GPK_AFFTYPE');

    if VH_GC.TOBPCaisse.FieldExists('GPK_AFFPORT') then
      Port := VH_GC.TOBPCaisse.GetValue('GPK_AFFPORT')
    else
      VH_GC.TOBPCaisse.AddChampSup('GPK_AFFPORT', False);
    if Trim(Port) = '' then
    begin
      Port := VH_GC.TOBPCaisse.GetValue('GPK_PRTPORT');
      VH_GC.TOBPCaisse.PutValue('GPK_AFFPORT', Port);
    end;

    if VH_GC.TOBPCaisse.FieldExists('GPK_AFFPARAM') then
      Param := VH_GC.TOBPCaisse.GetValue('GPK_AFFPARAM')
    else
      VH_GC.TOBPCaisse.AddChampSup('GPK_AFFPARAM', False);
    if Trim(Param) = '' then
    begin
      Param := VH_GC.TOBPCaisse.GetValue('GPK_PRTBAUDS') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTPARITE') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTNBBIT') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTSTOPBIT') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTCTRLFLUX') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTMODEBIDI');
      VH_GC.TOBPCaisse.PutValue('GPK_AFFPARAM', Param);
    end;
  end;
  Result := (Afficheur <> '');
end;

{==============================================================================================}
{========================== FONCTIONS POUR LE TIROIR CAISSE ===================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un tiroir caisse est défini pour la caisse courante
Mots clefs ... : FO
*****************************************************************}

function FOExisteTiroir: Boolean;
var Tiroir: string;
begin
  Tiroir := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_TIROIRTYPE')) then Tiroir := VH_GC.TOBPCaisse.GetValue('GPK_TIROIRTYPE');
  Result := ((Tiroir <> '') and (Tiroir <> 'AUC'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne les paramètres de la liaison avec le tiroir caisse
Mots clefs ... : FO
*****************************************************************}

function FODonneParamTiroir(var Tiroir, Port, Param: string): boolean;
begin
  Tiroir := '';
  Port := '';
  Param := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_TIROIRTYPE')) then
  begin
    Tiroir := VH_GC.TOBPCaisse.GetValue('GPK_TIROIRTYPE');

    if VH_GC.TOBPCaisse.FieldExists('GPK_TIROIRPORT') then
      Port := VH_GC.TOBPCaisse.GetValue('GPK_TIROIRPORT')
    else
      VH_GC.TOBPCaisse.AddChampSup('GPK_TIROIRPORT', False);
    if Trim(Port) = '' then
    begin
      Port := VH_GC.TOBPCaisse.GetValue('GPK_PRTPORT');
      VH_GC.TOBPCaisse.PutValue('GPK_TIROIRPORT', Port);
    end;

    if VH_GC.TOBPCaisse.FieldExists('GPK_TIROIRPARAM') then
      Param := VH_GC.TOBPCaisse.GetValue('GPK_TIROIRPARAM')
    else
      VH_GC.TOBPCaisse.AddChampSup('GPK_TIROIRPARAM', False);
    if Trim(Param) = '' then
    begin
      Param := VH_GC.TOBPCaisse.GetValue('GPK_PRTBAUDS') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTPARITE') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTNBBIT') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTSTOPBIT') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTCTRLFLUX') + ';'
        + VH_GC.TOBPCaisse.GetValue('GPK_PRTMODEBIDI');
      VH_GC.TOBPCaisse.PutValue('GPK_TIROIRPARAM', Param);
    end;
    if not VH_GC.TOBPCaisse.FieldExists('GPK_TIROIRPIN5') then
      VH_GC.TOBPCaisse.AddChampSupValeur('GPK_TIROIRPIN5', '-');
    Param := Param + ';' + VH_GC.TOBPCaisse.GetValue('GPK_TIROIRPIN5');
  end;
  Result := (Tiroir <> '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si l'ouverture du tiroir caisse doit être déclenchée en
Suite ........ : fonction des modes de paiement utilisés dans la pièce
Mots clefs ... : FO
*****************************************************************}

function FOVerifOuvreTiroir(TOBEches: TOB): boolean;
var SiEspece: string;
  Ind: integer;
  TOBL: TOB;
begin
  SiEspece := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_TIROIRESP')) then
    SiEspece := VH_GC.TOBPCaisse.GetValue('GPK_TIROIRESP');
  if SiEspece = 'X' then
  begin
    // il doit exister au moins une échéance de type espèce
    Result := False;
    if TOBEches = nil then Exit;
    for Ind := 0 to TOBEches.Detail.Count - 1 do
    begin
      TOBL := TOBEches.Detail[Ind];
      if (TOBL.FieldExists('TYPEMODEPAIE')) and
        (TOBL.GetValue('TYPEMODEPAIE') = TYPEPAIEESPECE) then
      begin
        Result := True;
        Exit;
      end;
    end;
  end else
    Result := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/10/2002
Modifié le ... : 10/10/2002
Description .. : Ouvre le tiroir caisse avec vérification du paramétrage du
Suite ........ : mode automatiquement depuis l'ouverture ou la fermeture
Suite ........ : de journée.
Mots clefs ... : FO
*****************************************************************}

procedure FOOuvreTiroir(FromUser: Boolean = False; FinJour: Boolean = False);
{$IFDEF FOS5}
var Ok: Boolean;
  Stg: string;
  {$ENDIF}
begin
  {$IFDEF FOS5}
  if FinJour then Ok := (FOGetParamCaisse('GPK_OUVTIROIRFIN') = 'X')
  else Ok := True;
  if (Ok) and (FOExisteTiroir) then
  begin
    if FromUser then Stg := TraduireMemoire('Demandé par l''utilisateur')
    else Stg := '';
    if Ouvretiroir then FOIncrJoursCaisse(jfoNbOuvTiroir, Stg);
  end;
  {$ENDIF}
end;

{==============================================================================================}
{======================= FONCTIONS POUR L'IMPRESSION DES TICKETS ==============================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si une imprimante ticket est définie pour la caisse
Suite ........ : courante
Mots clefs ... : FO
*****************************************************************}

function FOExisteLP: Boolean;
var Imprimante: string;
begin
  Imprimante := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_PRTTYPE')) then Imprimante := VH_GC.TOBPCaisse.GetValue('GPK_PRTTYPE');
  Result := ((Imprimante <> '') and (Imprimante <> 'AUC'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si l'imprimante ticket imprime aussi les chèques
Mots clefs ... : FO
*****************************************************************}

function FOExisteLPCheque: Boolean;
begin
  Result := False;
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_PRTTYPE')) and (VH_GC.TOBPCaisse.GetValue('GPK_PRTTYPE') <> '') then
  begin
    Result := (VH_GC.TOBPCaisse.GetValue('GPK_REMPLIRCHQ') = 'X');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne le type de l'imprimante ticket
Mots clefs ... : FO
*****************************************************************}

function FODonneTypeLP: string;
begin
  Result := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_PRTTYPE')) then Result := VH_GC.TOBPCaisse.GetValue('GPK_PRTTYPE');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne les paramètres de communication avec l'imprimante
Suite ........ : ticket
Mots clefs ... : FO
*****************************************************************}

function FODonneParamLP(var Imprimante, Port, Param: string): Boolean;
begin
  Imprimante := '';
  Port := '';
  Param := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_PRTTYPE')) then
  begin
    Imprimante := VH_GC.TOBPCaisse.GetValue('GPK_PRTTYPE');
    Port := VH_GC.TOBPCaisse.GetValue('GPK_PRTPORT');
    Param := VH_GC.TOBPCaisse.GetValue('GPK_PRTBAUDS') + ';'
      + VH_GC.TOBPCaisse.GetValue('GPK_PRTPARITE') + ';'
      + VH_GC.TOBPCaisse.GetValue('GPK_PRTNBBIT') + ';'
      + VH_GC.TOBPCaisse.GetValue('GPK_PRTSTOPBIT') + ';'
      + VH_GC.TOBPCaisse.GetValue('GPK_PRTCTRLFLUX') + ';'
      + VH_GC.TOBPCaisse.GetValue('GPK_PRTMODEBIDI');
  end;
  Result := (Imprimante <> '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Fabrique la clause where pour l'impression du ticket X
Mots clefs ... : FO
*****************************************************************}

function FOMakeWhereTicketX(Prefixe: string; Caisse: string = ''): string;
var NumZ: integer;
begin
  if Prefixe = '' then Exit;
  if Caisse = '' then Caisse := FOCaisseCourante;
  if Caisse = '' then Exit;
  NumZ := FOGetNumZCaisse(Caisse);
  Result := Prefixe + '_NATUREPIECEG=' + FOGetNatureTicket(False, True) + ' AND ' +
    Prefixe + '_CAISSE="' + Caisse + '" AND ' +
    Prefixe + '_NUMZCAISSE=' + IntToStr(NumZ) + ' AND ' +
    Prefixe + '_DATEPIECE="' + USDateTime(FOGetDateOuv(Caisse, NumZ)) + '"';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 22/08/2001
Modifié le ... : 22/08/2001
Description .. : Fabrique la clause where pour l'impression du ticket Z
Mots clefs ... : FO;
*****************************************************************}

function FOMakeWhereTicketZ(Prefixe, Caisse: string; DeNumZ, ANumZ: integer): string;
begin
  if Prefixe = '' then Exit;
  if Caisse = '' then Caisse := FOCaisseCourante;
  if Caisse = '' then Exit;
  Result := Prefixe + '_NATUREPIECEG=' + FOGetNatureTicket(False, True) + ' AND ' +
    Prefixe + '_CAISSE="' + Caisse + '" AND ' +
    Prefixe + '_NUMZCAISSE>=' + IntToStr(DeNumZ) + ' AND ' +
    Prefixe + '_NUMZCAISSE<=' + IntToStr(ANumZ) + ' AND ' +
    FOSelectNumZDate(Prefixe, DeNumZ, ANumZ, Caisse);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Fabrique la clause where pour l'impression du récapitulatif des vendeurs
Mots clefs ... : FO
*****************************************************************}

function FOMakeWhereRecapVendeurs(Prefixe, Caisse: string; DeNumZ, ANumZ: Integer): string;
begin
  if Prefixe = '' then Exit;
  if Caisse = '' then Caisse := FOCaisseCourante;
  if Caisse = '' then Exit;
  Result := Prefixe + '_NATUREPIECEG=' + FOGetNatureTicket(False, True) + ' AND ' +
    Prefixe + '_CAISSE="' + Caisse + '" AND ' +
    Prefixe + '_NUMZCAISSE>=' + IntToStr(DeNumZ) + ' AND ' +
    Prefixe + '_NUMZCAISSE<=' + IntToStr(ANumZ) + ' AND ' +
    FOSelectNumZDate(Prefixe, DeNumZ, ANumZ, Caisse);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne le code de l'état à utiliser pour un type d'impression
Mots clefs ... : FO
*****************************************************************}

function FODonneCodeEtat(TypeEtat: TTypeEtatFO; ValeurDefaut: Boolean; var Nature, CodeEtat, Titre: string; CodeOk : boolean=false): Boolean;
begin
  Titre := '';
  Nature := '';
  if not CodeOk then CodeEtat := '';
  case TypeEtat of
    efoTicket: // Ticket de vente
      begin
        Titre := 'Ticket de vente';
        Nature := NATUREMODTICKET;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODTIC');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'FFO';
      end;
    efoTicketX: // Ticket X
      begin
        Titre := 'Ticket X';
        Nature := NATUREMODTICKETZ;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODTICX');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'TKX';
      end;
    efoTicketZ: // Ticket Z
      begin
        Titre := 'Ticket Z';
        Nature := NATUREMODTICKETZ;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODTICZ');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'TKZ';
      end;
    efoRecapVend: // Récapitulatif par vendeur
      begin
        Titre := 'Récapitulatif par vendeurs';
        Nature := NATUREMODTICKETZ;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODRECAPV');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'VEN';
      end;
    efoListeRegle: // Liste des règlements
      begin
        Titre := 'Liste des règlements';
        Nature := NATUREMODTICKETZ;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODLISREG');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'REG';
      end;
    efoFdCaisse: // Liste du fond de caisse
      begin
        Titre := 'Fond de caisse';
        Nature := NATUREMODTICKETZ;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODFDC');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'FDC';
      end;
    efoTicketTest: // Ticket de test
      begin
        Titre := 'Ticket de test';
        Nature := NATUREMODTICKET;
        CodeEtat := 'XXX';
        ;
      end;
    efoCheque: // Chèque
      begin
        Titre := 'Chèque';
        Nature := NATUREMODCHEQUE;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODCHEQUE');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'CHQ';
        ;
      end;
    efoRemiseBq: // Remise en banque
      begin
        Titre := 'Remise en banque';
        Nature := NATUREMODCHEQUE;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODREMIS');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'BRC';
        ;
      end;
    efoBonAchat: // Bon d'achat
      begin
        Titre := 'Bon d''achat';
        Nature := NATUREMODTICKET;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODBONACH');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'BAC';
      end;
    efoBonAvoir: // Bon d'avoir
      begin
        Titre := 'Bon d''avoir';
        Nature := NATUREMODTICKET;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODBONAV');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'BAV';
      end;
    efoBonArrhes: // Bon de versement d'arrhes
      begin
        Titre := 'Bon de versement d''arrhes';
        Nature := NATUREMODTICKET;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODBONARR');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'BAR';
      end;
    efoTransfert: // Bon de transfert
      begin
        Titre := 'Bon de transfert';
        Nature := NATUREMODTICKET;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODBONTID');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'TEM';
      end;
    efoCommande: // Bon de commande
      begin
        Titre := 'Bon de commande';
        Nature := NATUREMODTICKET;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODBONCDE');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'FCF';
      end;
    efoReception: // Bon de réception
      begin
        Titre := 'Bon de réception';
        Nature := NATUREMODTICKET;
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODBONREC');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'BLF';
      end;
    efoStatFam: // Statistiques par famille
      begin
        Titre := 'Statistiques par famille';
        Nature := NATUREMODTICKETZ;
        CodeEtat := 'FAM';
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODSTAFAM');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'FAM';
      end;
    efoStatRem: // Statistiques sur les remises
      begin
        Titre := 'Statistiques sur les remises';
        Nature := NATUREMODTICKETZ;
        CodeEtat := 'REM';
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODSTAREM');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'REM';
      end;
    efoListeArtVendu: // Liste des articles vendus
      begin
        Titre := 'Liste des articles vendus';
        Nature := NATUREMODTICKETZ;
        if (V_PGI.LaSerie > S3) and (GetParamSoc('SO_ARTLOOKORLI')) then
          CodeEtat := 'AR5'
        else
          CodeEtat := 'ART';
        if FOVerifieVHGCCaisse then CodeEtat := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODLISART');
        if (ValeurDefaut) and (CodeEtat = '') then CodeEtat := 'ART';
      end;
    // JTR - Statistiques
    efoStatsGP: // Pièces
      begin
        Titre := 'Statistiques';
        Nature := NATUREMODSTATSGP;
      end;
    efoStatsGPE: // Encaissements
      begin
        Titre := 'Statistiques';
        Nature := NATUREMODSTATSGPE;
      end;
    efoStatsGPV: // Vendeurs
      begin
        Titre := 'Statistiques';
        Nature := NATUREMODSTATSGPV;
      end;
    efoStatsGPZ: // Spécifiques
      begin
        Titre := 'Statistiques';
        Nature := NATUREMODSTATSGPZ;
      end;
    // Fin JTR - Statistiques
  end;
  if Titre <> '' then Titre := TraduireMemoire(Titre);
  Result := (CodeEtat <> '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un code de l'état est défini pour un type
Suite ........ : d'impression
Mots clefs ... : FO
*****************************************************************}

function FOExisteCodeEtat(TypeEtat: TTypeEtatFO): Boolean;
{$IFDEF FOS5}
var Nature, CodeEtat, Titre, Langue: string;
  {$ENDIF}
begin
  Result := False;
  {$IFDEF FOS5}
  if FODonneCodeEtat(TypeEtat, False, Nature, CodeEtat, Titre) then
  begin
    Langue := V_PGI.LanguePerso;
    Result := ExisteModele('T', Nature, CodeEtat, Langue);   //, False
  end;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 09/10/2002
Description .. : Lance une impression en mode ticket à partir soit d'une
Suite ........ : TOB soit d'une clause WHERE
Mots clefs ... : FO
*****************************************************************}

procedure FOLanceImprimeLP(TypeEtat: TTypeEtatFO; Where: string; Apercu: Boolean; UneTOB: TOB; Code : string='');
{$IFDEF FOS5}
var Titre, Nature, CodeEtat, Imprimante, Port, Param: string;
  Err: TMC_Err;
  SQLBased: Boolean;
begin
  CodeEtat := Code;
  // Récupération des paramètres de l'impression
  FODonneCodeEtat(TypeEtat, False, Nature, CodeEtat, Titre, (Code <> ''));
  // Récupération des paramètres de l'imprimante ticket
  FODonneParamLP(Imprimante, Port, Param);
  if not Apercu then Apercu := (Imprimante = '');
  SQLBased := (UneTOB = nil);
  // Lancement de l'impression
  if (not FOImprimeLP('T', Nature, CodeEtat, Where, Imprimante, Port, Param, True, Apercu, SQLBased, UneTOB, Err))
    and (Err.Code <> 0) then
      FOMCErr(Err, Titre);
end;
{$ELSE}
begin
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 09/10/2002
Modifié le ... : 09/10/2002
Description .. : Lance une impression en mode ticket
Mots clefs ... : FO
*****************************************************************}
{$IFDEF FOS5}

function FOImprimeLP(TypeEtat, NatEtat, CodeEtat, Where, Imprimante, Port, Params: string; Initialise, PreView, SQLBased: Boolean; UneTOB: TOB; var Err:
  TMC_Err; IsImpControl: Boolean = FALSE; IsTestMateriel: Boolean = FALSE): Boolean; //XMG 22/06/01
var VersionDemo: Boolean;
begin
  // Pour imprimer "Version de démonstration" dans la base Formation
  VersionDemo := V_PGI.VersionDemo;
  if GetParamSoc('SO_GCFOBASEFORMATION') then V_PGI.VersionDemo := True;
  Result := ImprimeLP(TypeEtat, NatEtat, CodeEtat, Where, Imprimante, Port, Params, Initialise, PreView, SQLBased, UneTOB, Err, IsImpControl, IsTestMateriel);
  V_PGI.VersionDemo := VersionDemo;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/07/2002
Modifié le ... : 19/07/2002
Description .. : Affichage du message d'erreur retourné par le matériel
Suite ........ : caisse.
Mots clefs ... : FO
*****************************************************************}
{$IFDEF FOS5}

procedure FOMCErr(Erreur: TMC_Err; Titre: string);
var Stg: string;
begin
  Stg := Trim(Erreur.Libelle);
  if Stg = '' then Stg := TraduireMemoire('Erreur n°') + IntToStr(Erreur.Code);
  PGIBox(Stg, Titre);
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/12/2002
Modifié le ... : 31/12/2002
Description .. : Lancement de l'éditeur de ticket par DispatchTT
Mots clefs ... :
*****************************************************************}

procedure FODispatchParamModele(Num: integer; Action: TActionFiche; Lequel, TT, Range: string);
{$IFNDEF EAGLCLIENT}
var TypEtat, NatEtat, Stg: string;
  {$ENDIF}
begin
  {$IFNDEF EAGLCLIENT}
  // Extraction de la clause plus du type et de la nature de l'état
  Range := UpperCase(Range);
  Range := FindEtReplace(Range, ' AND ', ';', True);
  Range := FindEtReplace(Range, ' OR ', ';', True);
  Range := FindEtReplace(Range, ' ', '', True);
  repeat
    Stg := ReadTokenSt(Range);
    if Copy(Stg, 1, 8) = 'MO_TYPE=' then
    begin
      ReadTokenPipe(Stg, '"');
      TypEtat := ReadTokenPipe(Stg, '"');
    end else
      if Copy(Stg, 1, 10) = 'MO_NATURE=' then
    begin
      ReadTokenPipe(Stg, '"');
      NatEtat := ReadTokenPipe(Stg, '"');
    end;
  until Range = '';
  // Lancement de l'éditeur de ticket
  if TypEtat = 'T' then
  begin
    Param_LPTexte(nil, 'T', NatEtat, Lequel, True);
    LibereSauvgardeLP;
  end else
    if TypEtat = 'L' then
    EditDocument('L', NatEtat, Lequel, True)
  else
    if TypEtat = 'E' then
    EditEtat('E', NatEtat, Lequel, True, nil, '', '');
  {$ENDIF}
end;

{==============================================================================================}
{==================================== POLICE D'UN CHAMP =======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Modifie la taille de la police d'un contrôle
Mots clefs ... : FO
*****************************************************************}

procedure FoSetFontSize(Composant: TComponent; Action: string; Taille: Integer);
var Valeur: Integer;
begin
  if Composant is TControl then
  begin
    Valeur := TFOFontControl(Composant).Font.Size;
    if Action = '+' then Valeur := Valeur + Taille else
      if Action = '-' then Valeur := Valeur - Taille else Valeur := Taille;
    TFOFontControl(Composant).Font.Size := Valeur;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Modifie la couleur d'un contrôle
Mots clefs ... : FO
*****************************************************************}

procedure FoSetFontColor(FF: TForm; NomComp: string; Color: TColor);
var Composant: TComponent;
begin
  if FF = nil then exit;
  Composant := FF.FindComponent(NomComp);
  if (Composant <> nil) and (Composant is TControl) then TFOFontControl(Composant).Font.Color := Color;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Redimensionne avec affichage des étiquettes les boutons
Suite ........ : d'une fiche
Mots clefs ... : FO
*****************************************************************}

procedure FoShowEtiquette(F: Forms.TForm; Visu: boolean);
var i, j: integer;
  TB: TToolBar97;
  B: TToolBarButton97;
begin
  for i := 0 to F.ComponentCount - 1 do
    if (F.Components[i] is TToolBar97) and not (F.Components[i] is TMenu97) then
    begin
      TB := TToolBar97(F.Components[i]);
      for j := 0 to TB.ControlCount - 1 do
        if TB.Controls[j] is TToolBarButton97 then
        begin
          B := TToolBarButton97(TB.Controls[j]);
          B.Spacing := 0;
          if Visu then
          begin
            B.Height := {72} {48} 35;
            if B.Width < 40 then B.Width := {80} 60 else
              if B.Width = 40 then B.Width := {81} 61;
            if B.Layout = blGlyphTop then B.DisplayMode := dmBoth;
          end else
          begin
            B.Height := 24;
            if B.Width = {81} 61 then B.Width := 40 else
              if B.Width = {80} 60 then B.Width := 24;
            if B.Layout = blGlyphTop then B.DisplayMode := dmGlyphOnly;
          end;
        end;
    end;
end;

{==============================================================================================}
{============================= MANIPULATION DES CHAINES DE CARACTERES =========================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Détermine si la saisie est un entier
Suite ........ : (suppression des 0 éventuels en début de saisie)
Mots clefs ... : FO
*****************************************************************}

function FOIsInteger(st: string): string;
var i: integer;
  chiffre: string;
  stofzero: boolean; // la saisie ne comporte-t-elle que des zéros ? //
begin
  stofzero := true;
  for i := 1 to length(st) do
  begin
    chiffre := copy(st, i, 1);
    if chiffre <> '0' then stofzero := false;
    if pos(chiffre, '0123456789') = 0 then break;
  end;
  if i <> (length(st) + 1) then result := 'false' else
    if stofzero = true then result := '0' else
  begin
    i := 1;
    while i <= length(st) do
      if copy(st, i, 1) = '0' then st := copy(st, i + 1, length(st) - i) else break;
    ;
    result := st;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne une chaine
Mots clefs ... : FO
*****************************************************************}

function FOReverseSt(st: string): string;
var chaine: string;
  i: integer;
begin
  chaine := '';
  for i := 1 to length(st) do
    chaine := chaine + copy(st, length(st) + 1 - i, 1);
  result := chaine;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Supprime les espaces d'une chaine comportant des chiffres
Suite ........ : et ',' ou '.'
Mots clefs ... : FO
*****************************************************************}

function FOCutSpaceSt(St: string): string;
var Chaine, Signe: string;
  Ind, Compteur: integer;
begin
  Signe := '';
  Chaine := '';
  Compteur := 0;
  if copy(St, 1, 1) = '-' then Signe := '-';
  for Ind := 1 to length(St) do
  begin
    if copy(St, Ind, 1) = ',' then inc(Compteur);
    if (pos(copy(St, Ind, 1), '0123456789,.') <> 0) and (Compteur <= 1) then
      Chaine := Chaine + copy(St, Ind, 1)
    else if Compteur > 1 then Break;
  end;
  Result := Signe + Chaine;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 24/07/2001
Description .. : Extrait une sous-chaîne d'une chaîne.
Mots clefs ... : FO
*****************************************************************}

function FOExtract(var Stg: string; Index, Count: Integer): string;
begin
  Result := Copy(Stg, Index, Count);
  Delete(Stg, Index, Count);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si une chaîne de caractères appartient à une liste au
Suite ........ : format multi-combo
Mots clefs ... : FO
*****************************************************************}

function FOStrCmp(Chaine, Liste: string; GereTous: Boolean = False): Boolean;
begin
  Result := False;
  if Chaine = '' then Exit;
  if GereTous then
  begin
    // Cas du <<Tous>>
    if (Pos('<<', Liste) = 1) and (Pos('>>', Liste) = Length(Liste) - 1) then Liste := '';
    Result := True;
  end;
  if Liste = '' then Exit;
  if Liste[1] <> ';' then Liste := ';' + Liste;
  if Liste[Length(Liste)] <> ';' then Liste := Liste + ';';
  if Chaine[1] <> ';' then Chaine := ';' + Chaine;
  if Chaine[Length(Chaine)] <> ';' then Chaine := Chaine + ';';
  Result := (Pos(Chaine, Liste) <> 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : formate une chaîne de caractère à partir d'un montant selon
Suite ........ : les caractéristiques de la devise (nombre de décimales et
Suite ........ : symbole)
Mots clefs ... : FO
*****************************************************************}

function FOStrFMontant(Montant: Extended; Taille: Integer; Devise: string): string;
var NbDec: Integer;
  Symbole: string;
  TOBDev: TOB;
begin
  if Devise = V_PGI.DevisePivot then
  begin
    NbDec := V_PGI.OKDecV;
    Symbole := V_PGI.SymbolePivot;
  end else
  begin
    TOBDev := TOB.Create('DEVISE', nil, -1);
    TOBDev.SelectDB('"' + Devise + '"', nil, False);
    NbDec := TOBDev.GetValue('D_DECIMALE');
    Symbole := TOBDev.GetValue('D_SYMBOLE');
    TOBDev.Free;
  end;
  if FOTestCodeEuro(Devise) then Symbole := SIGLEEURO;
  Result := StrfMontant(Montant, Taille, NbDec, Symbole, True);
end;

{==============================================================================================}
{======================================== SQL =================================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 05/12/2002
Description .. : Extrait la condition de sélection d'une page critères d'un
Suite ........ : MUL sans la clause WHERE ni les ordres de tri et de
Suite ........ : regroupement
Mots clefs ... : FO
*****************************************************************}

function FoExtractWhere(PageCritere: TPageControl): string;
var Ind: integer;
begin
  Result := UpperCase(RecupWhereCritere(PageCritere));
  Ind := Pos('WHERE ', Result);
  if Ind > 0 then Delete(Result, Ind, 6);
  Ind := Pos('ORDER BY ', Result);
  if Ind > 0 then Delete(Result, Ind, (Length(Result) - Ind + 1));
  Ind := Pos('GROUP BY ', Result);
  if Ind > 0 then Delete(Result, Ind, (Length(Result) - Ind + 1));
  Result := Trim(Result);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si un select retroune des enregistrements
Mots clefs ... : FO
*****************************************************************}

function FOExisteSQL(sSql: string): Boolean;
var QQ: TQuery;
  Stg: string;
  Ind: integer;
begin
  Result := False;
  if UpperCase(Copy(sSql, 1, 6)) <> 'SELECT' then Exit;
  Ind := Pos(' FROM ', UpperCase(sSql));
  if Ind > 0 then Stg := 'select count(*)' + Copy(sSql, Ind, (Length(sSql) - Ind + 1));
  QQ := OpenSQL(Stg, True);
  if not QQ.EOF then Result := (QQ.Fields[0].AsInteger > 0);
  Ferme(QQ);
end;

{==============================================================================================}
{====================================== DEVISES ===============================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 10/01/2003
Description .. : charge en TOB une devise à partir de son code
Mots clefs ... : FO
*****************************************************************}

function FOChargeDevise(pD_: TOB; Code: string): boolean;
begin
  Result := False;
  if pD_ = nil then Exit;
  pD_.PutValue('D_DEVISE', Code);
  Result := pD_.LoadDB;
  if (Result) and (Code = CODEEURO) then pD_.PutValue('D_SYMBOLE', SIGLEEURO);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 18/10/2001
Description .. : Retourne le nombre de décimales d'une devise
Mots clefs ... : FO
*****************************************************************}

function FODonneDeviseNbDec(Code: string): Integer;
var QQ: TQuery;
begin
  if Code = CODEEURO then Result := V_PGI.OkDecE else Result := V_PGI.OkDecV;
  QQ := OpenSQL('SELECT D_DECIMALE FROM DEVISE WHERE D_DEVISE="' + Code + '"', TRUE);
  if not QQ.EOF then Result := QQ.FindField('D_DECIMALE').AsInteger;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 18/10/2001
Description .. : Retourne le symbole d'une devise
Mots clefs ... : FO
*****************************************************************}

function FODonneDeviseSymbole(Code: string): string;
var QQ: TQuery;
begin
  if Code <> CODEEURO then
  begin
    QQ := OpenSQL('SELECT D_SYMBOLE FROM DEVISE WHERE D_DEVISE="' + Code + '"', TRUE);
    if not QQ.EOF then Result := QQ.FindField('D_SYMBOLE').AsString;
    Ferme(QQ);
  end else Result := SIGLEEURO;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne le code de la devise EURO
Mots clefs ... : FO
*****************************************************************}

function FOChargeDeviseEuro: string;
var QQ: TQuery;
begin
  Result := CODEEURO;
  QQ := OpenSQL('SELECT D_DEVISE FROM DEVISE WHERE D_CODEISO="' + CODEEURO + '" AND D_FERME="X"', TRUE);
  if not QQ.EOF then Result := QQ.FindField('D_DEVISE').AsString;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne le code de la devise EURO défini pour la caisse
Suite ........ : courante
Suite ........ : (optimisation = le code de la devise Euro est ajouté dans le
Suite ........ : paramétrage de la caisse courante)
Mots clefs ... : FO
*****************************************************************}

function FODonneCodeEuro: string;
begin
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_CODEEURO')) then Result := VH_GC.TOBPCaisse.GetValue('GPK_CODEEURO')
  else Result := CODEEURO;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Teste si une devise est l'EURO
Mots clefs ... : FO
*****************************************************************}

function FOTestCodeEuro(CodeDev: string): Boolean;
begin
  Result := (CodeDev = FODonneCodeEuro);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne le nom de la devise et des "centimes", par exemple
Suite ........ : 'FRANC;CENTIME'
Suite ........ : Attention ne traite que les francs et les euros !
Mots clefs ... : FO
*****************************************************************}

function FODonneMaskDevise(CodeDev: string): string;
begin
  if FOTestCodeEuro(CodeDev) then Result := 'EURO;CENT' else Result := 'FRANC;CENTIME'; // en attendant mieux ????
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne le mode de paiement pour les écarts de change
Mots clefs ... : FO
*****************************************************************}

function FODonneEcartChange: string;
begin
  Result := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_MDPECART')) then Result := VH_GC.TOBPCaisse.GetValue('GPK_MDPECART');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne le mode de paiement pour le rendu monnaie
Mots clefs ... : FO
*****************************************************************}

function FODonneRendu: string;
begin
  Result := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_MDPRENDU')) then Result := VH_GC.TOBPCaisse.GetValue('GPK_MDPRENDU');
end;

{==============================================================================================}
{=================================== GESTION DES CLIENTS ======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne le code du client par défaut pour les tickets défini
Suite ........ : dans le paramétrage de la caisse courante
Mots clefs ... : FO
*****************************************************************}

function FODonneClientDefaut: string;
begin
  Result := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_TIERS')) then Result := VH_GC.TOBPCaisse.GetValue('GPK_TIERS');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/04/2004
Modifié le ... : 23/04/2004
Description .. : vérifie si un tiers est un tiers fictif.
Mots clefs ... : FO
*****************************************************************}

function FOTiersEstFictif(TOBTiers: TOB): boolean;
var
  Ind: integer;
begin
  Result := False;
  if TOBTiers = nil then Exit;

  Ind := TOBTiers.GetNumChamp('T_TIERSFICTIF');
  Result := (Ind >= 0) and (StrToBool(TOBTiers.GetValeur(Ind)));

  if not Result then Result := (TOBTiers.GetValue('T_TIERS') = FODonneClientDefaut) ;
end;

{==============================================================================================}
{=================================== GESTION DES DEMARQUES ====================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Indique si on gére le code démarque sur la caisse courante
Mots clefs ... : FO
*****************************************************************}

function FOGereDemarque(Obligatoire: Boolean): Boolean;
begin
  Result := False;
  if (VH_GC.TOBPCaisse = nil) or (not VH_GC.TOBPCaisse.FieldExists('GPK_DEMSAISIE')) then Exit;
  if VH_GC.TOBPCaisse.GetValue('GPK_DEMSAISIE') = 'X' then
  begin
    Result := True;
    if Obligatoire then Result := VH_GC.TOBPCaisse.GetValue('GPK_DEMOBLIG') = 'X';
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Charge en TOB les codes démarque
Mots clefs ... : FO
*****************************************************************}
procedure FOChargeDemarque;
var TOBDem: TOB;
begin
  if VH_GC.TOBPCaisse = nil then Exit;
  if VH_GC.TOBPCaisse.GetValue('GPK_DEMSAISIE') <> 'X' then Exit;
  // Libération si besoin de l'ancienne TOB
  TOBDem := FOGetPTobDemarque;
  if TOBDem <> nil then TOBDem.Free;
  // Chargement dans une  TOB fille de la TOB du paramétrage de la caisse
  TOBDem := TOB.Create('', VH_GC.TOBPCaisse, -1);
  TOBDem.Commentaire := 'Types remise';
  TOBDem.LoadDetailDB('TYPEREMISE', '', '', nil, False);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : JTR
Créé le ...... : 26/03/2004
Modifié le ... :
Description .. : Charge en TOB les articles financiers + tiers associés + infos comptable
Mots clefs ... : FO
*****************************************************************}
procedure FOChargeArtFiTiers;
var TOBDem: TOB;            
    Sql : String;
    Qry : TQuery;
begin
  if VH_GC.TOBPCaisse = nil then Exit;
  // Libération si besoin de l'ancienne TOB
  TOBDem := FOGetPTobArtFiTiers;
  if TOBDem <> nil then TOBDem.Free;
  TOBDem := TOB.Create('', VH_GC.TOBPCaisse, -1);
  TOBDem.Commentaire := 'Article Financiers + Tiers';
  Sql := 'SELECT GA_ARTICLE, GA_CODEARTICLE, GA_URLWEB, GA_TYPEARTFINAN, GA_DESIGNATION1, GA_DESIGNATION2, GFC_COMPTE, GFC_JOURNAL'
        +' FROM ARTICLE'
        +' LEFT JOIN ARTFINANCIERCOMPL ON GFC_ARTICLE=GA_ARTICLE AND GFC_ETABLISSEMENT="'+ VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT')+'"'
        +' WHERE GA_TYPEARTICLE="FI" AND GA_FERME="-" ORDER BY GA_CODEARTICLE';
  Qry:= OpenSQL(Sql, True);
  TOBDem.LoadDetailDB('ARTFITIERS', '', '', Qry, False);
  Ferme(Qry);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/11/2002
Modifié le ... : 18/11/2002
Description .. : Retourne la TOB des codes démarque de la caisse
Mots clefs ... : FO
*****************************************************************}

function FOGetPTobDemarque: TOB;
var Ind: Integer;
begin
  Result := nil;
  if VH_GC.TOBPCaisse = nil then Exit;
  if VH_GC.TOBPCaisse.GetValue('GPK_DEMSAISIE') <> 'X' then Exit;
  for Ind := 0 to VH_GC.TOBPCaisse.Detail.Count - 1 do
  begin
    if (VH_GC.TOBPCaisse.Detail[Ind].NomTable = '') and
      (VH_GC.TOBPCaisse.Detail[Ind].Commentaire = 'Types remise') then
    begin
      Result := VH_GC.TOBPCaisse.Detail[Ind];
      Break;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JTR
Créé le ...... : 26/03/2004
Modifié le ... :
Description .. : Retourne la TOB des articles financiers de la caisse
Mots clefs ... : FO
*****************************************************************}
function FOGetPTobArtFiTiers: TOB; // JTR Lien Opération Caisse avec Tiers
var Ind: Integer;
begin
  Result := nil;
  if VH_GC.TOBPCaisse = nil then Exit;
  for Ind := 0 to VH_GC.TOBPCaisse.Detail.Count - 1 do
  begin
    if (VH_GC.TOBPCaisse.Detail[Ind].NomTable = '') and
      (VH_GC.TOBPCaisse.Detail[Ind].Commentaire = 'Article Financiers + Tiers') then
    begin
      Result := VH_GC.TOBPCaisse.Detail[Ind];
      Break;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JTR
Créé le ...... : 06/05/2004
Modifié le ... :
Description .. : Retourne True s'il y a un article financier "spécial" dans la pièce
Mots clefs ... : FO
*****************************************************************}
function IsArtFiByType(TobPiece, TobArtFi : TOB; TypeFi : string): boolean;
var TobTmp : TOB;
begin
  Result := False;
  TobTmp := TobArtFi.FindFirst(['GA_TYPEARTFINAN'],[TypeFi],True);
  while TobTmp <> nil do
  begin
    if TobPiece.FindFirst(['GL_ARTICLE'],[TobTmp.GetValue('GA_ARTICLE')],True) <> nil then
    begin
      Result := True;
      break;
    end;
    TobTmp := TobArtFi.FindNext(['GA_TYPEARTFINAN'],[TypeFi],True);
  end;
end;

function IsArtFiPce(TobPiece : TOB) : Boolean;
var TobArtFi : TOB;
begin
  Result := False;
  TobArtFi := FOGetPTobArtFiTiers;
  if TobArtFi <> nil then
  begin
    if IsArtFiByType(TobPiece, TobArtFi, 'SOD') then // Sortie diverses
      Result := True;
    if (not Result) and (IsArtFiByType(TobPiece, TobArtFi, 'FCA')) then // Fond de caisse
      Result := True;
    if (not Result) and (IsArtFiByType(TobPiece, TobArtFi, 'ECA')) then // Ecart de caisse
      Result := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/11/2002
Modifié le ... : 18/11/2002
Description .. : Retourne un option d'un code démarque
Suite ........ : Option = '1' -> Client obligatoire (non par défaut)
Suite ........ :        = '2' -> Imprimable en caisse (oui par défaut)
Suite ........ :        = '3' -> Pourcentage maximum ('' par défaut)
Suite ........ :        = '4' -> Date début d'utilisation (01/01/1900 par défaut)
Suite ........ :        = '5' -> Date fin d'utilisation (31/12/2099 par défaut)
Mots clefs ... : FO
*****************************************************************}

function FODecodeLibreDemarque(TOBDem: TOB; Option: string): string;
begin
  Result := '';
  if Option = '1' then Result := '-' else
    if Option = '2' then Result := 'X' else
    if Option = '4' then Result := DateToStr(IDate1900) else
    if Option = '5' then Result := DateToStr(IDate2099);
  if TOBDem = nil then Exit;
  // Client obligatoire (1 caractère - 'N' par défaut)
  if Option = '1' then
  begin
    Result := TOBDem.GetValue('GTR_CLIOBLIGFO');
    if Result <> 'X' then Result := '-';
  end else
    // Imprimable en caisse (1 caractère - 'O' par défaut)
    if Option = '2' then
  begin
    Result := TOBDem.GetValue('GTR_IMPRIMABLE');
    if Result <> '-' then Result := 'X';
  end else
    // Pourcentage maximum (3 caractères - '' par défaut)
    if Option = '3' then
  begin
    Result := IntToStr(TOBDem.GetValue('GTR_REMPOURMAX'));
  end else
    // Début de la période d'utilisation (8 caractères - '19000101' par défaut)
    if Option = '4' then
  begin
    Result := DateToStr(TOBDem.GetValue('GTR_DATEDEBUT'));
  end else
    // Fin de la période d'utilisation (8 caractères - '20991231' par défaut)
    if Option = '5' then
  begin
    Result := DateToStr(TOBDem.GetValue('GTR_DATEFIN'));
  end;
end;

{==============================================================================================}
{=================================== GESTION DES ARTICLES =====================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne l'identifiant de l'article générique à partir de celui
Suite ........ : d'un article dimensionné
Mots clefs ... : FO
*****************************************************************}

function FOIdentArticleGen(CodeArt: string): string;
var Nbc: Integer;
begin
  Result := '';
  if CodeArt = '' then Exit;
  if Length(CodeArt) > 18 then Nbc := 18 else Nbc := Length(CodeArt);
  Result := Copy(CodeArt, 1, Nbc);
  Nbc := (18 - Length(Result)) + (3 * MaxDimension);
  if Nbc < 0 then Nbc := 0;
  Result := Result + StringOfChar(' ', Nbc) + 'X';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 24/07/2001
Description .. : Indique si l'image de l'article doit être affichée en saisie de
Suite ........ : ticket pour la caisse courante
Mots clefs ... : FO
*****************************************************************}

function FOAfficheImageArt: Boolean;
begin
  Result := False;
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.GetValue('GPK_AFFPHOTOART') <> '') then
  begin
    Result := (VH_GC.TOBPCaisse.GetValue('GPK_AFFPHOTOART') = 'X');
  end;
end;

{==============================================================================================}
{================================= COMPOSANT TFOProgressForm ==================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  Create : création du composant
///////////////////////////////////////////////////////////////////////////////////////

constructor TFOProgressForm.Create(AOwner: TWinControl; Caption, Titre: string);
begin
  inherited Create(AOwner);
  InitMoveProgressForm(AOwner, Caption, Titre, 60, False, True);
  FTimer := TTimer.Create(Self);
  FTimer.OnTimer := IsTimer;
  FTimer.Enabled := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Destroy : destruction du composant
///////////////////////////////////////////////////////////////////////////////////////

destructor TFOProgressForm.Destroy;
begin
  FTimer.Free;
  FiniMoveProgressForm;
  inherited Destroy;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  IsTimer : événement OnTimer du composant
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOProgressForm.IsTimer(Sender: TObject);
begin
  MoveCurProgressForm('');
end;

{==============================================================================================}
{=================================== GESTION DES BOUTONS ======================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  FOEnableButtonFiche : Active ou inactive les boutons d'une fiche ( à voir ??)
///////////////////////////////////////////////////////////////////////////////////////

procedure FOEnableButtonFiche(Fiche: TForm; Enabled, ReInit: Boolean);
var Ind, Nb: Integer;
  Btn: TRadioButton;
begin
  for Ind := 0 to Fiche.Componentcount - 1 do
  begin
    if Fiche.Components[Ind] is TRadioButton then
    begin
      Btn := TRadioButton(Fiche.Components[Ind]);
      if ReInit then
      begin
        // Remise en état du mode
        if Btn.Tag >= 1024 then
        begin
          Nb := Btn.Tag;
          Dec(Nb, 1024);
          Btn.Tag := Nb;
          Btn.Enabled := not (Btn.Enabled);
        end;
      end else
      begin
        // Inversion du mode
        if Btn.Enabled <> Enabled then
        begin
          Nb := Btn.Tag;
          Inc(Nb, 1024);
          Btn.Tag := Nb;
          Btn.Enabled := Enabled;
        end;
      end;
    end;
  end;
end;

{==============================================================================================}
{================================= GESTION DE LA REGISTRY =====================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/06/2002
Modifié le ... : 28/06/2002
Description .. : Enregistre une valeur d'un champ en registry
Mots clefs ... : FO
*****************************************************************}

procedure FOSaveInRegistry(LocalKey, NomChamp: string; Valeur: Variant);
var Key: string;
begin
  Key := 'Software\' + Apalatys + '\' + NomHalley + '\' + LocalKey;
  SaveInRegistry(GetRootKey, Key, NomChamp, Valeur);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/06/2002
Modifié le ... : 28/06/2002
Description .. : Lit une valeur d'un champ en registry
Mots clefs ... : FO
*****************************************************************}

function FOGetFromRegistry(LocalKey, NomChamp: string; ValeurDefaut: Variant): Variant;
var Key: string;
begin
  Key := 'Software\' + Apalatys + '\' + NomHalley + '\' + LocalKey;
  Result := GetFromRegistry(GetRootKey, Key, NomChamp, ValeurDefaut);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/03/2003
Modifié le ... : 17/03/2003
Description .. : Supprime un champ en registry
Mots clefs ... : FO
*****************************************************************}

procedure FODeleteFromRegistry(LocalKey, NomChamp: string);
var Key: string;
  Reg: TRegIniFile;
begin
  Key := 'Software\' + Apalatys + '\' + NomHalley + '\' + LocalKey;
  Reg := TRegIniFile.Create('');
  try
    Reg.RootKey := GetRootKey;
    if Reg.OpenKey(Key, False) then
      Reg.DeleteKey('', NomChamp);
  finally
    Reg.Free;
  end;
end;

{==============================================================================================}
{================================ FONCTIONS MATHEMATIQUES =====================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/06/2002
Modifié le ... : 28/06/2002
Description .. : Change de base un nombre exprimé en base 10
Mots clefs ... : FO
*****************************************************************}

function FOChangeBase(Valeur: string; Base: Byte; Taille: Integer): string;
var Ind: Byte;
  iVal: Longint;
begin
  Result := '';
  iVal := StrToInt(Valeur);
  while iVal <> 0 do
  begin
    Ind := iVal mod Base;
    if Ind < 10 then Result := Chr(Ord('0') + Ind) + Result
    else Result := Chr(Ord('A') + Ind - 10) + Result;
    iVal := iVal div Base;
  end;
  Result := StringOfChar('0', Taille - Length(Result)) + Result;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/06/2002
Modifié le ... : 19/09/2003
Description .. : Ajout un caractère alphabétique à un code numérique.
Suite ........ : Par exemple 001 => A01 ou 123 => B23  ou 999 => J99
Mots clefs ... : FO
*****************************************************************}

function FOAlphaCodeNumeric(const Code: string): string;
begin
  Result := Code;
  if Length(Code) > 0 then Result[1] := Chr(Ord('A') + Ord(Code[1]) - Ord('0'));
end;

{==============================================================================================}
{==================================== CONFIDENTIALITE =========================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/03/2003
Modifié le ... : 12/03/2003
Description .. : Decode les droits d'accès d'une ligne de menu pour un
Suite ........ : groupe d'utilisateurs.
Mots clefs ... : FO
*****************************************************************}

function FODecodeAccesGrp(AccesGrp: string; UserGrp: Integer): boolean;
var NivConcept: byte;
begin
  NivConcept := 0;
  if (UserGrp > 0) and (UserGrp <= Length(AccesGrp)) then
    case AccesGrp[UserGrp] of
      'X': NivConcept := 1;
      '-': NivConcept := 2;
    end;
  Result := (NivConcept <= 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/02/2003
Modifié le ... : 12/03/2003
Description .. : Vérifie si un concept est autorisé pour un groupe
Suite ........ : d'utilisateurs.
Mots clefs ... : FO
*****************************************************************}

function FOJaiLeDroitConcept(cc: TConcept; UserGrp: Integer = 0): boolean;
var QQ: TQuery;
  sSql, sTag, sAcces: string;
begin
  if (UserGrp = 0) or (UserGrp = V_PGI.UserGrp) then
  begin
    Result := ExJaiLeDroitConcept(cc, False);
    Exit;
  end;
  // recherche des droits sur le concept pour le groupe indiqué
  sAcces := '';
  sTag := IntToStr(Ord(cc) + 26000);
  sSql := 'SELECT MN_LIBELLE,MN_ACCESGRP FROM MENU WHERE MN_1=26 AND MN_TAG=' + sTag;
  QQ := OpenSQL(sSql, False);
  if not QQ.Eof then
  begin
    if QQ.FindField('MN_LIBELLE').AsString <> '-' then
      sAcces := QQ.FindField('MN_ACCESGRP').AsString;
  end;
  Ferme(QQ);
  Result := FODecodeAccesGrp(sAcces, UserGrp);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/03/2003
Modifié le ... : 12/03/2003
Description .. : Retourne la TOB des événements de confidentialité
Mots clefs ... : FO
*****************************************************************}

function FOGetPTobConfidentialite: TOB;
var Ind: Integer;
begin
  Result := nil;
  if VH_GC.TOBPCaisse = nil then Exit;
  for Ind := 0 to VH_GC.TOBPCaisse.Detail.Count - 1 do
  begin
    if (VH_GC.TOBPCaisse.Detail[Ind].NomTable = '') and
      (VH_GC.TOBPCaisse.Detail[Ind].Commentaire = 'Confidentialité') then
    begin
      Result := VH_GC.TOBPCaisse.Detail[Ind];
      Break;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/03/2003
Modifié le ... : 12/03/2003
Description .. : Charge en TOB les événements de confidentialité dont le
Suite ........ : tag est compris entre 107901 et 107999.
Mots clefs ... : FO
*****************************************************************}

procedure FOChargeConfidentialite(NoMenu: integer);
var TOBEvt, TOBAcc, TOBL: TOB;
  sSql, sAcces, TabAcces: string;
  QQ: TQuery;
  Ind, Tag: integer;
begin
  if VH_GC.TOBPCaisse = nil then Exit;
  // Libération de l'ancienne TOB si besoin
  TOBEvt := FOGetPTobConfidentialite;
  if TOBEvt <> nil then
  begin
    if (TOBEvt.GetValue('NOMENU') = NoMenu) and (TOBEvt.GetValue('USERGRP') = V_PGI.UserGrp) then
      Exit // les données sont encores valides, rechargement inutiles
    else
      TOBEvt.Free;
  end;
  TabAcces := StringOfChar('X', 99);
  // Lecture des événements (tags du menu 107 compris entre 107901 et 107999)
  TOBAcc := TOB.Create('', nil, -1);
  {$IFDEF GESCOM}
  sSql := 'SELECT MN_TAG,MN_LIBELLE,MN_ACCESGRP FROM MENU WHERE MN_1=' + IntToStr(NoMenu)
    + ' AND MN_TAG>=' + IntToStr(NoMenu) + '601'
    + ' AND MN_TAG<=' + IntToStr(NoMenu) + '699';
  {$ELSE}
  sSql := 'SELECT MN_TAG,MN_LIBELLE,MN_ACCESGRP FROM MENU WHERE MN_1=' + IntToStr(NoMenu)
    + ' AND MN_TAG>=' + IntToStr(NoMenu) + '901'
    + ' AND MN_TAG<=' + IntToStr(NoMenu) + '999';
  {$ENDIF}
  QQ := OpenSQL(sSql, False);
  if not QQ.Eof then TOBAcc.LoadDetailDB('MENU', '', '', QQ, False);
  Ferme(QQ);
  // Mise sous la forme d'une chaîne de caratères/booléens des droits d'accès du groupe
  for Ind := 0 to TOBAcc.Detail.Count - 1 do
  begin
    TOBL := TOBAcc.Detail[Ind];
    if TOBL.GetValue('MN_LIBELLE') <> '-' then
    begin
     {$IFDEF GESCOM}
      Tag := TOBL.GetValue('MN_TAG') - ((NoMenu * 1000) + 600);
     {$ELSE}
      Tag := TOBL.GetValue('MN_TAG') - ((NoMenu * 1000) + 900);
     {$ENDIF}
      if (Tag >= 1) or (Tag <= 99) then
      begin
        sAcces := TOBL.GetValue('MN_ACCESGRP');
        if not FODecodeAccesGrp(sAcces, V_PGI.UserGrp) then TabAcces[Tag] := '-';
      end;
    end;
  end;
  TOBAcc.Free;
  // Stockage des droits d'accès une TOB fille de la TOB du paramétrage de la caisse
  TOBEvt := TOB.Create('', VH_GC.TOBPCaisse, -1);
  TOBEvt.Commentaire := 'Confidentialité';
  TOBEvt.AddChampSupValeur('NOMENU', NoMenu);
  TOBEvt.AddChampSupValeur('USERGRP', V_PGI.UserGrp);
  TOBEvt.AddChampSupValeur('TABACCES', TabAcces);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/03/2003
Modifié le ... : 12/03/2003
Description .. : Convertir un code événement en numéro de tag ou numéro
Suite ........ : de concept.
Mots clefs ... : FO
*****************************************************************}

function FOEvenementToTag(CodeEvenement: integer; var Tag, NoConcept: integer): boolean;
begin
  Result := True;
  Tag := 0;
  NoConcept := 0;
  case CodeEvenement of
  {$IFDEF GESCOM}
    1: Tag := 36601; // Commentaire d'en-tête
    2: Tag := 36602; // Commentaire de pied
    3: Tag := 36603; // Complément d'en-tête
    4: Tag := 36604; // Saisir le régime de taxe
    5: Tag := 36605; // Saisir le code de taxe
    6: Tag := 36606; // Demande de détaxe
    7: Tag := 36607; // Complément de la ligne
    8: Tag := 36608; // Saisir une ligne de retour
    9: Tag := 36609; // Descriptif détaillé de l'article
    10: Tag := 36610; // Modification du dépôt de la ligne
    11: Tag := 36611; // Reprise d'un ticket en attente
    12: Tag := 36612; // Voir l'article
    13: Tag := 36613; // Voir le client
    14: Tag := 36614; // Voir le tarif
    15: Tag := 36615; // Voir le vendeur
    16: Tag := 36616; // Voir la devise
    17: Tag := 36617; // Détail de la nomenclature
    18: Tag := 36618; // Voir le stock de l'article
    19: Tag := 36619; // Voir le stock
    20: Tag := 36620; // Voir le stock distant
    21: Tag := 36621; // Réimpression du dernier ticket
    22: Tag := 36622; // Choix du modèle d'impression
    23: Tag := 36623; // Ouverture du tiroir caisse
    24: Tag := 36624; // Abandon de la pièce
    25: Tag := 36625; // Suppression d'une ligne
    26: Tag := 36626; // Modification de la référence saisie
    27: Tag := 36627; // Modification de la quantité
    28: Tag := 36628; // Modification du libellé
    29: Tag := 36629; // Modification du code démarque
    30: Tag := 36630; // Remise supérieure au seuil n°1
    31: Tag := 36631; // Remise supérieure au seuil n°2
    32: Tag := 36632; // Remise supérieure au seuil n°3
    33: Tag := 36633; // Outrepasser le maximum du motif de démarque
    34: Tag := 36634; // Majoration du prix
    35: Tag := 36635; // Modifier la fiche du client
    36: Tag := 36636; // Modification du vendeur  ?? NoConcept := gcSaisModifRepres ??
    37: Tag := 36637; // Mettre un ticket en attente
    38: Tag := 36638; // Rechercher un article
    39: Tag := 36639; // Réimprimer les bons du dernier ticket

    51: Tag := 36651; // Voir le mode de paiement
    52: Tag := 36652; // Détail de l'échéance
    53: Tag := 36653; // Solder la ligne d'échéance
    54: Tag := 36654; // Remise globale sur le ticket
    55: Tag := 36655; // Outrepasser les conditions sur les montants
    56: Tag := 36656; // Forcer le résultat du TPE
    57: Tag := 36657; // Outrepasser la liaison des règlements

    61: Tag := 36661; // Acquisition bon achat
    62: Tag := 36662; // Encaissement chèque différé
    63: Tag := 36663; // Encaissement crédit
    64: Tag := 36664; // Entrée diverse
    65: Tag := 36665; // Fond de caisse
    66: Tag := 36666; // Prélèvement
    67: Tag := 36667; // Remboursement arrhes
    68: Tag := 36668; // Remboursement avoir
    69: Tag := 36669; // Remboursement bon achat
    70: Tag := 36670; // Sortie Diverse
    71: Tag := 36671; // Versement arrhes
    72: Tag := 36672; // Espèce
    73: Tag := 36673; // Avoir
    74: Tag := 36674; // Chèque
    75: Tag := 36675; // Chèque différé
    76: Tag := 36676; // Carte bancaire
    77: Tag := 36677; // Arrhes déjà versées
    78: Tag := 36678; // Reste dû
    79: Tag := 36679; // Bon d'achat
    80: Tag := 36680; // Autre mode de règlement
    81: Tag := 36681; // Saisir des documents si la journée ne correspond pas à la date système
    82: Tag := 36682; // Modifier les actions complémentaires de la fermeture
    83: Tag := 36683; // Paramétrage de la Situation Flash
    84: Tag := 36684; // Affichage du détail des tickets en contrôle caisse
    85: Tag := 36685; // Affichage du solde du client dans tous les établissements
    86: Tag := 36686; // Conserver les montants saisis dans le contrôle caisse
    87: Tag := 36687; // Corriger le cumul de fidélité

    // à partir de 100 les événements ne sont pas chargés en TOB
    101: Tag := 36301; // Ticket X
    102: Tag := 36302; // Ticket Z
    103: Tag := 36401; // Situation Flash
    104: NoConcept := gcCliCreat; // Création clients
  {$ELSE}
    1: Tag := 107901; // Commentaire d'en-tête
    2: Tag := 107902; // Commentaire de pied
    3: Tag := 107903; // Complément d'en-tête
    4: Tag := 107904; // Saisir le régime de taxe
    5: Tag := 107905; // Saisir le code de taxe
    6: Tag := 107906; // Demande de détaxe
    7: Tag := 107907; // Complément de la ligne
    8: Tag := 107908; // Saisir une ligne de retour
    9: Tag := 107909; // Descriptif détaillé de l'article
    10: Tag := 107910; // Modification du dépôt de la ligne
    11: Tag := 107911; // Reprise d'un ticket en attente
    12: Tag := 107912; // Voir l'article
    13: Tag := 107913; // Voir le client
    14: Tag := 107914; // Voir le tarif
    15: Tag := 107915; // Voir le vendeur
    16: Tag := 107916; // Voir la devise
    17: Tag := 107917; // Détail de la nomenclature
    18: Tag := 107918; // Voir le stock de l'article
    19: Tag := 107919; // Voir le stock
    20: Tag := 107920; // Voir le stock distant
    21: Tag := 107921; // Réimpression du dernier ticket
    22: Tag := 107922; // Choix du modèle d'impression
    23: Tag := 107923; // Ouverture du tiroir caisse
    24: Tag := 107924; // Abandon de la pièce
    25: Tag := 107925; // Suppression d'une ligne
    26: Tag := 107926; // Modification de la référence saisie
    27: Tag := 107927; // Modification de la quantité
    28: Tag := 107928; // Modification du libellé
    29: Tag := 107929; // Modification du code démarque
    30: Tag := 107930; // Remise supérieure au seuil n°1
    31: Tag := 107931; // Remise supérieure au seuil n°2
    32: Tag := 107932; // Remise supérieure au seuil n°3
    33: Tag := 107933; // Outrepasser le maximum du motif de démarque
    34: Tag := 107934; // Majoration du prix
    35: Tag := 107935; // Modifier la fiche du client
    36: Tag := 107936; // Modification du vendeur  ?? NoConcept := gcSaisModifRepres ??
    37: Tag := 107937; // Mettre un ticket en attente
    38: Tag := 107938; // Rechercher un article
    39: Tag := 107939; // Réimprimer les bons du dernier ticket

    51: Tag := 107951; // Voir le mode de paiement
    52: Tag := 107952; // Détail de l'échéance
    53: Tag := 107953; // Solder la ligne d'échéance
    54: Tag := 107954; // Remise globale sur le ticket
    55: Tag := 107955; // Outrepasser les conditions sur les montants
    56: Tag := 107956; // Forcer le résultat du TPE
    57: Tag := 107957; // Outrepasser la liaison des règlements

    61: Tag := 107961; // Acquisition bon achat
    62: Tag := 107962; // Encaissement chèque différé
    63: Tag := 107963; // Encaissement crédit
    64: Tag := 107964; // Entrée diverse
    65: Tag := 107965; // Fond de caisse
    66: Tag := 107966; // Prélèvement
    67: Tag := 107967; // Remboursement arrhes
    68: Tag := 107968; // Remboursement avoir
    69: Tag := 107969; // Remboursement bon achat
    70: Tag := 107970; // Sortie Diverse
    71: Tag := 107971; // Versement arrhes
    72: Tag := 107972; // Espèce
    73: Tag := 107973; // Avoir
    74: Tag := 107974; // Chèque
    75: Tag := 107975; // Chèque différé
    76: Tag := 107976; // Carte bancaire
    77: Tag := 107977; // Arrhes déjà versées
    78: Tag := 107978; // Reste dû
    79: Tag := 107979; // Bon d'achat
    80: Tag := 107980; // Autre mode de règlement
    81: Tag := 107981; // Saisir des documents si la journée ne correspond pas à la date système
    82: Tag := 107982; // Modifier les actions complémentaires de la fermeture
    83: Tag := 107983; // Paramétrage de la Situation Flash
    84: Tag := 107984; // Affichage du détail des tickets en contrôle caisse
    85: Tag := 107985; // Affichage du solde du client dans tous les établissements
    86: Tag := 107986; // Conserver les montants saisis dans le contrôle caisse
    87: Tag := 107987; // Corriger le cumul de fidélité

    // à partir de 100 les événements ne sont pas chargés en TOB
    101: Tag := 107310; // Ticket X
    102: Tag := 107320; // Ticket Z
    103: Tag := 107410; // Situation Flash
    104: NoConcept := gcCliCreat; // Création clients
  {$ENDIF}
  else Result := False;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/03/2003
Modifié le ... : 12/03/2003
Description .. : Vérifie la confidentialté pour une code événement
Suite ........ : de concept.
Mots clefs ... : FO
*****************************************************************}

function FOEvenementConfidentiel(CodeEvenement: integer; UserGrp: integer): boolean;
var Tag, NoConcept, NoMenu, Ind: integer;
  TabAcces: string;
  TOBEvt: TOB;
begin
  Result := True;
  if UserGrp = 0 then UserGrp := V_PGI.UserGrp;
  if not FOEvenementToTag(CodeEvenement, Tag, NoConcept) then
  begin
    if V_PGI.SAV then PGIBox(IntToStr(CodeEvenement), 'Code événement inconnu');
    Exit;
  end;
  // recherche dans la TOB de la confidentialité
  TOBEvt := FOGetPTobConfidentialite;
  if TOBEvt <> nil then
  begin
    NoMenu := TOBEvt.GetValue('NOMENU');
   {$IFDEF GESCOM}
    if (Tag >= (NoMenu * 1000) + 601) and (Tag <= (NoMenu * 1000) + 699) and
      (UserGrp = TOBEvt.GetValue('USERGRP')) then
   {$ELSE}
    if (Tag >= (NoMenu * 1000) + 901) and (Tag <= (NoMenu * 1000) + 999) and
      (UserGrp = TOBEvt.GetValue('USERGRP')) then
   {$ENDIF}
    begin
      TabAcces := TOBEvt.GetValue('TABACCES');
     {$IFDEF GESCOM}
      Ind := Tag - ((NoMenu * 1000) + 600);
     {$ELSE}
      Ind := Tag - ((NoMenu * 1000) + 900);
     {$ENDIF}
      if (Ind >= 1) and (Ind <= Length(TabAcces)) then
      begin
        Result := (TabAcces[Ind] <> '-');
        Exit;
      end;
    end;
  end;
  // recherche dans la table des menus ou dans les concepts
  if Tag <> 0 then
    Result := JaiLeDroitTag(Tag, UserGrp)
  else if NoConcept > 0 then
    Result := FOJaiLeDroitConcept(TConcept(NoConcept), UserGrp)
  else
    Result := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/02/2003
Modifié le ... : 12/03/2003
Description .. : Vérifie si un événement est autorisé à un groupe
Suite ........ : d'utilisateurs.
Suite ........ : Un événement correspond soit à un concept, soit à une
Suite ........ : ligne de menu.
Mots clefs ... : FO
*****************************************************************}

function FOJaiLeDroit(CodeEvenement: integer; Parle: boolean = True; ChgUser: boolean = True; ErrorMsg: string = ''; UserGrp: integer = 0): boolean;
var Stg, Titre, CodeGrp: string;
  ZoomOLE: boolean;
begin
  Result := FOEvenementConfidentiel(CodeEvenement, UserGrp);
  if (Parle) and (not Result) then
  begin
    if ErrorMsg = '' then
      Stg := TraduireMemoire('Vous n''avez pas le droit d''effectuer cette opération.')
    else
      Stg := TraduireMemoire(ErrorMsg);
    Titre := TraduireMemoire('Action interdite !');
    if ChgUser then
    begin
      Stg := Stg + #10#10 + TraduireMemoire('Voulez-vous saisir un mot de passe pour continuer ?');
      if PGIAsk(Stg, Titre) = mrYes then
      begin
        ZoomOLE := V_PGI.ZoomOLE;
        V_PGI.ZoomOLE := True;
        CodeGrp := AGLLanceFiche('MFO', 'MFOUSERCHG', '', '', '');
        V_PGI.ZoomOLE := ZoomOLE;
        if CodeGrp <> '' then
          Result := FOJaiLeDroit(CodeEvenement, False, False, '', StrToInt(CodeGrp));
      end;
    end else
      PGIBox(Stg, Titre);
  end;
end;

{==============================================================================================}
{==================================== PAVÉ TACTILE ============================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : X. MALUENDA
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. :
Mots clefs ... : FO
*****************************************************************}

function CountTokenPipe(LargeSt, Sep: string): Integer;
var St: string;
begin
  Result := 0;
  Sep := Trim(Sep); //xmg 8/6/00
  LargeSt := Trim(LargeSt);
  if (Length(LargeSt) > 0) and (Sright(LargeST, length(Sep)) <> sep) then LargeSt := LargeSt + Sep; ///XMG 8/6/00
  while LargeSt <> '' do
  begin
    St := ReadTokenPipe(LargeSt, Sep);
    inc(Result, ord(St <> ''));
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : X. MALUENDA
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. :
Mots clefs ... : FO
*****************************************************************}

function KeytoChar(Key: Word; Shift: TShiftState): Char;
var Kb_etat: TKeyboardState;
  Scan: Cardinal;
  KL: HKL;
  St: string;
begin
  Result := #0;
  St := format_String(#0, 255);
  if ((ssCtrl in Shift) or (ssAlt in Shift)) and ((not (ssShift in Shift)) and ([ssctrl, ssAlt] <> Shift)) then
  begin
    Result := chr(Key);
    exit;
  end;
  GetKeyboardState(KB_etat);
  KL := Getkeyboardlayout(0);
  Scan := mapVirtualKeyEx(Key, 0, kl);
  if ToAscIIex(Key, scan, KB_Etat, pchar(St), 0, kl) <> 1 then St := #0;
  if length(St) > 0 then Result := St[1];
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : X. MALUENDA
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. :
Mots clefs ... : FO
*****************************************************************}

procedure DisableControlsFiche(F: TCustomForm; FCaisse: TWinControl; Disable, OkEnableCaisse: Boolean; OkCaisse: Boolean = FALSE);
  function OkCOntrol(X: TComponent): Boolean;
  begin
    Result := TRUE;
    if (x is THValComboBox) then exit;
    if (x is THCritMaskEdit) then exit;
    if (x is THGrid) then exit;
    if (x is THNumEdit) then exit;
    if (x is TCheckBox) then exit;
    if (x is TPageControl) then exit;
    if (x is TEdit) then exit;
    if (x is TRadioButton) then exit;
    //  if  (x is ) then exit ;
    Result := FALSE;
  end;

var i: Integer;
  C: TControl;
  x: TComponent;
begin
  for i := 0 to F.Componentcount - 1 do
  begin
    X := F.Components[i];
    if OkControl(X) then //(X is TControl) and (not (X is THLabel)) then
    begin
      C := TControl(F.Components[i]);
      if Disable then
      begin
        if C <> FCaisse then C.Enabled := FALSE
        else C.Enabled := TRUE;
      end else
      begin
        if (OkENableCaisse) and (C = FCaisse) then C.Enabled := FALSE
        else C.Enabled := TRUE;
      end;
    end;
  end;
  if (Disable) and (FCaisse <> nil) then
  begin
    C := FCaisse.Parent;
    while C <> nil do
    begin
      if not C.Enabled then C.Enabled := TRUE;
      C := C.Parent;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : X. MALUENDA
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. :
Mots clefs ... : FO
*****************************************************************}

function AbsoluteGoForward(Ctrl1, Ctrl2: TWinControl): Boolean;
var Prnt1, prnt2, oPrnt1, oprnt2: TWinControl;
  Form: TCustomForm;
  ok: Boolean;
begin
  Result := TRUE;
  if (Ctrl1 = nil) or (Ctrl2 = nil) then exit;
  Form := GetParentform(Ctrl1);
  if Form <> GetParentForm(Ctrl2) then exit;
  Prnt1 := Ctrl1.parent;
  Prnt2 := nil;
  oPrnt1 := nil;
  oPrnt2 := nil;
  while (Prnt1 <> Prnt2) do
  begin
    Prnt2 := Ctrl2;
    ok := FALSE;
    while (not Ok) and (prnt2 <> Form) do
    begin
      oPrnt2 := Prnt2;
      Prnt2 := prnt2.parent;
      ok := (Prnt2 = Prnt1);
    end;
    if not ok then
    begin
      oPrnt1 := prnt1;
      Prnt1 := Prnt1.parent;
    end;
  end;
  if prnt1 <> ctrl1.parent then
  begin
    Prnt1 := oprnt1;
    prnt2 := oprnt2;
  end else
    if prnt1 = prnt2 then
  begin
    Prnt1 := Ctrl1;
    prnt2 := Ctrl2;
  end;
  Result := Prnt1.tabOrder < Prnt2.TabOrder;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : X. MALUENDA
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : -TPGIErr-Erreur
Mots clefs ... : FO
*****************************************************************}

function MsgErrDefautS1(Code: integer): string;
begin
  // liste des codes erreurs
  case Code of
    0: result := 'Une erreur non référencée s''est produite.';
    1: result := 'L''enregistrement est impossible !';
    2: result := 'L''enregistrement est introuvable !';
    3: result := 'Une information obligatoire est absente';
    4: result := 'Une information est incohérente';
    5: result := 'Fichier absent';
    6: result := 'Erreur matérielle';
    7: result := 'Enregistrement interdit';
    8: result := 'L''enregistrement a été modifié par un autre utilisateur.';
    9: result := 'Le code existe déjà';
    10: result := 'L''importation a été interrompue par l''utilisateur.';
    11: result := 'L''opération n''est pas permise.';
    710: Result := 'Voulez-vous enregistrer les modifications ?';
    744: Result := 'Confirmez-vous la sortie du paramétrage ?';
    752: Result := 'Ce raccourci clavier est déjà utilisé.';
    756: Result := 'Désirez-vous recréer le pavé '; //le ? est mis par le code
    768: Result := 'Cet enregistrement n''est pas modifiable.';
    787: Result := 'Il y a déjà des boutons sur les autres pages à cette position.';
    788: Result := 'Il existe déjà des boutons à la position choisie.' + #13 + ' Désirez-vous les intervertir ?';
    789: Result := 'Plusieurs boutons ont été déplacés du fait de l''existence de boutons fixes.';
    790: Result := 'Plusieurs boutons ont été déplacés du fait de l''existence d''un bouton fixe.';

  else Result := 'Erreur n°' + IntToStr(Code);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : X. MALUENDA
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. :
Mots clefs ... : FO
*****************************************************************}

{$IFNDEF EAGLCLIENT}
function OpenSQLDB(SQL: string; DB: TDatabase): TQuery; //Copie des HCtrl.OpenSQL
begin
  Result := DBOpenSQL(DB, SQL, StToDriver(DB.DriverName, FALSE));
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/11/2003
Modifié le ... : 14/11/2003
Description .. : vérifie la cohérence des boutons 'vendeur' du pavé pour
Suite ........ : une liste de caisse
Mots clefs ... : 
*****************************************************************}

function FOVerifieVendeurPave (Etablis: string; Caisse: array of string; RemplaceAuto: boolean): boolean;
var
  Ind, IndVen: Integer;
  sSql, sSep, Stg, sConcept, sCode, sNumCais: string;
  TOBBtns, TOBb, TOBVen, TOBv: TOB;
  Remplace: boolean;
begin
  Result := True;
  Remplace := False;
  sNumCais := '';
  // chargement des boutons 'Vendeur'
  sSql := 'SELECT * FROM CLAVIERECRAN WHERE CE_CAISSE IN (';
  for Ind := Low(Caisse) to High(Caisse) do
  begin
    if Ind = Low(Caisse) then sSep := '' else sSep := ',';
    sSql := sSql + sSep + '"' + Caisse[Ind] + '"';
  end;
  sSql := sSql + ') AND CE_TEXTE LIKE "VEN%" ORDER BY CE_CAISSE';
  TOBBtns := TOB.Create('', nil, -1);
  TOBBtns.LoadDetailDBFromSQL('CLAVIERECRAN', sSql);
  if TOBBtns.Detail.Count > 0 then
  begin
    // chargement des vendeurs de l'établissement
    sSql := 'SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE GCL_TYPECOMMERCIAL="VEN" '
      + 'AND GCL_ETABLISSEMENT="' + Etablis + '" AND GCL_DATESUPP>"' + USDateTime(Date) + '"';
    TOBVen := TOB.Create('', nil, -1);
    TOBVen.LoadDetailFromSQL(sSql);
    IndVen := 0;

    for Ind := 0 to TOBBtns.Detail.Count -1 do
    begin
      TOBb := TOBBtns.Detail[Ind];
      Stg := vString(TOBb.GetValue('CE_TEXTE'));
      sConcept := ReadTokenSt(Stg);
      sCode := ReadTokenSt(Stg);
      if (sConcept = 'VEN') and (sCode <> '') then
      begin
        TOBv := TOBVen.FindFirst(['GCL_COMMERCIAL'], [sCode], False);
        if TOBv = nil then
        begin
          if (not Remplace) and (RemplaceAuto) then
          begin
            Stg := TraduireMemoire('Des boutons font référence à des vendeurs d''un autre établissement.')
              + #10#10 + TraduireMemoire('Voulez-vous les remplacer par des vendeurs de l''établisssement ?');
            if PGIAsk (Stg) = mrYes then Remplace := True;
          end;

          if Remplace then
          begin
            if sNumCais <> TOBb.GetValue('CE_CAISSE') then
            begin
              // changement de caisse on repart depuis le début de la liste de vendeur
              IndVen := 0;
              sNumCais := TOBb.GetValue('CE_CAISSE');
            end;

            if IndVen < TOBVen.Detail.Count then
            begin
              // remplacement du vendeur du bouton
              TOBv := TOBVen.Detail[IndVen];
              Stg := 'VEN;' + TOBv.GetValue('GCL_COMMERCIAL') + ';';
              TOBb.PutValue('CE_TEXTE', Stg);
              TOBb.PutValue('CE_CAPTION', #255+#255+'LIB');
              TOBb.SetAllModifie(True);
              TOBBtns.Modifie := True;
              Inc(IndVen);
            end else
            begin
              // suppression du bouton
              TOBb.DeleteDB;
              TOBb.SetAllModifie(False);
            end;
          end;
        end;
        if (TOBv = nil) and (not Remplace) then
        begin
          Result := False;
          break;
        end;
      end;
    end;
    TOBVen.Free;
  end;
  if (Remplace) and (TOBBtns.Modifie) then TOBBtns.UpdateDB;
  TOBBtns.Free;
end;

end.
