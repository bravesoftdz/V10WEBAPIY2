{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/05/2002
Modifié le ... :   /  /    
Description .. : Saisie des tables de correspondances pour l'importation des 
Suite ........ : données GPAO dans BO MODE
Mots clefs ... : 
*****************************************************************}
unit SaisieCorrespGPAO;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, HTB97, StdCtrls, Mask, ExtCtrls, HPanel, UIUtil, HSysMenu,
  LooKup, HMsgBox, HEnt1, SAISUTIL, UTOB,
{$IFDEF EAGLCLIENT}
  Math;
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Math, PrintDBG;
{$ENDIF}


type
  TFCorrespGPAO = class(TForm)
    PENTETE: TPanel;
    LGRO_TYPEREPRISE: TLabel;
    LibNomTable: TLabel;
    GRO_TYPEREPRISE: THValComboBox;
    GRO_NOMCHAMP: THCritMaskEdit;
    Panel4: TPanel;
    GS: THGrid;
    LGRO_NOMCHAMP: TLabel;
    HMTrad: THSystemMenu;
    FindLigne: TFindDialog;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BChercher: TToolbarButton97;
    bNewligne: TToolbarButton97;
    bDelLigne: TToolbarButton97;
    BtestGPAOenPGI: TToolbarButton97;
    BtestPGIenGPAO: TToolbarButton97;
    VALDEFAUTGPAO: THCritMaskEdit;
    VALDEFAUTPGI1: THCritMaskEdit;
    VALDEFAUTPGI5: THCritMaskEdit;
    VALDEFAUTPGI2: THCritMaskEdit;
    VALDEFAUTPGI3: THCritMaskEdit;
    VALDEFAUTPGI4: THCritMaskEdit;
    LVALDEFAUTGPAO: TLabel;
    LVALDEFAUTPGI: TLabel;
    VALDEFAUTPGI6: THCritMaskEdit;
    VALDEFAUTPGI7: THCritMaskEdit;
    VALDEFAUTPGI8: THCritMaskEdit;
    TWTest: TToolWindow97;
    HPanel5: THPanel;
    BValideTest: TToolbarButton97;
    BFermeTest: TToolbarButton97;
    LVALTESTGPAO: TLabel;
    VALTESTGPAO: THCritMaskEdit;
    LVALTESTPGI1: TLabel;
    VALTESTPGI1: THCritMaskEdit;
    VALTESTPGI2: THCritMaskEdit;
    VALTESTPGI3: THCritMaskEdit;
    VALTESTPGI4: THCritMaskEdit;
    VALTESTPGI8: THCritMaskEdit;
    VALTESTPGI7: THCritMaskEdit;
    VALTESTPGI6: THCritMaskEdit;
    VALTESTPGI5: THCritMaskEdit;
    TVALTESTGPAO: THLabel;
    BGenere: TToolbarButton97;
    BVoirTob: TToolbarButton97;
    BImprime: TToolbarButton97;
    FAutoSave: TCheckBox;
    LIBDEFAUTGPAO: THCritMaskEdit;
    procedure GSElipsisClick(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure bDelLigneClick(Sender: TObject);
    procedure bNewligneClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BFermeTestClick(Sender: TObject);
    procedure BValideTestClick(Sender: TObject);
    procedure BtestGPAOenPGIClick(Sender: TObject);
    procedure BtestPGIenGPAOClick(Sender: TObject);
    procedure TWTestClose(Sender: TObject);
    procedure BGenereClick(Sender: TObject);
    procedure BVoirTobClick(Sender: TObject);
    procedure BImprimeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure VALDEFAUTChange(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Déclarations privées }
    Action       : TActionFiche ;            // action de saisie de la fiche
    LookUpMode   : Array[1..8] of Boolean ;  // type de présentation (LookUp ou Combo) de la colonne de la grille
    NomTable     : Array[1..8] of String ;   // nom de la table associée à la colonne de la grille
    NbCol        : Integer ;                 // nombre de colonnes actives correspondant au nombre de champs PGI
    ColDebut     : Integer ;                 // 1ère colonne de saisie des valeurs PGI
    COL_CodeGpao : Integer ;                 // Rang de la colonne de saisie
    StCellCur    : String ;                  // valeur avant modification de la cellule courante
    FindFirst    : Boolean ;                 // indicateur pour la recherche de texte
    Modifie      : Boolean ;                 // indicateur de la saisie d'une modification
    TOB_Corr     : TOB ;                     // TOB de la table de correspondance
    procedure CompleteTOBCorr ;
    procedure InsereTOBCorr ( ARow : Integer ; SiBesoin : Boolean = False ) ;
    procedure InitTOBCorr ( TOBLigne : TOB ) ;
    procedure DepileTOBCorr ( TOBPrinc : TOB ) ;
    procedure AjoutValeurDefaut ( TOBPrinc : TOB ) ;
    function  GetTobCorr ( ARow : Integer ) : TOB ;
    procedure EnregistreSaisie ;
    procedure ChargeTableCorrespondance ;
    procedure AffichageTableCorrespondance ;
    procedure ChargeGrille ( ARow : Integer ) ;
    procedure BrancheTablette ( Var Rang : Integer ; Tablette : String ) ;
    procedure AdapteGrille ;
    procedure ActiveBoutons ( Actif : Boolean ) ;
    procedure ActiveChampTest ( GPAOenPGI : Boolean ) ;
    function  ValideSaisie : Boolean ;
    procedure GenereValeurs ( Champs : Array of String ; Valeurs : Array of Variant ; Rang : Integer ; Var First : Boolean ; Var NbLig, ARow : Integer ) ;
  public
    { Déclarations publiques }
  end;

Const TypeRepriseGB2000 = 'GB'  ;       // Reprise des données de GB 2000
      TypeRepriseParam  = 'GBP' ;       // Paramétrage de la repise de données GB 2000
      TypeRepriseCASH   = 'CAS' ;       // Reprise des données de CASH 2000

Function  TraduitGPAOenPGI ( TypeReprise, NomChamp, Valeur : String ; TOBPrinc : TOB = Nil ; ToutPreCharge : Boolean = False ) : String ;
Function  TraduitPGIenGPAO ( TypeReprise, NomChamp : String ; Valeurs : Array of String ; AvecLibelle : Boolean = False ; TOBPrinc : TOB = Nil ; ToutPreCharge : Boolean = False ) : String ;
Procedure AppelSaisieCorrespGPAO ( Action : TActionFiche ; TypeReprise, NomChamp : string ) ;

implementation

{$R *.DFM}

///////////////////////////////////////////////////////////////////////////////
// Définition des champs spécifiques
///////////////////////////////////////////////////////////////////////////////
Const
  ChampsSpecifiques : array[1..4,1..9] of String = (
    ('$$_CODETVA'  , 'TTREGIMETVA'        , 'GCFAMILLETAXE1', '', '', '', '', '', '' ),
    ('$$_OPCAISSE' , 'GCARTICLEFINANCIER' , 'GCMODEPAIE', '', '', '', '', '', '' ),
    ('$$_CODEFAMSF', 'GCFAMILLENIV1'      , 'GCFAMILLENIV2' , 'GCFAMILLENIV3', 'GCFAMILLENIV4', 'GCFAMILLENIV5', 'GCFAMILLENIV6', 'GCFAMILLENIV7', 'GCFAMILLENIV8' ),
    (''            , ''                   , ''              , '', '', '', '', '', '' )) ;

///////////////////////////////////////////////////////////////////////////////
// Définition de la table des synonymes
///////////////////////////////////////////////////////////////////////////////
Const
  ChampsSynonymes : array[1..3,1..2] of String = (
    ('DEVISEFO'     , 'DEVISE'   ),
    ('DEVISEESP'    , 'DEVISE'   ),
    (''             , ''         )) ;

///////////////////////////////////////////////////////////////////////////////
// Définition des diverses constantes
///////////////////////////////////////////////////////////////////////////////
Const NbRowsInit     = 15 ;     // nombre de lignes lors de la création de la grille
      NbRowsPlus     = 5 ;      // nombre de ligne à ajouter à la grille
      MaxNbLigGenere = 100 ;    // nombre maximum de lignes à créer lors de la génération automatique
      IdValeurDefaut = '...' ;  // marque de l'enregistrement qui contient les valeurs par défaut


///////////////////////////////////////////////////////////////////////////////
// ExtractSuffixeSynonyme : Extrait le suffixe d'un nom de champ en tenant compte des synonymes
///////////////////////////////////////////////////////////////////////////////
Function ExtractSuffixeSynonyme ( NomChamp : String ) : String ;
Var Ind : Integer ;
BEGIN
Result := ExtractSuffixe(NomChamp) ;
for Ind := Low(ChampsSynonymes) to High(ChampsSynonymes) do
   BEGIN
   if ChampsSynonymes[Ind, 1] = Result then
      BEGIN
      Result := ChampsSynonymes[Ind, 2] ;
      Break;
      END ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////
// MiseEnFormeTradGPAOenPGI : Met en forme des valeurs PGI qui correspondent à une valeur GPAO d'un champ.
///////////////////////////////////////////////////////////////////////////////
Function MiseEnFormeTradGPAOenPGI ( TOB_Corr : TOB ; Valeur : String ) : String ;
Var Stg : String ;
    Ind : Integer ;
BEGIN
Result := Valeur ;
// Mise en forme des valeurs correspondantes trouvées.
Stg := TOB_Corr.GetValue('GRO_VALEURPGI') ;
if Stg <> '' then Result := Stg ;
for Ind := 1 to 7 do
   BEGIN
   Stg := TOB_Corr.GetValue('GRO_VALEURPGI' + IntToStr(Ind)) ;
   if Trim (Stg) <> '' then Result := Result + ';' + Stg ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////
// TradValTOBGPAOenPGI : traduit pour PGI une valeur GPAO d'un champ en utilisant une TOB pour minimiser les requêtes
///////////////////////////////////////////////////////////////////////////////
Function TradValTOBGPAOenPGI ( TypeReprise, NomChamp, Valeur : String ; TOBPrinc : TOB ; ToutPreCharge : Boolean ) : String ;
         ///////////////////////////////////////////////////////////////////////////////
         Function TrouveTOB ( ValeurDefaut : Boolean ) : TOB ;
         Var ValeurCle : String ;
         BEGIN
         if ValeurDefaut then ValeurCle := IdValeurDefaut else ValeurCle := Valeur ;
         Result := TOBPrinc.FindFirst(['GRO_TYPEREPRISE','GRO_NOMCHAMP','GRO_VALEURGPAO'],
                                      [TypeReprise, ExtractSuffixeSynonyme(NomChamp), ValeurCle], False) ;
         if (Result <> Nil) and (Result.GetValue('GRO_VALEURPGI') = IdValeurDefaut) then Result := Nil ;
         END ;
         ///////////////////////////////////////////////////////////////////////////////
Var TOB_Corr : TOB ;
BEGIN
Result := Valeur ;
// Recherche de la valeur dans la TOB principale
TOB_Corr := TrouveTOB(False) ;
// Recherche de la valeur par défaut
if TOB_Corr = Nil then TOB_Corr := TrouveTOB(True) ;
// Recherche des valeurs du champ dans la base
if (TOB_Corr = Nil) and (not ToutPreCharge) then
BEGIN
   if TOBPrinc.LoadDetailDB('REPRISECOGPAO', '"'+TypeReprise+'";"'+ExtractSuffixeSynonyme(NomChamp)+'"', '', Nil, True) then
   BEGIN
      //GcVoirTob(TOBPRINC);
      // Recherche de la valeur dans la TOB principale
      TOB_Corr := TrouveTOB(False) ;
      // Recherche de la valeur par défaut
      if TOB_Corr = Nil then TOB_Corr := TrouveTOB(True) ;
   END ;
END ;
if TOB_Corr = Nil then
   BEGIN
   if not ToutPreCharge then
      BEGIN
      // Insertion de la valeur par défaut à vide pour éviter de chercher à nouveau dans la base
      TOB_Corr := TOB.Create ('REPRISECOGPAO', TOBPrinc, -1) ;
      TOB_Corr.InitValeurs ;
      TOB_Corr.PutValue('GRO_TYPEREPRISE', TypeReprise) ;
      TOB_Corr.PutValue('GRO_NOMCHAMP', ExtractSuffixeSynonyme(NomChamp)) ;
      TOB_Corr.PutValue('GRO_VALEURGPAO', IdValeurDefaut) ;
      END ;
   END else
   BEGIN
   // Mise en forme des valeurs correspondantes trouvées.
   Result := MiseEnFormeTradGPAOenPGI(TOB_Corr, Valeur) ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////
// TradValGPAOenPGI : traduit pour PGI une valeur GPAO d'un champ en accédant à la table
///////////////////////////////////////////////////////////////////////////////
Function TradValGPAOenPGI ( TypeReprise, NomChamp, Valeur : String ) : String ;
Var TOB_Corr : TOB ;
    Ok       : Boolean ;
BEGIN
Result := Valeur ;
// Recherche de la valeur dans la base
TOB_Corr := TOB.Create ('REPRISECOGPAO', Nil, -1) ;
Ok := TOB_Corr.SelectDB('"'+TypeReprise+'";"'+ExtractSuffixeSynonyme(NomChamp)+'";"'+Valeur+'"', Nil) ;
if (TOB_Corr <> Nil) and (TOB_Corr.GetValue('GRO_VALEURPGI') = IdValeurDefaut) then Ok := False ;
// Recherche de la valeur par défaut
if not Ok then Ok := TOB_Corr.SelectDB('"'+TypeReprise+'";"'+ExtractSuffixeSynonyme(NomChamp)+'";"'+IdValeurDefaut+'"', Nil) ;
if (TOB_Corr <> Nil) and (TOB_Corr.GetValue('GRO_VALEURPGI') = IdValeurDefaut) then Ok := False ;
// Mise en forme des valeurs correspondantes trouvées.
if Ok then  Result := MiseEnFormeTradGPAOenPGI(TOB_Corr, Valeur) ;
TOB_Corr.Free ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 29/03/2002
Modifié le ... : 02/04/2002
Description .. : Retourne les valeurs correspondantes dans PGI à la valeur
Suite ........ : GPAO d'un champ sous la forme :
Suite ........ :   xxx;yyy;zzz
Mots clefs ... :
*****************************************************************}
Function TraduitGPAOenPGI ( TypeReprise, NomChamp, Valeur : String ; TOBPrinc : TOB = Nil ; ToutPreCharge : Boolean = False ) : String ;
BEGIN
Valeur := Trim(Valeur) ;
if TOBPrinc = Nil then Result := TradValGPAOenPGI(TypeReprise, NomChamp, Valeur)
                  else Result := TradValTOBGPAOenPGI(TypeReprise, NomChamp, Valeur, TOBPrinc, ToutPreCharge) ;
END ;

///////////////////////////////////////////////////////////////////////////////
// TradValTOBGPAOenPGI : traduit pour PGI une valeur GPAO d'un champ en utilisant une TOB pour minimiser les requêtes
///////////////////////////////////////////////////////////////////////////////
Function TradValTOBPGIenGPAO ( TypeReprise, NomChamp : String ; Valeurs : Array of String ; AvecLibelle : Boolean ; TOBPrinc : TOB ; ToutPreCharge : Boolean ) : String ;
         ///////////////////////////////////////////////////////////////////////////////
         Function TrouveTOB ( ValeurDefaut : Boolean ) : TOB ;
         Var NomCle    : Array of String ;
             ValeurCle : Array of Variant ;
             Ind       : Integer ;
         BEGIN
         if ValeurDefaut then
            BEGIN
            // valeur par defaut
            SetLength(NomCle, 3) ;
            NomCle[0] := 'GRO_TYPEREPRISE' ;
            NomCle[1] := 'GRO_NOMCHAMP' ;
            NomCle[2] := 'GRO_VALEURGPAO' ;
            SetLength(ValeurCle, 3) ;
            ValeurCle[0] := TypeReprise ;
            ValeurCle[1] := ExtractSuffixeSynonyme(NomChamp) ;
            ValeurCle[2] := IdValeurDefaut ;
            Result := TOBPrinc.FindFirst(NomCle, ValeurCle, False) ;
            END else
            BEGIN
            // valeur exacte
            SetLength(NomCle, High(Valeurs)+3) ;
            NomCle[0] := 'GRO_TYPEREPRISE' ;
            NomCle[1] := 'GRO_NOMCHAMP' ;
            SetLength(ValeurCle, High(Valeurs)+3) ;
            ValeurCle[0] := TypeReprise ;
            ValeurCle[1] := ExtractSuffixeSynonyme(NomChamp) ;
            for Ind := Low(Valeurs) to High(Valeurs) do
               BEGIN
               NomCle[Ind+2] := 'GRO_VALEURPGI' ;
               if Ind > 0 then NomCle[Ind+2] := NomCle[Ind+2] + IntToStr(Ind) ;
               ValeurCle[Ind+2] := Valeurs[Ind] ;
               END ;
            Result := TOBPrinc.FindFirst(NomCle, ValeurCle, False) ;
            if (Result <> Nil) and (Result.GetValue('GRO_VALEURGPAO') = IdValeurDefaut) then
               BEGIN
               repeat
                  Result := TOBPrinc.FindNext(NomCle, ValeurCle, False) ;
               until (Result = Nil) or (Result.GetValue('GRO_VALEURGPAO') <> IdValeurDefaut) ;
               END ;
            END ;
         END ;
         ///////////////////////////////////////////////////////////////////////////////
Var TOB_Corr       : TOB ;
    sNomChamp, Stg : String ;
BEGIN
Result := Valeurs[Low(Valeurs)] ;
// Recherche de la valeur dans la TOB principale
sNomChamp := 'GRO_VALEURGPAO' ;
TOB_Corr := TrouveTOB(False) ;
// Recherche de la valeur par défaut
if TOB_Corr = Nil then
   BEGIN
   sNomChamp := 'GRO_OPTIONS' ;
   TOB_Corr := TrouveTOB(True) ;
   END ;
// Recherche des valeurs du champ dans la base
if (TOB_Corr = Nil) and (not ToutPreCharge) then
   BEGIN
   if TOBPrinc.LoadDetailDB('REPRISECOGPAO', '"'+TypeReprise+'";"'+ExtractSuffixeSynonyme(NomChamp)+'"', '', Nil, True) then
      BEGIN
      // Recherche de la valeur dans la TOB principale
      sNomChamp := 'GRO_VALEURGPAO' ;
      TOB_Corr := TrouveTOB(False) ;
      // Recherche de la valeur par défaut
      if TOB_Corr = Nil then
         BEGIN
         sNomChamp := 'GRO_OPTIONS' ;
         TOB_Corr := TrouveTOB(True) ;
         END ;
      END ;
   END ;
if TOB_Corr = Nil then
   BEGIN
   if not ToutPreCharge then
      BEGIN
      // Insertion de la valeur par défaut à vide pour éviter de chercher à nouveau dans la base
      TOB_Corr := TOB.Create ('REPRISECOGPAO', TOBPrinc, -1) ;
      TOB_Corr.InitValeurs ;
      TOB_Corr.PutValue('GRO_TYPEREPRISE', TypeReprise) ;
      TOB_Corr.PutValue('GRO_NOMCHAMP', ExtractSuffixeSynonyme(NomChamp)) ;
      TOB_Corr.PutValue('GRO_VALEURGPAO', IdValeurDefaut) ;
      END ;
   END else
   BEGIN
   Stg := TOB_Corr.GetValue(sNomChamp) ;
   if Stg <> '' then Result := Stg ;
   if AvecLibelle then
      BEGIN
      Stg := TOB_Corr.GetValue('GRO_LIBELLEGPAO') ;
      if Stg <> '' then Result := Result + ';' + Stg ;
      END ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////
// TradValPGIenGPAO : traduit en une valeur GPAO un ensemble de valeurs d'un champ en accédant à la table
///////////////////////////////////////////////////////////////////////////////
Function TradValPGIenGPAO ( TypeReprise, NomChamp : String ; Valeurs : Array of String ; AvecLibelle : Boolean ) : String ;
Var TOB_Corr              : TOB ;
    sSql, sNom, sNomChamp : String ;
    Ind                   : Integer ;
    Ok                    : Boolean ;
    QQ                    : TQuery ;
BEGIN
Result := Valeurs[Low(Valeurs)] ;
TOB_Corr := TOB.Create('REPRISECOGPAO', Nil, -1) ;
Ok := False ;
// Recherche de la valeur dans la base
sNomChamp := 'GRO_VALEURGPAO' ;
sSql := 'Select * from REPRISECOGPAO where GRO_TYPEREPRISE="'+ TypeReprise +'"'
      + ' and GRO_NOMCHAMP="'+ ExtractSuffixeSynonyme(NomChamp) +'"'
      + ' and GRO_VALEURGPAO<>"'+IdValeurDefaut+'"' ;
for Ind := Low(Valeurs) to High(Valeurs) do
   BEGIN
   if Ind in [0..7] then
      BEGIN
      sNom := 'GRO_VALEURPGI' ;
      if Ind > 0 then sNom := sNom + IntToStr(Ind) ;
      sSql := sSql + ' and '+ sNom +'="'+ Valeurs[Ind] +'"'
      END ;
   END ;
QQ := OpenSQL(sSql, True,-1,'',true) ;
if not QQ.Eof then
   BEGIN
   TOB_Corr.SelectDB('', QQ) ;
   Ok := True ;
   END ;
Ferme(QQ) ;
// Recherche de la valeur par défaut
if not Ok then
   BEGIN
   sNomChamp := 'GRO_OPTIONS' ;
   sSql := 'Select * from REPRISECOGPAO where GRO_TYPEREPRISE="'+ TypeReprise +'"'
         + ' and GRO_NOMCHAMP="'+ ExtractSuffixeSynonyme(NomChamp) +'"'
         + ' and GRO_VALEURGPAO="'+IdValeurDefaut+'"' ;
   QQ := OpenSQL(sSql, True,-1,'',true) ;
   if not QQ.Eof then
      BEGIN
      TOB_Corr.SelectDB('', QQ) ;
      Ok := True ;
      END ;
   Ferme(QQ) ;
   END ;
if Ok then
   BEGIN
   Result := TOB_Corr.GetValue(sNomChamp) ;
   if AvecLibelle then Result := Result + ';' + TOB_Corr.GetValue('GRO_LIBELLEGPAO') ;
   END ;
TOB_Corr.Free ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 29/03/2002
Modifié le ... : 02/04/2002
Description .. : Retourne la valeur GPAO correspondante à un ensemble
Suite ........ : de valeurs PGI d'un champ sous la forme :
Suite ........ :     code;libellé
Mots clefs ... :
*****************************************************************}
Function TraduitPGIenGPAO ( TypeReprise, NomChamp : String ; Valeurs : Array of String ; AvecLibelle : Boolean = False ; TOBPrinc : TOB = Nil ; ToutPreCharge : Boolean = False ) : String ;
BEGIN
if TOBPrinc = Nil then Result := TradValPGIenGPAO(TypeReprise, NomChamp, Valeurs, AvecLibelle)
                  else Result := TradValTOBPGIenGPAO(TypeReprise, NomChamp, Valeurs, AvecLibelle, TOBPrinc, ToutPreCharge) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 29/03/2002
Modifié le ... : 29/03/2002
Description .. : Saisie des tables de correspondances de la reprise GPAO
Mots clefs ... :
*****************************************************************}
Procedure AppelSaisieCorrespGPAO ( Action : TActionFiche ; TypeReprise, NomChamp : string ) ;
var XX  : TFCorrespGPAO ;
    PP  : THPanel ;
    Ind : Integer ;
    Ok  : Boolean ;
BEGIN
// contrôle d'existence d'une tablette associée au champ
Ok := False ;
// Recherche s'il s'agit d'un champ spécifique
for Ind := Low(ChampsSpecifiques) to High(ChampsSpecifiques) do
   BEGIN
   if NomChamp = ChampsSpecifiques[Ind][1] then
      BEGIN
      Ok := True ;
      Break ;
      END ;
   END ;
// Recherche de la tablette associée
if not Ok then Ok := (Get_Join(NomChamp) <> '') ;
if not Ok then
   BEGIN
   PGIBox('Il n''existe aucune tablette associée au champ.', 'Saisie des tables de correspondances') ;
   Exit ;
   END ;
// Affichage de l'écran de saisie
PP := FindInsidePanel ;
XX := TFCorrespGPAO.Create(Application) ;
XX.Action := Action ;
XX.GRO_TYPEREPRISE.Value := TypeReprise ;
XX.GRO_NOMCHAMP.Text := NomChamp ;
if PP = Nil then
   BEGIN
   try
      XX.ShowModal ;
    finally
      XX.Free ;
    end ;
   END else
   BEGIN
   InitInside(XX, PP) ;
   XX.Show ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  CompleteTOBCorr : compléte la TOB pour atteindre un nombre minimum de lignes
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.CompleteTOBCorr ;
Var Ind : Integer ;
BEGIN
if GS.RowCount < NbRowsInit then
  BEGIN
  for Ind := GS.RowCount to NbRowsInit do InsereTOBCorr(Ind) ;
  GS.RowCount := NbRowsInit ;
  END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  InsereTOBCorr : insertion d'une TOB
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.InsereTOBCorr ( ARow : Integer ; SiBesoin : Boolean = False ) ;
Var TOBL : TOB ;
    Ind  : Integer ;
BEGIN
if (SiBesoin) and (GetTobCorr(ARow) <> Nil) then Exit ;
if TOB_Corr.Detail.Count < ARow then Ind := -1 else Ind := ARow -1 ;
TOBL := TOB.Create('REPRISECOGPAO', TOB_Corr, Ind) ;
InitTOBCorr(TOBL) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  InitTOBCorr : initialisation d'une TOB
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.InitTOBCorr ( TOBLigne : TOB ) ;
BEGIN
if TOBLigne <> Nil then
   BEGIN
   TOBLigne.InitValeurs ;
   TOBLigne.PutValue('GRO_TYPEREPRISE', GRO_TYPEREPRISE.Value) ;
   TOBLigne.PutValue('GRO_NOMCHAMP', ExtractSuffixeSynonyme(GRO_NOMCHAMP.Text)) ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  GetTobCorr : retourne la TOB d'une ligne
///////////////////////////////////////////////////////////////////////////////////////
function TFCorrespGPAO.GetTobCorr ( ARow : Integer ) : TOB ;
BEGIN
if ((ARow > 0) and (ARow <= TOB_Corr.Detail.Count)) then Result := TOB_Corr.Detail[ARow -1] else Result := Nil ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  DepileTOBCorr : supprime les TOB vides
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.DepileTOBCorr ( TOBPrinc : TOB ) ;
Var TOBL     : TOB ;
    Ind, jj  : Integer ;
    sColName : String ;
    Libere, ConserveValeurPgiVide : Boolean ;
BEGIN
ConserveValeurPgiVide := True ;
for Ind := TOBPrinc.Detail.Count -1 downto 0 do
   BEGIN
   TOBL := TOBPrinc.Detail[Ind] ;
   Libere := True ;
   if TOBL.GetValue('GRO_VALEURGPAO') <> '' then
      BEGIN
      if ConserveValeurPgiVide then
         BEGIN
         Libere := False ;
         END else
         BEGIN
           for jj := Low(NomTable) to High(NomTable) do
              BEGIN
              sColName := 'GRO_VALEURPGI' ;
              if jj > 1 then sColName := sColName + IntToStr(jj -1) ;
              if TOBL.GetValue(sColName) <> '' then
                 BEGIN
                 Libere := False ;
                 Break ;
                 END ;
             END;
         END;
       END;
   if Libere then TOBL.Free ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  AjoutValeurDefaut : ajout dans la TOB de la table de correspondance les valeurs par defaut
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.AjoutValeurDefaut ( TOBPrinc : TOB ) ;
Var TOBL          : TOB ;
    Stg, sColName : String ;
    Ind           : Integer ;
    Ctrl          : TComponent ;
BEGIN
TOBL := Nil ;
// Valeur par défaut GPAO
Stg := Trim(VALDEFAUTGPAO.Text) ;
if Stg <> '' then
   BEGIN
   TOBL := TOB.Create('REPRISECOGPAO', TOBPrinc, -1) ;
   InitTOBCorr(TOBL) ;
   TOBL.PutValue('GRO_VALEURGPAO', IdValeurDefaut) ;
   TOBL.PutValue('GRO_OPTIONS', Stg) ;
   TOBL.PutValue('GRO_LIBELLEGPAO', LIBDEFAUTGPAO.Text) ;
   END ;
// Valeurs par défaut PGI
Stg := Trim(VALDEFAUTPGI1.Text) ;
if Stg <> '' then
   BEGIN
   if TOBL = Nil then
      BEGIN
      TOBL := TOB.Create('REPRISECOGPAO', TOBPrinc, -1) ;
      InitTOBCorr(TOBL) ;
      TOBL.PutValue('GRO_VALEURGPAO', IdValeurDefaut) ;
      END ;
   for Ind := Low(NomTable) to High(NomTable) do
       BEGIN
       Ctrl := FindComponent('VALDEFAUTPGI' + IntToStr(Ind)) ;
       sColName := 'GRO_VALEURPGI' ;
       if Ind > 1 then sColName := sColName + IntToStr(Ind -1) ;
       if (Ctrl <> Nil) and (Ctrl is THCritMaskEdit) then Stg := Trim(THCritMaskEdit(Ctrl).Text) ;
       if Stg <> '' then TOBL.PutValue(sColName, Stg) ;
       END ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  EnregistreSaisie : enregistre les données saisies
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.EnregistreSaisie ;
Var Nb   : Integer ;
    sSql : String ;
BEGIN
// Suppression des valeurs saisies précédemment
sSql := 'DELETE FROM REPRISECOGPAO WHERE GRO_TYPEREPRISE="'+ GRO_TYPEREPRISE.Value+'"'
      + ' AND GRO_NOMCHAMP="'+ExtractSuffixeSynonyme(GRO_NOMCHAMP.Text)+'"' ;
Nb := ExecuteSQL(sSql) ;
if Nb < 0 then
   BEGIN
   V_PGI.IoError := oeUnknown ;
   Exit ;
   END ;
// Enregistre la TOB des valeurs saisies
TOB_Corr.SetAllModifie(True) ;
if not TOB_Corr.InsertDB(Nil, True) then
   BEGIN
   V_PGI.IoError := oeUnknown ;
   Exit ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////
// ChargeTableCorrespondance : charge en TOB la table de correspondance associé au champ
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.ChargeTableCorrespondance ;
Var sColName : String ;
    Ind      : Integer ;
    TOBL     : TOB ;
    Ctrl     : TComponent ;
BEGIN
if Action = taConsult then Exit ;
if TOB_Corr = Nil then Exit ;
// chargement de la table de correspondance
TOB_Corr.LoadDetailDB('REPRISECOGPAO', '"'+GRO_TYPEREPRISE.Value+'";"'+ExtractSuffixeSynonyme(GRO_NOMCHAMP.Text)+'"', '', Nil, False) ;
// traitement des valeurs par défaut
TOBL := TOB_Corr.FindFirst(['GRO_VALEURGPAO'], [IdValeurDefaut], False) ;
if TOBL <> Nil then
   BEGIN
   VALDEFAUTGPAO.Text := TOBL.GetValue('GRO_OPTIONS') ;
   LIBDEFAUTGPAO.Text := TOBL.GetValue('GRO_LIBELLEGPAO') ;
   for Ind := Low(NomTable) to High(NomTable) do
       BEGIN
       Ctrl := FindComponent('VALDEFAUTPGI' + IntToStr(Ind)) ;
       sColName := 'GRO_VALEURPGI' ;
       if Ind > 1 then sColName := sColName + IntToStr(Ind -1) ;
       if (Ctrl <> Nil) and (Ctrl is THCritMaskEdit) then THCritMaskEdit(Ctrl).Text := TOBL.GetValue(sColName) ;
       END ;
   TOBL.Free ;
   END ;
AffichageTableCorrespondance ;
Modifie := False ;
END ;

///////////////////////////////////////////////////////////////////////////////
// AffichageTableCorrespondance : affichage dans la grille de la TOB de la table de correspondance
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.AffichageTableCorrespondance ;
Var sColName : String ;
    Ind      : Integer ;
BEGIN
// affichage dans la grille
sColName := 'GRO_VALEURGPAO;GRO_LIBELLEGPAO;GRO_VALEURPGI' ;
for Ind := 1 to NbCol do sColName := sColName + ';GRO_VALEURPGI' + IntToStr(Ind) ;
TOB_Corr.PutGridDetail(GS , False, False, sColName, True) ;
END ;

///////////////////////////////////////////////////////////////////////////////
// ChargeGrille : met à jour la TOB lavec les valeur d'une ligne de la grille de saisie
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.ChargeGrille ( ARow : Integer ) ;
Var sColName : String ;
    Ind      : Integer ;
    TOBL     : TOB ;
BEGIN
if Action = taConsult then Exit ;
if TOB_Corr = Nil then Exit ;
TOBL := GetTobCorr(ARow) ;
if TOBL = Nil then Exit ;
sColName := ';GRO_VALEURGPAO;GRO_LIBELLEGPAO;GRO_VALEURPGI' ;
for Ind := 1 to NbCol do sColName := sColName + ';GRO_VALEURPGI' + IntToStr(Ind) ;
TOBL.GetLigneGrid(GS, ARow, sColName) ;
END ;

///////////////////////////////////////////////////////////////////////////////
// BrancheTablette : associe une tablette à une colonne de la grille
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BrancheTablette ( Var Rang : Integer ; Tablette : String ) ;
Var Ind, ARow : Integer ;
    Ctrl      : TComponent ;
    sType     : String ;
    Combo     : THValComboBox ;
BEGIN
if Tablette = '' then Exit ;
// contrôle si la tablette est renseignée
sType := TTToTipe(Tablette) ;
if sType <> '' then
   BEGIN
   Combo := THValComboBox.Create(Self) ;
   Combo.Parent := Self ;
   Combo.Visible := False ;
   Combo.DataType := Tablette ;
   Ind := Combo.Values.Count ;
   Combo.Free ;
   if Ind <= 0 then Exit ;
   END ;
Inc(Rang) ;
ARow := ColDebut + Rang -1 ;
LookUpMode[Rang] := (sType = '') ;
NomTable[Rang] := Tablette ;
// associe une tablette à une colonne de la grille
if LookUpMode[Rang] then GS.ColFormats[ARow] := ''
                    else GS.ColFormats[ARow] := 'CB='+ Tablette ;
for Ind := (GS.Titres.Count -1) to (ARow) do GS.Titres.Add('') ;
GS.Titres.Strings[ARow] := TTToLibelle(Tablette) ;
// associe une tablette à la zone de saisie de la valeur par defaut
Ctrl := FindComponent('VALDEFAUTPGI' + IntToStr(Rang)) ;
if (Ctrl <> Nil) and (Ctrl is THCritMaskEdit) then
   BEGIN
   THCritMaskEdit(Ctrl).DataType := Tablette ;
   THCritMaskEdit(Ctrl).Visible := True ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////
// AdapteGrille : adapte les prpriétés de la grille en fonction du champ choisi
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.AdapteGrille ;
Var iRang, Ind : Integer ;
BEGIN
// Recherche s'il s'agit d'un champ spécifique
iRang := 0 ;
for Ind := Low(ChampsSpecifiques) to High(ChampsSpecifiques) do
   BEGIN
   if GRO_NOMCHAMP.Text = ChampsSpecifiques[Ind][1] then
      BEGIN
      iRang := Ind ;
      Break ;
      END ;
   END ;
NbCol := 0;
if iRang > 0 then
   BEGIN
   // Cas d'un champ spécifique
   for Ind := 2 to High(ChampsSpecifiques[iRang]) do
      BrancheTablette(NbCol, ChampsSpecifiques[iRang][Ind]) ;
   END else
   BEGIN
   // Recherche de la tablette associée
   BrancheTablette(NbCol, Get_Join(GRO_NOMCHAMP.Text)) ;
   END ;
// colonnes inaccessibles
for Ind := GS.FixedCols to GS.ColCount -1 do
   BEGIN
   if Ind < (ColDebut + NbCol) then
      BEGIN
      GS.ColLengths[Ind] := 35 ;
      GS.ColWidths[Ind] := 35 ;
      END else
      BEGIN
      GS.ColLengths[Ind] := -1 ;
      GS.ColWidths[Ind] := -1 ;
      END ;
   END ;
GS.UpdateTitres ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  ActiveBoutons : active ou non les boutons
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.ActiveBoutons ( Actif : Boolean ) ;
BEGIN
if Action <> taConsult then
   BEGIN
   BValider.Enabled := Actif ;
   bNewligne.Enabled := Actif ;
   bDelLigne.Enabled := Actif ;
   BGenere.Enabled := Actif ;
   END ;
BFerme.Enabled := Actif ;
BChercher.Enabled := Actif ;
BtestGPAOenPGI.Enabled := Actif ;
BtestPGIenGPAO.Enabled := Actif ;
END ;

///////////////////////////////////////////////////////////////////////////////
// GSEnter
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.GSEnter(Sender: TObject);
Var bc, Cancel : Boolean ;
    ACol, ARow : Integer ;
BEGIN
if Action = taConsult then Exit ;
bc := False ;
Cancel := False ;
ACol := GS.Col ;
ARow := GS.Row ;
GSRowEnter(GS, GS.Row, bc, False) ;
GSCellEnter(GS, ACol, ARow, Cancel) ;
END ;

///////////////////////////////////////////////////////////////////////////////
// GSRowEnter
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN
if GS.Row >= GS.RowCount -1 then GS.RowCount := GS.RowCount + NbRowsPlus ;
if Action <> taConsult then InsereTOBCorr(GS.Row, True) ;
END ;

///////////////////////////////////////////////////////////////////////////////
// GSRowExit
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN
if Action = taConsult then Exit ;
// récupération des éventuelles valeurs saisies ou modifié
ChargeGrille(Ou) ;
END ;

///////////////////////////////////////////////////////////////////////////////
// GSCellEnter
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Var Ind : Integer ;
BEGIN
if Action = taConsult then Exit ;
if (GS.Col > COL_CodeGpao) and (Trim(GS.Cells[COL_CodeGpao, GS.Row]) = '') then Cancel := True ;
if Not Cancel then
   BEGIN
   GS.ElipsisButton := False ;
   if (GS.Col >= ColDebut) and (GS.Col < (ColDebut + NbCol)) and (GS.ColLengths[GS.Col] > 0) and (GS.ColWidths[GS.Col] > 0) then
      BEGIN
      Ind := (GS.Col - ColDebut + 1) ;
      GS.ElipsisButton := ((LookUpMode[Ind]) and (NomTable[Ind] <> '')) ;
      END ;
   StCellCur := GS.CellValues[GS.Col, GS.Row] ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////
// GSCellExit
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Var Ind  : Integer ;
    TOBL : TOB ;
    Stg  : String ;
BEGIN
if csDestroying in ComponentState then Exit ;
if Action = taConsult then Exit ;
if GS.CellValues[ACol, ARow] = StCellCur then Exit ;
Modifie := True ;
if ACol = COL_CodeGpao then
   BEGIN
   // utilisation de la marque des valeurs par défaut interdite
   if GS.CellValues[ACol, ARow] = IdValeurDefaut then
      BEGIN
      Cancel := True ;
      PGIBox('L''utilisation de cette valeur est interdite.', GS.Titres.Strings[ACol]) ;
      END else
      BEGIN
      // contrôle d'unicité de la valeurs GPAO
      TOBL := TOB_Corr.FindFirst(['GRO_VALEURGPAO'], [GS.CellValues[ACol, ARow]], False) ;
      if TOBL <> Nil then
         BEGIN
         if TOBL.GetIndex = ARow -1 then
            TOBL := TOB_Corr.FindNext(['GRO_VALEURGPAO'], [GS.CellValues[ACol, ARow]], False) ;
         if TOBL <> Nil then
            BEGIN
            Cancel := True ;
            PGIBox('Valeur déjà saisie.', GS.Titres.Strings[ACol]) ;
            END ;
         END ;
      END ;
   END else
if (ACol >= ColDebut) and (ACol < (ColDebut + NbCol)) and (GS.ColLengths[ACol] > 0) and (GS.ColWidths[ACol] > 0) then
   BEGIN
   // contrôle d'existence des valeurs PGI
   Ind := (ACol - ColDebut + 1) ;
   Stg := RechDom(NomTable[Ind], GS.CellValues[ACol, ARow], False) ;
   if UpperCase(Stg) = 'ERROR' then
      BEGIN
      Cancel := True ;
      PGIBox('Valeur inconnue.', GS.Titres.Strings[ACol]) ;
      END ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////
// GSElipsisClick
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.GSElipsisClick(Sender: TObject);
Var sPrefixe, sTitre, sTable, sColonne, sSelect, sWhere, sOrdre : String ;
    Ind, ACol : Integer ;
BEGIN
if Action = taConsult then Exit ;
if (GS.Col <= 0) or (GS.Col >= GS.ColCount) then Exit ;
if (GS.ColLengths[GS.Col] < 0) or (GS.ColWidths[GS.Col] < 0) then Exit ;
ACol := (GS.Col - ColDebut + 1) ;
if (not LookUpMode[ACol]) or (NomTable[ACol] = '') then Exit ;
GetCorrespType(NomTable[ACol], sTable, sSelect, sWhere, sPrefixe, sColonne) ;
//ChangeWhereCombo(sWhere) ;
ChangeWhereTT(sWhere, '', FALSE) ;
Ind := Pos('DISTINCT ', sSelect) ;
if Ind >0 then
   BEGIN
   sColonne := sSelect ;
   sSelect := Copy(sSelect, Ind +9, Length(sSelect)) ;
   sOrdre := sSelect ;
   END else
   BEGIN
   sOrdre := sColonne ;
   Ind := Pos(' ORDER BY ', sWhere) ;
   if Ind > 0 then
      BEGIN
      sOrdre := Copy(sWhere, Ind +9, Length(sWhere)) ;
      Delete(sWhere, Ind, Length(sWhere)) ;
      END ;
   END ;
sTitre := GS.Titres.Strings[GS.Col] ;
LookupList(GS, sTitre, sTable, sSelect, sColonne, sWhere, sOrdre, True, 0) ;
END ;

///////////////////////////////////////////////////////////////////////////////
// ValideSaisie : enregistre les éléments saisis
///////////////////////////////////////////////////////////////////////////////
function TFCorrespGPAO.ValideSaisie : Boolean ;
Var bc, cancel : Boolean ;
    ACol, ARow : Integer ;
    io         : TIOErr ;
BEGIN
Result := False ;
if Action = taConsult then Exit ;
if not BValider.Enabled then Exit ;
BValider.Enabled := False ;
if not BGenere.Enabled then Exit ;
BGenere.Enabled := False ;
bc := False ;
Cancel := False ;
ACol := GS.Col ;
ARow := GS.Row ;
GSCellExit(GS, ACol, ARow, Cancel) ;
if Cancel then
   BEGIN
   BValider.Enabled := True ;
   BGenere.Enabled := True ;
   Exit ;
   END ;
GSRowExit(GS, ARow, bc, False) ;
DepileTOBCorr(TOB_Corr) ;
AjoutValeurDefaut(TOB_Corr) ;
if (Modifie) or (TOB_Corr.IsOneModifie) then
   BEGIN
   SourisSablier ;
   io := Transactions(EnregistreSaisie, 2) ;
   Case io of
      oeOk      : BEGIN Modifie := False ; Result := True ; END ;
      oeUnknown : PGIBox('ATTENTION : Données non enregistrées !', Caption) ;
      oeSaisie  : PGIBox('ATTENTION : Ces données en cours de traitement par un autre utilisateur n''ont pas été enregistrées !', Caption) ;
      END ;
   SourisNormale ;
   END ;
BValider.Enabled := True ;
BGenere.Enabled := True ;
END ;

///////////////////////////////////////////////////////////////////////////////
// VALDEFAUTChange
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.VALDEFAUTChange(Sender: TObject);
BEGIN
Modifie := True ;
END ;

///////////////////////////////////////////////////////////////////////////////
// BValiderClick
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BValiderClick(Sender: TObject);
BEGIN
if ValideSaisie then Close ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  bNewligneClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.bNewligneClick(Sender: TObject);
Var ARow : Integer ;
BEGIN
if Action = taConsult then Exit ;
if ((GS.Row < 1) or (GS.Row > TOB_Corr.Detail.Count)) then Exit ;
ARow := GS.Row ;
GS.CacheEdit ;
GS.SynEnabled := False ;
// Insertion d'une ligne
GS.InsertRow(ARow) ;
InsereTOBCorr(ARow) ;
GS.Col := GS.FixedCols ;
GS.Row := ARow ;
GS.MontreEdit ;
GS.SynEnabled := True ;
END;

///////////////////////////////////////////////////////////////////////////////////////
//  bDelLigneClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.bDelLigneClick(Sender: TObject);
Var bc, Cancel        : Boolean ;
    ARow, ACol, NoLig : Integer ;
BEGIN
if Action = taConsult then Exit ;
if ((GS.Row < 1) or (GS.Row > TOB_Corr.Detail.Count)) then Exit ;
ARow := GS.Row ;
GS.CacheEdit ;
GS.SynEnabled := False ;
// Suppression de la ligne
NoLig:=GS.TopRow ;
GS.DeleteRow(ARow) ;
if NoLig > 1 then GS.TopRow := (NoLig -1) else GS.TopRow := 1 ;
if ((TOB_Corr.Detail.Count > 1) and (ARow <> TOB_Corr.Detail.Count)) then
     TOB_Corr.Detail[ARow-1].Free
else InitTOBCorr(TOB_Corr.Detail[ARow-1]) ;
if GS.RowCount < NbRowsInit then GS.RowCount := GS.RowCount + NbRowsPlus ;
GS.MontreEdit ;
GS.SynEnabled := True ;
bc := False ;
Cancel := False ;
GS.Col := GS.FixedCols ;
if GS.Row > GS.FixedRows then
   BEGIN
   ARow := GS.Row ;
   Dec(ARow) ;
   GS.Row := ARow ;
   END ;
ACol := GS.Col ;
GSRowEnter(GS, ARow, bc, False) ;
GSCellEnter(GS, ACol, ARow, Cancel) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  GenereValeurs : Génération des valeurs PGI sans correspondance
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.GenereValeurs ( Champs : Array of String ; Valeurs : Array of Variant ; Rang : Integer ; Var First : Boolean ; Var NbLig, ARow : Integer ) ;
Var Ctrl    : TComponent ;
    TOBL    : TOB ;
    Ind, jj : Integer ;
BEGIN
Ctrl := FindComponent('ComboGenere_' + IntToStr(Rang)) ;
if (Ctrl = Nil) or not (Ctrl is THValComboBox) then Exit ;
for Ind := 0 to THValComboBox(Ctrl).Values.Count -1 do
   BEGIN
   // on compléte les éléments de recherche
   Champs[Rang] := 'GRO_VALEURPGI' ;
   if Rang > 0 then Champs[Rang] := Champs[Rang] + IntToStr(Rang) ;
   Valeurs[Rang] := THValComboBox(Ctrl).Values[Ind] ;
   if Rang = (NbCol -1) then
      BEGIN
      // recherche si les valeurs existent dans la TOB
      TOBL := TOB_Corr.FindFirst(Champs, Valeurs, False) ;
      if TOBL = Nil then
         BEGIN
         if First then
            BEGIN
            DepileTOBCorr(TOB_Corr) ;
            First := False ;
            ARow := TOB_Corr.Detail.Count + GS.FixedRows ;
            END ;
         TOBL := TOB.Create('REPRISECOGPAO', TOB_Corr, -1) ;
         InitTOBCorr(TOBL) ;
         for jj := Low(Champs) to High(Champs) do TOBL.PutValue(Champs[jj], Valeurs[jj]) ;
         Inc(NbLig) ;
         END ;
      END else
      BEGIN
      // appel du niveau suivant
      GenereValeurs(Champs, Valeurs, (Rang +1), First, NbLig, ARow) ;
      if NbLig >= MaxNbLigGenere then Break ;
      END ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BGenereClick : Génération des valeurs PGI sans correspondance
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BGenereClick(Sender: TObject);
Var Combo                 : THValComboBox ;
    Ctrl                  : TComponent ;
    Ind, NbLig, Rep, ARow : Integer ;
    First                 : Boolean ;
    Champs                : Array of String ;
    Valeurs               : Array of Variant ;
BEGIN
if not BGenere.Enabled then Exit ;
BGenere.Enabled := False ;
if not BValider.Enabled then Exit ;
BValider.Enabled := False ;
SourisSablier ;
// création des combos pour récupérer les codes
for Ind := 0 to NbCol -1 do
   BEGIN
   Combo := THValComboBox.Create(Self) ;
   Combo.Parent := Self ;
   Combo.Name := 'ComboGenere_' + IntToStr(Ind) ;
   Combo.Visible := False ;
   Combo.DataType := NomTable[Ind +1] ;
   END ;
// calcul du nombre de croissements à étudier
NbLig := 1 ;
for Ind := 0 to NbCol -1 do
   BEGIN
   Ctrl := FindComponent('ComboGenere_' + IntToStr(Ind)) ;
   if (Ctrl <> Nil) and (Ctrl is THValComboBox) then NbLig := NbLig * THValComboBox(Ctrl).Values.Count ;
   END ;
if NbLig > MaxNbLigGenere then Rep := PGIAsk(IntToStr(NbLig)+' éléments vont être étudiés. Voulez-vous continuer ?', 'Génération des valeurs sans correspondance')
                          else Rep := mrYes ;
// Génération des valeurs PGI sans correspondance
if Rep = mrYes then
   BEGIN
   First := True ;
   NbLig := 0 ;
   SetLength(Champs, NbCol) ;
   SetLength(Valeurs, NbCol) ;
   GenereValeurs(Champs, Valeurs, 0, First, NbLig, ARow) ;
   if NbLig > 0 then PGIInfo(IntToStr(NbLig)+' lignes ont été ajoutées.', 'Génération des valeurs sans correspondance') ;
   if not First then
      BEGIN
      AffichageTableCorrespondance ;
      CompleteTOBCorr ;
      if GS.CanFocus then GS.SetFocus ;
      GS.GotoRow(ARow) ;
      END ;
   END ;
// suppression des combos pour récupérer les codes
for Ind := 0 to NbCol -1 do
   BEGIN
   Ctrl := FindComponent('ComboGenere_' + IntToStr(Ind)) ;
   if (Ctrl <> Nil) and (Ctrl is THValComboBox) then THValComboBox(Ctrl).Free ;
   END ;
BValider.Enabled := True ;
BGenere.Enabled := True ;
SourisNormale ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BVoirTobClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BVoirTobClick(Sender: TObject);
BEGIN
//GCVoirTob(TOB_Corr) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BImprimeClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BImprimeClick(Sender: TObject);
BEGIN
PrintDBGrid(GS, PENTETE, Caption, '');
END ;

///////////////////////////////////////////////////////////////////////////////
// BChercherClick
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BChercherClick(Sender: TObject);
begin
if GS.RowCount <= 2 then Exit ;
FindFirst := True ;
FindLigne.Execute ;
end;

///////////////////////////////////////////////////////////////////////////////
// FindLigneFind
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.FindLigneFind(Sender: TObject);
begin
Rechercher(GS, FindLigne, FindFirst) ;
end;

///////////////////////////////////////////////////////////////////////////////
// FormKeyDown
///////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG : Boolean ;
BEGIN
OkG := (Screen.ActiveControl = GS) ;
if Key = VK_RECHERCHE then
   BEGIN
   if (OkG) and (Shift = []) then
       BEGIN
       Key := 0 ;
       GSElipsisClick(GS) ;
       END ;
   END else
if Key = VK_VALIDE then
   BEGIN
   if Shift = [] then
      BEGIN
      Key := 0 ;
      BValiderClick(Nil) ;
      END ;
   END else
if Key = VK_ESCAPE then
   BEGIN
   if Shift = [] then
      BEGIN
      Key := 0 ;
      Close ;
      END ;
   END else
if Key = VK_INSERT then
   BEGIN
   if ((OkG) and (Shift = [])) then
      BEGIN
      Key := 0 ;
      bNewligneClick(Nil) ;
      END ;
   END else
if Key = VK_DELETE then
   BEGIN
   if (OkG) and (Shift = [ssCtrl]) then
      BEGIN
      Key := 0 ;
      bDelLigneClick(Nil) ;
      END ;
   END else
if Key = VK_RETURN then
   BEGIN
   END else ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.FormShow(Sender: TObject);
Var Ind : Integer ;
BEGIN
Modifie := False ;
TOB_Corr := TOB.Create('Table de correspondance', Nil, -1) ;
// Mise en place des colonnes de la grille
COL_CodeGpao := GS.FixedCols ;
ColDebut := GS.FixedCols + 2 ;
NbCol := 1 ;
GS.ColFormats[COL_CodeGpao] := 'UPPER' ;
for Ind := (ColDebut + NbCol) to GS.ColCount -1 do
   BEGIN
   GS.ColLengths[Ind] := -1 ;
   GS.ColWidths[Ind] := -1 ;
   END ;
AffecteGrid(GS, Action) ;
// Recherche de la tablette associée au champ
AdapteGrille ;
Case Action of
   taCreat   : BEGIN
               for Ind := 1 to NbRowsInit do InsereTOBCorr(Ind) ;
               GS.RowCount := NbRowsInit ;
               END ;
   taConsult : BEGIN
               ChargeTableCorrespondance ;
               BValider.Enabled := False ;
               bNewligne.Enabled := False ;
               bDelLigne.Enabled := False ;
               END ;
   taModif   : BEGIN
               ChargeTableCorrespondance ;
               CompleteTOBCorr ;
               END ;
   END ;
// Retaillage des contrôles de saisie
if NbCol <= 2 then VALDEFAUTPGI1.Width := VALDEFAUTGPAO.Width ;
if NbCol = 2 then
   BEGIN
   VALDEFAUTPGI2.Width := LIBDEFAUTGPAO.Width ;
   VALDEFAUTPGI2.Left  := LIBDEFAUTGPAO.Left ;
   END ;
if (NbCol < 5) and (PENTETE.Height > 100) then PENTETE.Height := PENTETE.Height - 28 ;
HMTrad.ResizeGridColumns(GS) ;
BVoirTob.Visible := V_PGI.SAV ;
// Appel de la fonction d'empilage dans la liste des fiches
AglEmpileFiche(Self) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCloseQuery
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.FormDestroy(Sender: TObject);
BEGIN
// Appel de la fonction de dépilage dans la liste des fiches       
AglDepileFiche ; 
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCloseQuery
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
Var Rep : Integer ;
BEGIN
if Modifie then
   BEGIN
   Rep := PGIAskCancel('Voulez-vous enregistrer les modifications ?', Caption) ;
   case Rep of
      mrYes : CanClose := (ValideSaisie) ;
      mrNo  : ;
      else    CanClose := False ;
      END ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  FormClose
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.FormClose(Sender: TObject; var Action: TCloseAction);
BEGIN
if TOB_Corr <> Nil then
   BEGIN
   TOB_Corr.Free ;
   TOB_Corr := Nil ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BFermeTestClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BFermeTestClick(Sender: TObject);
BEGIN
TWTest.Visible := False ;
ActiveBoutons(True) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BValideTestClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BValideTestClick(Sender: TObject);
Var Stg     : String ;
    Ind     : Integer ;
    Ctrl    : TComponent ;
    Valeurs : Array of String ;
    TOBPrinc    : TOB ;
BEGIN
// constitution de la TOB de la table de correspondance avec les valeurs par défaut
TOBPrinc := TOB.Create ('La table de correspondance', Nil, -1) ;
TOBPrinc.Dupliquer(TOB_Corr, True, True) ;
DepileTOBCorr(TOBPrinc) ;
AjoutValeurDefaut(TOBPrinc) ;
// lancement de la fonction de traduction
if VALTESTGPAO.Enabled then
   BEGIN
   // Recherche des valeurs PGI correspondantes
   Stg := TraduitGPAOenPGI(GRO_TYPEREPRISE.Value, GRO_NOMCHAMP.Text, VALTESTGPAO.Text, TOBPrinc, True) ;
   for Ind := 1 to NbCol do
      BEGIN
      Ctrl := FindComponent('VALTESTPGI' + IntToStr(Ind)) ;
      if (Ctrl <> Nil) and (Ctrl is THCritMaskEdit) then THCritMaskEdit(Ctrl).Text := ReadTokenSt(Stg) ;
      END ;
   END else
   BEGIN
   // Recherche la valeur GPAO correspondante
   SetLength(Valeurs, NbCol) ;
   for Ind := 1 to NbCol do
      BEGIN
      Ctrl := FindComponent('VALTESTPGI' + IntToStr(Ind)) ;
      if (Ctrl <> Nil) and (Ctrl is THCritMaskEdit) then Valeurs[Ind -1] := THCritMaskEdit(Ctrl).Text ;
      END ;
   Stg := TraduitPGIenGPAO(GRO_TYPEREPRISE.Value, GRO_NOMCHAMP.Text, Valeurs, True, TOBPrinc, True) ;
   VALTESTGPAO.Text := ReadTokenSt(Stg) ;
   TVALTESTGPAO.Caption := ReadTokenSt(Stg) ;
   END ;
TOBPrinc.Free ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BtestGPAOenPGIClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BtestGPAOenPGIClick(Sender: TObject);
BEGIN
ActiveChampTest(True) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BtestPGIenGPAOClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.BtestPGIenGPAOClick(Sender: TObject);
BEGIN
ActiveChampTest(False) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  ActiveChampTest : active les champs de saisie du module de test
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.ActiveChampTest ( GPAOenPGI : Boolean ) ;
Var Ind    : Integer ;
    Ctrl   : TComponent ;
    sTitre : String ;
BEGIN
TWTest.Visible := True ;
if TWTest.CanFocus then TWTest.SetFocus ;
// mise en forme du titre
Ind := Pos(' :', TWTest.Caption) ;
if Ind > 0 then
   BEGIN
   sTitre := Copy(TWTest.Caption, 1, Ind +1) ;
   if GPAOenPGI then TWTest.Caption := sTitre + TraduireMemoire(' GPAO vers PGI')
                else TWTest.Caption := sTitre + TraduireMemoire(' PGI vers GPAO') ;
   END ;
// inactivation des boutons de la fiche principale
ActiveBoutons(False) ;
// activation des champs de saisie et d'affichage des valeurs pour le test
VALTESTGPAO.Enabled := (GPAOenPGI) ;
VALTESTGPAO.Text := '' ;
TVALTESTGPAO.Caption := '' ;
for Ind := Low(NomTable) to High(NomTable) do
   BEGIN
   Ctrl := FindComponent('VALTESTPGI' + IntToStr(Ind)) ;
   if (Ctrl <> Nil) and (Ctrl is THCritMaskEdit) then
      BEGIN
      THCritMaskEdit(Ctrl).Text := '' ;
      THCritMaskEdit(Ctrl).Enabled := (not GPAOenPGI) ;
      THCritMaskEdit(Ctrl).Visible := (Ind <= NbCol) ;
      THCritMaskEdit(Ctrl).ElipsisButton := (not GPAOenPGI) ;
      if (not GPAOenPGI) and (Ind <= NbCol) then THCritMaskEdit(Ctrl).DataType := NomTable[Ind] ;
      END ;
   END ;
if NbCol = 1 then VALTESTPGI1.Width := VALTESTGPAO.Width ;
if GPAOenPGI then
   BEGIN
   if VALTESTGPAO.CanFocus then VALTESTGPAO.SetFocus ;
   END else
   BEGIN
   if VALTESTPGI1.CanFocus then VALTESTPGI1.SetFocus ;
   END
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  TWTestClose
///////////////////////////////////////////////////////////////////////////////////////
procedure TFCorrespGPAO.TWTestClose(Sender: TObject);
BEGIN
ActiveBoutons(True) ;
END ;

end.
