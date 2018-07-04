{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/07/2003
Modifié le ... : 10/02/2004
Description .. :
Suite ........ : Contient la définition des types et de la classe
Suite ........ : TTraitementCloture utilisés pour effectuer le
Suite ........ : traitement de la cloture comptable aussi bien en mode
Suite ........ : manuel qu'en process serveur
Suite ........ :
Suite ........ : 17/11/2003 : Ajout des fonctions d'annulation de cloture
Suite ........ : 10/02/2004 : JP : suppression des champs euro
Mots clefs ... : CLOTURE;
*****************************************************************}
unit ULIBCLOTURE;
 
interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    forms,
    ENT1,
    HEnt1,
    utoB,
{$IFDEF EAGLCLIENT}
{$ELSE}
    DB,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
{$ENDIF}
    SaisUtil,
    SaisComm,    // pour le TSAJAL
    ParaClo,
    SoldeCpt,
    UiUtil,
    Paramsoc,
    HCompte,        // pour le TGGeneral
    HStatus,        // pour le MoveCur, FiniMove, InitMove...
    HCtrls,         // pour le OpenSQL
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ENDIF MODENT1}
    utilPGI   ,UentCommun        // CanCloseExoFromGC
    ;

//==================================================
// Definition de Type
//==================================================
Type tPointVent = ( BQPointable, BQPointableVentilable, AutrePointable,
                    AutrePointableVentilable, PVRien);

Type
    ttDCClo = Record
        D : double ;
        C : double ;
    end;

Type
    TFiltreClo = Record
        // Données exercice à cloturer
        CloExo            : TExoDate;
        // Données Devise en cours de traitement
        CloDev            : String3;
        CloDec            : Integer;
        CloQuotite        : double;
        CloLibDevise      : String;
        // Données Etablissement en cours de traitement
        CloEtab           : String3;
        CloLibEtab        : String;
        // Données pour écritures générées
        JalGenere         : String3;
        ExoGenere         : String3;
        QualifGenere      : String3;
        EcrANouveauGenere : String3;
        TypGenere         : TTypeEcr;
        DateGenere        : TDateTime;
        ModePaieGenere    : String3;
        RefGenere         : String35;
        LibGenere         : String35;
        // Cumuls crédit / débit dans les différentes devises
        tot               : ttDCClo;
        totDev            : ttDCClo;
        // ...???
        EnSolde           : Boolean;
        Contrepartie      : ttContrepartie;
        CptContrepGenere  : String;
        CptInit           : String;
        CreerAna          : Boolean;
        // Indicateurs ???
        OkGenere          : Boolean;
        ModeGenerePiece   : ttmodegenerepiece;
        // Données relatives au journal
        SAJAL             : TSAJAL;
    end;


// ------------------------------------------
// - Code erreurs utilisés par les méthodes -
// ------------------------------------------
// Rem :
// - ERR = Erreur
// - WNG = Warning (pas bloquant)


const CLO_PASERREUR          =  -1 ;
      CLO_CLOTUREOK          =  38 ;
      CLO_ERRCLOTURE         =  44 ;
      CLO_ERRPROCESSSERVER   =  63 ;
      CLO_ERRCLOTUREIFRS     =  66 ;

// Pour la vérif des paramètres
      CLO_ERRCPTBILOUV       =  16 ;
      CLO_ERRCPTBENOUV       =  17 ;
      CLO_ERRCPTPEROUV       =  18 ;
      CLO_ERRJALOUV          =  19 ;
      CLO_ERRCPTBILFER       =  20 ;
      CLO_ERRCPTBENFER       =  21 ;
      CLO_ERRCPTPERFER       =  22 ;
      CLO_ERRJALFER          =  23 ;
      CLO_ERRCPTRES          =  24 ;
      CLO_WNGCOMPTEURJALCLO  =  32 ;  // Attention au compteur du journal de clôture : des journaux autres que celui de clôture ont le même compteur. Confirmez-vous ?
      CLO_WNGCOMPTEURJALANO  =  32 ;  // Attention au compteur du journal d'A-Nouveau : des journaux autres que celui d'A-Nouveau ont le même compteur. Confirmez-vous ?

// Erreur avant traitement
      CLO_EXO2NONOUVERT        = 45 ;    // Exo suivant non ouvert
      CLO_ERRBALEXO1           = 46 ;    // Balance de l'exo à cloture non équilibré
      CLO_ERRBALEXO1IFRS       = 53 ;    // Balance IFRS de l'exo à cloture non équilibré
      CLO_WNGNATCPT6           = 60 ;    // Certains comptes 6 ne sont pas de nature charge.
      CLO_WNGNATCPT7           = 61 ;    // Certains comptes 7 ne sont pas de nature produit.
      CLO_WNJALEXTRACOMPTABLE  = 63 ;    // Erreur si Journaux Extra-Comptable mouvementés
      CLO_WNGENEEXTRACOMPTABLE = 64 ;    // Erreur si Generaux Extra-Comptable non soldés
      CLO_WNCPTBILAN           = 65 ;    // Erreur de compte de Bilan de nature "CHA" ou "PRO"

      CLO_ERRIMMONONCLO      =  52 ;    // La clôture des immobilisations a-t-elle été faite?
      CLO_ECRNOVALIDE        =  67 ;    // Message d'alerte si présence d'écritures non validées sur l'exercice
// Erreur de décloture
      CLO_ERREXOPASCLO       =  02 ;    // Aucune clôture comptable n''a été faite sur le produit
      CLO_ERRANNULCLO        =  05 ;    // ATTENTION ! Problème pendant l'annulation de clôture
      CLO_ERRLETTANO         =  12 ;    // Vous devez supprimer auparavant le lettrage et le pointage sur les à-nouveaux.
{-------------- pas utile pour le moment
      CLO_INEXISTANT         =  25 ;
      CLO_COLLECTIF          =  26 ;
      CLO_LETTRABLE          =  27 ;
      CLO_POINTABLE          =  28 ;
      CLO_ERRNATURE          =  29 ;  // de nature incorrecte
      CLO_NOFACTASSO         =  30 ;  // : pas de facturier associé
----------------}

//==================================================
// Definition de la classe TTraitementCloture
//==================================================
type
    TTraitementCloture = class
      private
        // Nouveaux Attributs
        Ecran         : TForm ;
        Auto          : Boolean;          // mode auto -> mode silencieux...
        modeServeur   : Boolean;          // mode serveur -> déroulement autonome...
        modeIFRS      : Boolean;          // mode IFRS -> Action sur les écritures "I"...
        modeAnoDyna   : boolean ;

        // Exercice à cloturer et suivant
        Exo1         : tExoDate;
        Exo2         : tExoDate;

        // Liste des paramètres d'ouverture
        StOuvRef : String ;
        StOuvLib : String ;
        StOuvJal : String ;
        StOuvBil : String ;
        StOuvPer : String ;
        StOuvBen : String ;
        DtOuvDeb : TDateTime ;
        DtOuvFin : TDateTime ;

        // Liste des paramètres de fermeture
        StFermRef : String ;
        StFermLib : String ;
        StFermJal : String ;
        StFermBil : String ;
        StFermPer : String ;
        StFermBen : String ;
        StFermRes : String ;
        DtFermDeb : TDateTime ;
        DtFermFin : TDateTime ;

        FStLastErrorMsg : string;

        // Attributs utilisés pour le traitement de cloture
        LLEtabDev     : TStringList;      // Liste des couples devise / etablissement
        CloDef        : Boolean;          // indicateur de cloture définitive

        ParaClo       : TParamCloture;    // Paramètres de cloture (cf paraclo.pas)

        ExisteCptPV   : Boolean;          // indicateur d'existance des comptes pointables et ventilables
        LCptBQE       : TStringList;      // liste des comptes de banque
        LCptPV        : TStringList;      // liste des comptes pointables et ventilbales

// Initialisation
        Procedure Initialisation ;
        Procedure ReinitPourIFRS ;
        Procedure AlimParamDefaut ;
        Procedure AffectParamDepuisTOB ( TobParam : TOB ) ;
// Fonctions utiles
        Procedure ChercheCpt(Gen,Aux : String ; var EtatLettrage,Ana,Eche : String ; var NumEche : Integer ; var Coll : Boolean ; var PointVent : TPointVent ; var Setfb : TsetFichierBase);
        Function  PositionneStUpdate( L : TStringList ) : String;
        function  PbSurMontantDevise (What : Integer ; Deb,Cre,DebDev,CreDev : double ; CloDec : Integer) : Boolean;
        function  ModePaiementDefaut : string;
        procedure QuelParam(var F,F2 : tFiltreClo);
        Function  IsPointableATraiterCommeVentilable(Cpt :String) : Boolean;
        procedure AlimListeCpt(PourBanque : Boolean);
        Function  ChangeEcartDeChange(Sens : integer) : Boolean;
        Function  PresenceLettragePointage: boolean; // Ajout 17/11/2003
// Fonctions Traitements
        Procedure DeleteMouvementResultat(var F : TFiltreClo);
        procedure majExo ;
        Procedure InsertMvtResultat(var F : TFiltreClo ; Cpt : String);
        // pour l'analytique
        function  AddMvtAna(What : Byte ; F : TFiltreClo ; Gen,Sect : String17 ; Deb,Cre,DebDev,CreDev: double ; var NumLig,NumVentil : LongInt ; NumPiece : LongInt ; SensEcr : Byte ; Axe : String ; totMnt,totMntDev: double; splan1 : string = ''; splan2 : string = ''; splan3 : string = ''; splan4 : string = ''; splan5 : string = '') : Boolean;
        procedure CalculANCLOANA(F : TFiltreClo ; Gen,Aux : String ; Sens : Byte ; NumPiece,NumLig : LongInt ; totMnt,totMntDev: double);
        // pour les écritures générales
        function  AddMvt(What : Byte ; F : TFiltreClo ; Gen,Aux : String17 ; var Deb,Cre,DebDev,CreDev: double ; CloDev,EtatLettrage,Ana,Eche,ModePaie : String ; var NumLig : LongInt ; NumPiece,NumEche : LongInt ; forceRef,forceLib : String ; CreerAnaSurAttente : Boolean ; Setfb : TsetFichierBase) : Boolean;
        procedure CalculANCLO(var F,F2 : TFiltreClo ; Recopie : Boolean ; QEcr : TQuery ; PourPieceCloture : Boolean ; TraitePointage : Boolean);
        procedure TraiteInsertAnClo (var F : TFiltreClo ; Gen,Aux : String ; Deb,Cre,DebDev,CreDev: double ; var NumPiece : LongInt ; var OkInsert : Boolean ; var NumLig : LongInt ; var totDeb,totCre,totDebDev,totCreDev: double ; PourPieceCloture : Boolean ; var SOldColl : String ; RefP : String);

        // Uniquement pour cloture définitive
        procedure AnnuleRefPointageAnalytique(Sens : integer ; StCompteUpdate : String);
        procedure ChangeRefPointage(Sens : integer ; StCompteUpdate : String);

        // Pour l'annulation de la cloture (Ajout 17/11/2003)
        function DeleteANO : Integer ;
        function DeleteCLO : Integer ;
        function DeleteANOH: Integer;
        procedure MajExoAnnulClo ;
        procedure MajExoV8;

// Init TFiltreClo / Ecran suivant étape
        procedure InitQCHAPRO(Nat : String3 ; var F : TFiltreClo);
        procedure InitGeneral(var F : TFiltreClo ; I : Integer);
        procedure InitSupp(var F : TFiltreClo ; What : Byte);
        procedure InitQResultatClo(var F : TFiltreClo ; totResultat,totResultatDev: ttDCClo ; Cpt : String);
        procedure InitQResultatAN(var F : TFiltreClo ; totResultat,totResultatDev: ttDCClo ; Cpt : String);
// Intéraction écran
        function  GestionEcran : Boolean ;
        procedure EcranCursorSynchr ;
        procedure EcranMessage1( i : Integer ) ;
        procedure EcranMessage2( i : Integer ; St1, St2 : String) ;
        procedure EcranInitGeneral;
        procedure EcranInitMove ( vInQte : Integer ; vStTitre : String ) ;
        procedure EcranMoveCur ;
        procedure EcranFiniMove ;
        procedure EcranUpdateXX ( vStCaption : String ) ;
        procedure EcranUpdateYY ( vStCaption : String ) ;
        procedure EcranEtapeSuivante( vInEtape : Integer ) ;
// NOUVEAU TRAITEMENT BASE
        procedure UpdateTiersSurAna ;
//----------
        procedure ClotureGCD ;
      PUBLIC
//----------
        property LastErrorMsg : string read FStLastErrorMsg;

        // Constructeur / Destructeur
        Constructor Create ( vEcran : TForm; vStCodeExo1, vStCodeExo2 : String ; vBoCloDef, vBoAuto : Boolean ; vAnoDyna : boolean = false ) ;
        Destructor  Destroy ; override ;
        // Affectation
        Procedure SetParamCloture ( vParamClo : TParamCloture ) ;
        Procedure SetParamFerm ( vStRef, vStLib, vStJal, vStBil, vStPer, vStBen, vStRes : String ; vDtDeb, vDtFin : TDateTime ) ;
        Procedure SetParamOuv ( vStRef, vStLib, vStJal, vStBil, vStPer, vStBen : String ; vDtDeb, vDtFin : TDateTime ) ;
        Procedure SetModeServeur ( vBoMode : Boolean ) ;
        // Lecture
        Function  GetExisteCptPV : boolean ;
        Function  EstModeIFRS : boolean ;
        Function  AuMoinsUneIFRS : Boolean ;

        // Vérifications
        Function  BalanceOk            : Integer ;
        Function  JalExtraComptableOk  : Integer ; // GCO - FQ 13049
        Function  GeneExtraComptableOk : Integer ; // GCO - FQ 13049
        Function  CompteBilanOk        : Integer ; // GCO - FQ 18842
        Function  CompteChargeOk       : Integer ;
        Function  CompteProduitOk      : Integer ;
        Function  VerifParamOk         : Integer ;
        Function  CanCloseExoGC        : boolean ; // SBO DEV3228 FQ 16676

        // Fonction utiles
        Function GetQteEtabDev : Integer ;
        Function NbEtabDev : Boolean;
        Function CreerTobParam : TOB ;
        Function CreerTobParamCloture : TOB ;
        // Traitements
        procedure DetruitANO;
        procedure TraitementPVBQE(i : Integer);
        function  Cloture  ( var ClotureOk, OkTraite : Boolean ) : Integer;
        procedure VerifCloture ( var OnSort : Boolean ) ;
        function  AnnuleCloture ( var OnSort : Boolean ) : Integer ;
        // Procedure pour le process server
        Function  ExecuteCloture ( TobParam : TOB ) : Tob;
        function  ExecuteAnnuleCloture ( TobParam : TOB ) : Tob ;
        Function  EcritValidOk: Integer ;
    end;


//================================================
//========= Fonctions externes ===================
//================================================
Function VerifCpt(Cpt : String) : Integer;
Function VerifJal(Jal,NatJal : String ; var Facturier : String) : Integer;
Function Multifacturier( NatJal, FactJal : String ) : Boolean;
procedure TraitementSituationDynamique;
// GCO - 04/09/2006 - FQ 18662
// Pour la mise à jour de la validation des journaux lors de la cloture
procedure MajValidationJournaux( vBoCloture : Boolean );

implementation

uses
  {$IFDEF MODENT1}
  CPVersion, 
  CPProcMetier,
  CPProcGen,
  ULibExercice,
  {$ENDIF MODENT1}
  {$IFNDEF EAGLSERVER}
  Cloture, // pour TFClos
  {$ENDIF}
  ULibAnalytique,
  balsit,
  uLibWindows; // IIF


// Pb initialisation des variables en EAGL-Serveur FQ 15943 SBO 01/08/2005
Const  FOkDecV       : Integer = 2 ;
       FDevisePivot  : String  = 'EUR' ;
       FCodeSociete  : String  = '' ;

//==================================================
// Autre Fonction hors class
//==================================================

function Multifacturier( NatJal, FactJal: String ) : Boolean ;
begin
  result := ExisteSQL('SELECT J_JOURNAL,J_NATUREJAL,J_COMPTEURNORMAL FROM JOURNAL WHERE J_NATUREJAL<>"'+NatJal+'" AND J_COMPTEURNORMAL="'+FactJal+'"');
end;

Function StAxe(fb : TFichierBase) : String3;
begin
    StAxe := 'A'+InttoStr(ord(fb)+1);
end;

Function VerifCpt(Cpt : String) : Integer;
var CGen : TGGeneral;
begin
  CGen := TGGeneral.Create(Cpt);
  Result := 25;
  if (Cgen <> nil) then
    begin
    if (CGen.General = '')
      then Result := 25
      else if (Cgen.Lettrable)
        then Result := 27
        else if (CGen.Collectif)
          then Result := 26
          else if (CGen.Pointable)
            then Result := 28
            else if (Cgen.Ventilable[1] or Cgen.Ventilable[2] or Cgen.Ventilable[3] or Cgen.Ventilable[4] or Cgen.Ventilable[5])
              then Result := 29
              else Result := 0;
    CGen.Free ;
    end;
end;

Function VerifJal(Jal,NatJal : String ; var Facturier : String) : Integer;
var
    SAJAL : TSAJAL;
begin
    SAJAL := TSaJal.Create(Jal,false);
    Result := 25;
    facturier := '';

    if (SAJAL <> nil) then
    begin
        if (SAJAL.JOURNAL <> JAL)           then Result := 25
        else if (SAJAL.NATUREJAL <> NatJAL) then result := 30
        else if (SAJAL.COMPTEURNorMAL = '') then Result := 31
        else result := 0;

        Facturier := SAJAL.COMPTEURNorMAL;
        SAJAL.Free;
    end;
end;

procedure SigneMoins(var M,M1{,M2} : double);
begin
    M  := -1*M;
    M1 := -1*M1;
//    M2 := -1*M2;
end;

Procedure ChangeF(var F : TFiltreClo);
begin
    if (F.ModeGenerePiece = UneParCpt) then F.Contrepartie := ChaqueLigne;
end;

procedure Alimtot(var tot,totDev : ttDcClo ; var F : TFiltreClo ; Inverse : Boolean);
begin
    if (Inverse) then
    begin
        tot.D := F.tot.C;
        tot.C := F.tot.D;
        totDev.D := F.totDev.C;
        totDev.C := F.totDev.D;
    end
    else
    begin
        tot.D := F.tot.D;
        tot.C := F.tot.C;
        totDev.D := F.totDev.D;
        totDev.C := F.totDev.C;
    end;
end;

Function CalculResultat(var totCha,totChaDev,totPro,totProDev,totResultat,totResultatDev : ttDCClo ; Inverse : Boolean ; var F : TFiltreClo) : Boolean;
var
    D,C,DDev,CDev : double;
begin
    Result := false;
    totResultat.D := 0;
    totResultat.C := 0;
    totResultatDev.D := 0;
    totResultatDev.C := 0;
    D := Arrondi(totCha.D+totPro.D,FOkDecV);
    C := Arrondi(totCha.C+totPro.C,FOkDecV);
    DDev := Arrondi(totChaDev.D+totProDev.D,F.CloDec);
    CDev := Arrondi(totChaDev.C+totProDev.C,F.CloDec);

    if (Arrondi(D-C,FOkDecV) = 0) then
    begin
        if ((Arrondi(DDev-CDev,F.CloDec) <> 0)) then
        begin
            if (CDev > DDev) then
            begin
                Result := true;
                if (Inverse) then
                begin
                    totResultat.D := Arrondi(C-D,FOkDecV);
                    totResultatDev.D := Arrondi(CDev-DDev,F.CloDec);
                end
                else
                begin
                    totResultat.C := Arrondi(C-D,FOkDecV);
                    totResultatDev.C := Arrondi(CDev-DDev,F.CloDec);
                end;
            end
            else
            begin
                Result := false;
                if (Inverse) then
                begin
                    totResultat.C := Arrondi(C-D,FOkDecV);
                    totResultatDev.C := Arrondi(CDev-DDev,F.CloDec);
                end
                else
                begin
                    totResultat.D := Arrondi(D-C,FOkDecV);
                    totResultatDev.D := Arrondi(DDev-CDev,F.CloDec);
                end;
            end;
        end;
    end
    else
    begin
        if (C > D) then
        begin
            Result := true;
            if (Inverse) then
            begin
                totResultat.D := Arrondi(C-D,FOkDecV);
                totResultatDev.D := Arrondi(CDev-DDev,F.CloDec);
            end
            else
            begin
                totResultat.C := Arrondi(C-D,FOkDecV);
                totResultatDev.C := Arrondi(CDev-DDev,F.CloDec);
            end;
        end
        else
        begin
            Result := false;
            if (Inverse) then
            begin
                totResultat.C := Arrondi(C-D,FOkDecV);
                totResultatDev.C := Arrondi(CDev-DDev,F.CloDec);
            end
            else
            begin
                totResultat.D := Arrondi(D-C,FOkDecV);
                totResultatDev.D := Arrondi(DDev-CDev,F.CloDec);
            end;
        end;
    end;
end;

// fonction de 'CPTESAV' et 'Souche' reprisse parce que ces unité ne sont pas eAGL ...

function RecalculSoucheSurUnExo(StWhereJal,Souche : String ; Exo : TExoDate ; Ecrit : Boolean) : Integer;
var
    Num : Integer;
    SQL : String;
    QS1 : TQuery;
begin
    Num := 0;

    SQL := 'SELECT MAX(E_NUMEROPIECE) FROM ECRITURE WHERE '+StWhereJal+' AND E_QUALIFPIECE="N"';
    if (Exo.Code <> '') then SQL := SQL+' AND E_EXERCICE="'+Exo.Code+'" ';

    QS1 := OpenSql(Sql,true);
    if (Not QS1.Eof) then Num := QS1.Fields[0].AsInteger;
    Ferme( QS1 ) ;

    Inc(Num) ;

    if (Ecrit) then
    begin
        SQL:='' ;

        if (Exo.Code = '') then SQL:='UPDATE SOUCHE SET SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_SOUCHE="'+Souche+'"'
        else if (Exo.Code = VH^.Suivant.Code) then SQL:='UPDATE SOUCHE SET SH_NUMDEPARTS='+IntToStr(Num)+' WHERE SH_SOUCHE="'+Souche+'"'
        else if (Exo.Code = VH^.EnCours.Code) then SQL:='UPDATE SOUCHE SET SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_SOUCHE="'+Souche+'"'
        else if (Exo.Code = VH^.Precedent.Code) then SQL:='UPDATE SOUCHE SET SH_NUMDEPARTP='+IntToStr(Num)+' WHERE SH_SOUCHE="'+Souche+'"';

        if (SQL <> '') Then ExecuteSQL(SQL);
    end;
    Result := Num;
end;

procedure RecalculSouchePourCloture;
var
    Jal,Souche,StWhereJal : String;
    QS1,QS2 : TQuery;
    Exo : TExoDate;
begin
    if (not VH^.MultiSouche) then exit;

    try
        BeginTrans;

        QS1 := OpenSql('SELECT * FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_ANALYTIQUE="-" AND SH_SIMULATION="-" AND SH_SOUCHEEXO="X"',True);

        while (not QS1.Eof) do
        begin
            Souche := QS1.FindField('SH_SOUCHE').AsString;
            Fillchar(Exo,SizeOf(Exo),#0);

            QS2 := OpenSql('SELECT J_JOURNAL FROM JOURNAL WHERE J_COMPTEURNORMAL="'+souche+'"',True);

            StWhereJal := '';

            while (not QS2.Eof) do
            begin
                Jal := QS2.FindField('J_JOURNAL').AsString;
                if (StWhereJal = '') then StWhereJal := 'E_JOURNAL="'+Jal+'" '
                else StWhereJal := StWhereJal+' OR E_JOURNAL="'+Jal+'" ' ;
                QS2.Next;
            end;

            Ferme( QS2 ) ;

            if (StWhereJal <> '') then
            begin
                StWhereJal:='('+StWhereJal+') ' ;
                RecalculSoucheSurUnExo(StWhereJal,Souche,VH^.Precedent,true);
                RecalculSoucheSurUnExo(StWhereJal,Souche,VH^.EnCours,true);
                RecalculSoucheSurUnExo(StWhereJal,Souche,VH^.Suivant,true);
            end;
            QS1.Next;
        end;

        Ferme( QS1 );
        CommitTrans;
    except
        Rollback;
    end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TraitementSituationDynamique;
begin
  if GetParamSocSecur('SO_CPBDSMENSUELLE', False) then
    CreationBDSDynamique('M');

  if GetParamSocSecur('SO_CPBDSTRIMESTRIELLE', False) then
    CreationBDSDynamique('T');

  if GetParamSocSecur('SO_CPBDSSEMESTRIELLE', False) then
    CreationBDSDynamique('S');

  if GetParamSocSecur('SO_CPBDSANNUELLE', False) then
    CreationBDSDynamique('A');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 24/04/2007
Modifié le ... :   /  /
Description .. : Fonction utilisée par la compta en cas de mise à jour des
Suite ........ : dates d'exercice. Positionne le stock initial au stock de la
Suite ........ : date de début d'exercice.
Mots clefs ... :
*****************************************************************}
function UpdateStockInitial(Const DebutExercice: tDateTime): Boolean;
var
  sql: string;
begin
{$IFNDEF CLO23}
  sql := 'UPDATE DISPO'
       + ' SET GQ_STOCKINITIAL=(SELECT SUM(GSM_PHYSIQUE)'
       +                      ' FROM STKMOUVEMENT'
       +                      ' WHERE GSM_DEPOT=GQ_DEPOT'
       +                      ' AND GSM_ARTICLE=GQ_ARTICLE'
       +                      ' AND GSM_STKTYPEMVT IN ("000","PHY")'
       +                      ' AND GSM_DATEMVT<"' + UsDateTime(DebutExercice) + '")'
       + ' WHERE GQ_CLOTURE="-"'
       ;
  ExecuteSql(Sql);
{$ENDIF}
  Result := V_Pgi.iOError = oeOk;

end;

////////////////////////////////////////////////////////////////////////////////
//==================================================
//     Méthodes de la classe TTraitementCloture
//==================================================

{ TTraitementCloture }

function TTraitementCloture.AddMvt(What: Byte; F: TFiltreClo; Gen,  Aux: String17; var Deb, Cre, DebDev, CreDev: double;  CloDev, EtatLettrage, Ana, Eche, ModePaie: String; var NumLig: Integer;  NumPiece, NumEche: Integer; forceRef, forceLib: String;  CreerAnaSurAttente: Boolean; Setfb: TsetFichierBase): Boolean;
var OnDev                       : Boolean;
    Debit,Credit                : double;
    DebitDev,CreditDev          : double;
    TauxDev, Coeff              : double;
    Mnt,MntDev                  : double;
    totAna,totAnaDev            : double;
    LeDebit,LeCredit,LeMnt      : double;
    Sens                        : Byte;
    fb                          : tfichierbase;
    NumVentil                   : LongInt;
    Sect                        : String17;
    QAdd                        : TQuery;
    splan1, splan2, splan3, splan4, splan5 : string;
 {   lTOB : TOB ;
    lInfo : TInfoEcriture ;
    lRecError : TRecError ;  }
begin

//  MAJ affichage de l'interface
  EcranMessage2(1,Gen,Aux);

  Result := false;

  OnDev := (F.CloDev <> FDevisePivot);

{  lTOB := TOB.Create('',nil,-1) ;
  lInfo := TInfoEcriture.Create;   }

  // Init montants et var calcul
  Debit       := 0;
  Credit      := 0;
  DebitDev    := 0;
  CreditDev   := 0;
  totAna      := 0;
  totAnaDev   := 0;
  TauxDev     := 1;
  Mnt         := 0;
  MntDev      := 0;
  Sens        := 1;

  // Arrondi des paramètre au cas ou ...
  Deb     := Arrondi(Deb,FOkDecV);
  Cre     := Arrondi(Cre,FOkDecV);
  DebDev  := Arrondi(DebDev,F.CloDec);
  CreDev  := Arrondi(CreDev,F.CloDec);

  // alimentation des montants en fonction du type d'écriture
  Case What Of
    // débit
    1 : begin
        Sens      := 1;
        Debit     := Deb;
        DebitDev  := DebDev;
        if ((Arrondi(Debit,4) = 0) and (Arrondi(DebitDev,4) = 0)) then
          Exit;
        totAna     := Debit;
        totAnaDev  := DebitDev;
        end;
    // crédit
    2 : begin
        Sens        := 2;
        Credit      := Cre;
        CreditDev   := CreDev;
        if ((Arrondi(Credit,4) = 0) and (Arrondi(CreditDev,4) = 0)) then
          Exit;
        totAna      := Credit;
        totAnaDev   := CreditDev;
        end;
    // en solde
    3 : begin
        if (deb > Cre) then
          begin
          Sens := 1;
          Debit := Deb-Cre;
          DebitDev := DebDev-CreDev;
          Deb := Debit;
          DebDev := Debitdev;
          Cre := 0;
          CreDev := 0;
          if ((Arrondi(Debit,4) = 0) and (Arrondi(DebitDev,4) = 0)) then
            begin
            Deb := 0;
            DebDev := 0;
            Exit;
            end;
          Mnt := Debit;
          MntDev := DebitDev;
          end
        else
          begin
          Sens := 2;
          Credit := Cre-Deb;
          CreditDev := CreDev-DebDev;
          Cre := Credit;
          CreDev := CreditDev;
          Deb := 0;
          DebDev := 0;
          if ((Arrondi(Credit,4) = 0) and (Arrondi(CreditDev,4) = 0)) then
            begin
            Cre := 0;
            CreDev := 0;
            Exit;
            end;
          Mnt := Credit;
          MntDev := CreditDev;
          end;

        totAna := Mnt;
        totAnaDev := MntDev;
        end;
    end;

  // cas spécial si monnaie de référence <> Euro
  Coeff := 1;
  LeDebit := Debit;
  LeCredit := Credit;
  LeMnt := Mnt;

  // détermination du taux de la devise de l'écriture
  if (not OnDev)
    then TauxDev := 1
    else
      Case What Of
        1 : if (DebitDev <> 0) then
              TauxDev := F.CLoQuotite * (LeDebit / DebitDev) * Coeff ;
        2 : if (CreditDev <> 0) then
              TauxDev := F.CLoQuotite * (LeCredit / CreditDev ) * Coeff ;
        3 : if (MntDev <> 0) then
              TauxDev := F.CLoQuotite * ( LeMnt / MntDev ) * Coeff ;
        end;

  // Pourkoi seulement maintenant ???
  if (not F.OkGenere) then
    Exit;

  Inc(NumLig);

  if (F.TypGenere = EcrClo) then
    begin
    ModePaie := '';
    NumEche := 0;
    end;

  // Requête bidon pour insertion base...
  Qadd := OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL="###"',false);
  // Nouvel enregistrement
  Qadd.Insert;
  InitNew(Qadd);

  // Remplissage des champs
  Qadd.FindField('E_EXERCICE').AsString         := F.ExoGenere;
  Qadd.FindField('E_JOURNAL').AsString          := F.JalGenere;
  Qadd.FindField('E_DATECOMPTABLE').AsDateTime  := F.DateGenere;
  Qadd.FindField('E_NUMEROPIECE').AsInteger     := NumPiece;
  Qadd.FindField('E_NUMLIGNE').AsInteger        := NumLig;
  Qadd.FindField('E_GENERAL').AsString          := Gen;
  Qadd.FindField('E_AUXILIAIRE').AsString       := Aux;
  Qadd.FindField('E_DEBIT').AsFloat             := Arrondi(Debit,FOkDecV);
  Qadd.FindField('E_CREDIT').AsFloat            := Arrondi(Credit,FOkDecV);
  Qadd.FindField('E_ENCAISSEMENT').asString     := SENSENC(Debit,Credit) ;

  if (forceRef <> '')
    then Qadd.FindField('E_REFINTERNE').AsString := forceRef
    else Qadd.FindField('E_REFINTERNE').AsString := F.RefGenere;

  if (forceLib <> '')
    then Qadd.FindField('E_LIBELLE').AsString   := forceLib
    else Qadd.FindField('E_LIBELLE').AsString   := F.LibGenere;

  Qadd.FindField('E_NATUREPIECE').AsString      := 'OD';
  Qadd.FindField('E_MODESAISIE').AsString       := '-';
  Qadd.FindField('E_TRESOSYNCHRO').AsString     := 'RIE';
  Qadd.FindField('E_QUALIFPIECE').AsString      := F.QualifGenere;
  Qadd.FindField('E_VALIDE').AsString           := 'X';
  Qadd.FindField('E_DATECREATION').AsDateTime   := date; // CA -22/04/2003 - FQ 12215
  Qadd.FindField('E_DATEMODIF').AsDateTime      := date; // CA -22/04/2003 - FQ 12215
  Qadd.FindField('E_SOCIETE').AsString          := FCodeSociete;
  Qadd.FindField('E_ETABLISSEMENT').AsString    := F.CloEtab;
  Qadd.FindField('E_MODEPAIE').AsString         := ModePaie;
  Qadd.FindField('E_DATEECHEANCE').AsDateTime   := F.DateGenere;
  Qadd.FindField('E_DEVISE').AsString           := F.CloDev;
  Qadd.FindField('E_DEBITDEV').AsFloat          := Arrondi(DebitDev,F.CloDec);
  Qadd.FindField('E_CREDITDEV').AsFloat         := Arrondi(CreditDev,F.CloDec);
  Qadd.FindField('E_TAUXDEV').AsFloat           := TauxDev;
  Qadd.FindField('E_ECRANOUVEAU').AsString      := F.EcrAnouveauGenere;
  Qadd.FindField('E_ETATLETTRAGE').AsString     := EtatLettrage;
  Qadd.FindField('E_NUMECHE').AsInteger         := NumEche;
  Qadd.FindField('E_ANA').AsString              := Ana;
  Qadd.FindField('E_ECHE').AsString             := Eche;
  Qadd.FindField('E_COTATION').AsFloat          := 1;
  Qadd.FindField('E_REGIMETVA').AsString        := GetParamSocSecur('SO_REGIMEDEFAUT','');
{$IFNDEF SPEC302}
    Qadd.FindField('E_PERIODE').AsInteger       := GetPeriode(F.DateGenere);
    Qadd.FindField('E_SEMAINE').AsInteger       := NumSemaine(F.DateGenere);
{$ENDIF}

  // MAj du pointage de la pièce
  {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
  if EstSpecif('51216') then begin
    if (CloDef) then
      if ((forceRef <> '') and (F.QualifGenere = 'N')) then
        begin
        Qadd.FindField('E_REFPOINTAGE').AsString := forceRef;
        Qadd.FindField('E_DATEPOINTAGE').AsDateTime := Exo1.Fin;
        end
      else
        begin
        Qadd.FindField('E_REFPOINTAGE').AsString := '';
        Qadd.FindField('E_DATEPOINTAGE').AsDateTime := IDate1900;
        end;
  end;

  if modeAnoDyna then
   QAdd.FindField('E_IO').AsString := '-'
    else
     QAdd.FindField('E_IO').AsString := 'X';


{  lTOB.selectdb('',QAdd) ;
  lInfo.LoadJournal(    lTob.GetValue('E_JOURNAL')           ) ; // Journal
  lInfo.Etabliss.load(  [ lTob.GetValue('E_ETABLISSEMENT') ] ) ; // Etablissement
  lInfo.Devise.load(    [ lTob.GetValue('E_DEVISE') ]        ) ; // Devise
  lInfo.LoadCompte( lTob.GetValue('E_GENERAL')    ) ; // Compte général
  lInfo.LoadAux(    lTob.GetValue('E_AUXILIAIRE') ) ; // compte auxiliaire
  lInfo.AjouteErrIgnoree([RC_DATEINCORRECTE]) ;
  lRecError := CIsValidLigneSaisie(lTob,lInfo) ;
  if ( lRecError.RC_Error <> RC_PASERREUR ) then
   begin
    Showmessage( intTostr(lRecError.RC_Error) + ' : '  + VartoStr(lTob.GetValue('E_GENERAL')) ) ;
    exit ;
   end ; }


  // Enregistrement en base
  Qadd.Post ;

  Ferme( QAdd ) ;


  Result := true;

  // Traitement de l'analytique associée
  if (F.CreerAna and (Ana = 'X')) then
    if (CreerAnaSurAttente) then
      begin
      if not VH^.AnaCroisaxe then
      begin
        NumVentil := 1;
        for fb := fbAxe1 to fbAxe5 do if (fb in Setfb) then
        begin
          Sect := VH^.Cpta[fb].Attente;
          AddMvtAna(3,F,Gen,Sect,Debit,Credit,DebitDev,CreditDev,NumLig,NumVentil,NumPiece,Sens,fbtoAxe(fb),totAna,totAnaDev);
        end;
      end
      else
      begin
        NumVentil := 1 ;
        Sect := VH^.Cpta[ AxeToFb('A' + IntToStr(RecherchePremDerAxeVentil.premier_axe)) ].Attente;
        splan1 := '' ; splan2 := '' ; splan3 := '' ; splan4 := ''; splan5 := '';
        if GetParamSocSecur('SO_VENTILA1',False) then splan1 := VH^.Cpta[AxeToFb('A1') ].Attente;
        if GetParamSocSecur('SO_VENTILA2',False) then splan2 := VH^.Cpta[AxeToFb('A2') ].Attente;
        if GetParamSocSecur('SO_VENTILA3',False) then splan3 := VH^.Cpta[AxeToFb('A3') ].Attente;
        if GetParamSocSecur('SO_VENTILA4',False) then splan4 := VH^.Cpta[AxeToFb('A4') ].Attente;
        if GetParamSocSecur('SO_VENTILA5',False) then splan5 := VH^.Cpta[AxeToFb('A5') ].Attente;
        AddMvtAna(3,F,Gen,Sect,Debit,Credit,DebitDev,CreditDev,NumLig,NumVentil,NumPiece,Sens,'A'+IntToStr(RecherchePremDerAxeVentil.premier_axe),totAna,totAnaDev, splan1, splan2, splan3, splan4, splan5);
      end;

      end
    else
      CalculANCLOAna(F,Gen,Aux,Sens,NumPiece,NumLig,totAna,totAnaDev) ;

end;

function TTraitementCloture.AddMvtAna(What: Byte; F: TFiltreClo; Gen, Sect: String17; Deb, Cre, DebDev, CreDev: double; var NumLig, NumVentil: Integer; NumPiece: Integer; SensEcr: Byte; Axe: String; totMnt, totMntDev: double; splan1 : string = ''; splan2 : string = ''; splan3 : string = ''; splan4 : string = ''; splan5 : string = ''): Boolean;
var OnDev,OnDeb             : Boolean;
    Debit, Credit           : double;
    DebitDev, CreditDev     : double;
    Mnt,MntDev              : double;
    TauxDev,Pourcent,Coeff  : double;
    SensAna                 : Byte;
    LeDebit,LeCredit,LeMnt  : double;
    QAddAna                 : TQuery;
begin

  Result := false;

  // Requête bidon pour insertion base
  QAddAna := OpenSQL('SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="###"',false);

  // ??
  OnDev := (F.CloDev <> FDevisePivot);

  // init montants
  Debit       := 0;
  Credit      := 0;
  DebitDev    := 0;
  CreditDev   := 0;
  Mnt         := 0;
  MntDev      := 0;
  // Init var de calcul
  Pourcent    := 100.0;
  TauxDev     := 1;

  // Les montants sont arrondis au cas ou...
  Deb     := Arrondi(Deb,FOkDecV);
  Cre     := Arrondi(Cre,FOkDecV);
  DebDev  := Arrondi(DebDev,F.CloDec);
  CreDev  := Arrondi(CreDev,F.CloDec);

  Case What Of
    // débit
    1 : begin
        SensAna   := 1;
        Debit     := Deb;
        DebitDev  := DebDev;
        if ((Arrondi(Debit,4) = 0) and (Arrondi(DebitDev,4) = 0)) then
          Exit;
        if (SensAna <> SensEcr) then
          SigneMoins(Debit,DebitDev);
        if (OnDev) then
          begin
          if (Arrondi(totMntDev,4) <> 0) then
            Pourcent := Arrondi(DebitDev/totMntDev*100,ADecimP);
          end
        else
          if (Arrondi(totMnt,4) <> 0) then
            Pourcent := Arrondi(Debit/totMnt*100,ADecimP);
        end;
    // crédit
    2 : begin
        SensAna    := 2;
        Credit     := Cre;
        CreditDev  := CreDev;
        if ((Arrondi(Credit,4) = 0) and (Arrondi(CreditDev,4) = 0)) then
          Exit;
        if (SensAna <> SensEcr) then
          SigneMoins(Credit,CreditDev);
        if (OnDev) then
          begin
          if (Arrondi(totMntDev,4) <> 0) then
            Pourcent := Arrondi(CreditDev/totMntDev*100,ADecimP);
          end
        else
          begin
          if (Arrondi(totMnt,4) <> 0) then
            Pourcent := Arrondi(Credit/totMnt*100,ADecimP);
          end;
        end;
    // en solde
    3 : begin
        // CA - 16/05/2002 - Cas des montants négatifs - FQ 12308
        if (abs (Deb) > abs (Cre) ) // remplace if (Deb > Cre) then
          then OnDeb := not ( (Cre-Deb < 0) and (SensEcr = 2) and (Arrondi(Cre-Deb-totMnt,4) = 0) )
          else OnDeb := ( (Deb-Cre < 0) and (SensEcr = 1) and (Arrondi(Deb-Cre-totMnt,4) = 0) ) ;

        if (OnDeb) then
          begin
          SensAna := 1;
          Debit := Deb-Cre;
          DebitDev := DebDev-CreDev;

          if ((Arrondi(Debit,4) = 0) and (Arrondi(DebitDev,4) = 0)) then
            Exit;
          if (SensAna <> SensEcr) then
            SigneMoins(Debit,DebitDev);

          Mnt := Debit;
          MntDev := DebitDev;
          end
        else
          begin
          SensAna := 2;
          Credit := Cre-Deb;
          CreditDev := CreDev-DebDev;

          if ((Arrondi(Credit,4) = 0) and (Arrondi(CreditDev,4) = 0)) then
            Exit;
          if (SensAna <> SensEcr) then
            SigneMoins(Credit,CreditDev);

          Mnt := Credit;
          MntDev := CreditDev;
          end;

        if ((totMnt = 0) and (totMntDev = 0)) then
          begin
          totMnt := Mnt;
          totMntDev := MntDev;
          end;

        if (OnDev) then
          begin
          if (Arrondi(totMntDev,4) <> 0) then
            Pourcent := Arrondi(MntDev/totMntDev*100,ADecimP);
          end
        else
          begin
          if (Arrondi(totMnt,4) <> 0) then
            Pourcent := Arrondi(Mnt/totMnt*100,ADecimP);
          end;
        end;
    end;

  Coeff := 1;
  LeDebit   := Debit;
  LeCredit  := Credit;
  LeMnt     := Mnt;

  if (not OnDev)
    then TauxDev := 1
    else
      Case (What) Of
        1 : if (DebitDev <> 0) then
              TauxDev := F.CLoQuotite * ( LeDebit / DebitDev ) * Coeff ;
        2 : if (CreditDev <> 0) then
              TauxDev := F.CLoQuotite * ( LeCredit / CreditDev ) * Coeff;
        3 : if (MntDev <> 0) then
              TauxDev := F.CLoQuotite * ( LeMnt / MntDev ) * Coeff;
        end;

  if (not F.OkGenere) then
    Exit ;

  Inc(NumVentil) ;

  Qaddana.Insert;
  InitNew(Qaddana) ;
  Qaddana.FindField('Y_EXERCICE').AsString        := F.ExoGenere;
  Qaddana.FindField('Y_JOURNAL').AsString         := F.JalGenere;
  Qaddana.FindField('Y_DATECOMPTABLE').AsDateTime := F.DateGenere;
  Qaddana.FindField('Y_NUMEROPIECE').AsInteger    := NumPiece;
  Qaddana.FindField('Y_NUMLIGNE').AsInteger       := NumLig;
  Qaddana.FindField('Y_GENERAL').AsString         := Gen;
  Qaddana.FindField('Y_SECTION').AsString         := Sect;
  Qaddana.FindField('Y_DEBIT').AsFloat            := Arrondi(Debit,FOkDecV);
  Qaddana.FindField('Y_CREDIT').AsFloat           := Arrondi(Credit,FOkDecV);
  Qaddana.FindField('Y_REFINTERNE').AsString      := F.RefGenere;
  Qaddana.FindField('Y_LIBELLE').AsString         := F.LibGenere;
  Qaddana.FindField('Y_NATUREPIECE').AsString     := 'OD';
  Qaddana.FindField('Y_QUALifPIECE').AsString     := F.QualifGenere;
  Qaddana.FindField('Y_VALIDE').AsString          := 'X';
  Qaddana.FindField('Y_DATECREATION').AsDateTime  := date;  // CA -22/04/2003 - FQ 12215
  Qaddana.FindField('Y_DATEMODIF').AsDateTime     := date;  // CA -22/04/2003 - FQ 12215
  Qaddana.FindField('Y_SOCIETE').AsString         := FCodeSociete;
  Qaddana.FindField('Y_ETABLISSEMENT').AsString   := F.CloEtab;
  Qaddana.FindField('Y_DEVISE').AsString          := F.CloDev;
  Qaddana.FindField('Y_DEBITDEV').AsFloat         := Arrondi(DebitDev,F.CloDec);
  Qaddana.FindField('Y_CREDITDEV').AsFloat        := Arrondi(CreditDev,F.CloDec);
  Qaddana.FindField('Y_TAUXDEV').AsFloat          := TauxDev;
  Qaddana.FindField('Y_ECRANOUVEAU').AsString     := F.EcrAnouveauGenere;
  Qaddana.FindField('Y_NUMVENTIL').AsInteger      := NumVentil;
  Qaddana.FindField('Y_AXE').AsString             := Axe;
  Qaddana.FindField('Y_toTALECRITURE').AsFloat    := totMnt;
  Qaddana.FindField('Y_toTALDEVISE').AsFloat      := totMntDev;
  Qaddana.FindField('Y_POURCENTAGE').AsFloat      := PourCent;
  Qaddana.FindField('Y_PERIODE').AsInteger      := GetPeriode(F.DateGenere);
  Qaddana.FindField('Y_SEMAINE').AsInteger      := NumSemaine(F.DateGenere);
  //SG6 07.03.05 mode croisaxe
  Qaddana.FindField('Y_SOUSPLAN1').AsString := splan1;
  Qaddana.FindField('Y_SOUSPLAN2').AsString := splan2;
  Qaddana.FindField('Y_SOUSPLAN3').AsString := splan3;
  Qaddana.FindField('Y_SOUSPLAN4').AsString := splan4;
  Qaddana.FindField('Y_SOUSPLAN5').AsString := splan5;

  Qaddana.Post;

  Ferme(QAddAna) ;

  Result := true;

end;

procedure TTraitementCloture.AlimParamDefaut;
begin
  // Liste des paramètres d'ouverture
  stOuvRef := '';
  stOuvLib := '';
  stOuvJal := GetParamSocSecur('SO_JALOUVRE','');
  stOuvBil := GetParamSocSecur('SO_OUVREBIL','');
  stOuvPer := GetParamSocSecur('SO_OUVREPERTE','');
  stOuvBen := GetParamSocSecur('SO_OUVREBEN','');
  dtOuvDeb := Exo2.Deb ;
  dtOuvFin := Exo2.Fin ;
  // Liste des paramètres de fermeture
  stFermRef := '' ;
  stFermLib := '' ;
  stFermJal := GetParamSocSecur('SO_JALFERME','');
  stFermBil := GetParamSocSecur('SO_FERMEBIL','');
  stFermRes := GetParamSocSecur('SO_RESULTAT','');
  stFermPer := GetParamSocSecur('SO_FERMEPERTE','');
  stFermBen := GetParamSocSecur('SO_FERMEBEN','');
  stFermRes := '' ;
  dtFermDeb := Exo1.Deb ;
  dtFermFin := Exo1.Fin ;
end;

procedure TTraitementCloture.AlimListeCpt(PourBanque: Boolean);
var QCptPV : TQuery;
    St     : String;
begin
  {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
  if not EstSpecif('51216') or not CloDef then Exit;

  if (PourBanque) then
    begin
    St := 'SELECT G_GENERAL FROM GENERAUX WHERE (G_POINTABLE="X" AND G_NATUREGENE="BQE") OR (G_POINTABLE="X" AND G_NATUREGENE<>"BQE" AND G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO" AND G_NATUREGENE<>"EXT" AND G_VENTILABLE="-")';
    QCPTPV := OpenSQL(St,true);
    while (not QCPTPV.EOF) do
      begin
      LCPTBQE.Add(QCPTPV.FindField('G_GENERAL').AsString);
      QCPTPV.Next;
      end;
    end
  else
    begin
    ExisteCptPV := false;
    St:='SELECT G_GENERAL FROM GENERAUX WHERE G_POINTABLE="X" AND G_VENTILABLE="X" AND G_NATUREGENE<>"BQE" AND G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO" AND G_NATUREGENE<>"EXT"';
    QCPTPV:=OpenSQL(St,true);
    while (not QCPTPV.EOF) do
      begin
      ExisteCptPV := true;
      LCPTPV.Add(QCPTPV.FindField('G_GENERAL').AsString);
      QCPTPV.Next;
      end;
    end;

  Ferme(QCPTPV) ;

  EcranCursorSynchr ;

end;

procedure TTraitementCloture.AnnuleRefPointageAnalytique(Sens: integer; StCompteUpdate: String);
var StChampUpdate   : String;
    WhereSup        : String;
    lStReq          : String ;
begin
  {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
  if not EstSpecif('51216') or not CloDef then Exit;

  if (StCompteUpdate = '') then Exit ;
  Case Sens Of
    0 : begin
        StChampUpdate:='E_REFLETTRAGE=E_REFPOINTAGE, E_REFPOINTAGE="" ' ;
        WhereSup:=' E_REFPOINTAGE<>""'
        end;
    1 : begin
        StChampUpdate:='E_REFPOINTAGE=E_REFLETTRAGE, E_REFLETTRAGE="" ' ;
        WhereSup:=' E_REFLETTRAGE<>""'
        end;
    end;

  lStReq := 'UPDATE ECRITURE SET '+StChampUpdate+' WHERE '+StCompteUpdate+' AND E_EXERCICE="'+Exo1.Code+'" AND E_DATECOMPTABLE>="'+UsDatetime(Exo1.Deb)+'" AND E_DATECOMPTABLE<="'+USDatetime(Exo1.Fin)+'"' ;
  // gestion mode IFRS
  if EstComptaIFRS
    then lStReq := lStReq + ' AND (E_QUALIFPIECE="I" OR E_QUALIFPIECE="N") AND ' + WhereSup
    else lStReq := lStReq + ' AND E_QUALIFPIECE="N" AND ' + WhereSup ;
  ExecuteSQL( lStReq );

  EcranCursorSynchr ;
end;

function TTraitementCloture.BalanceOk: Integer ;
Var QBal : TQuery ;
begin
  // -------------------------------------
  // ----- VERIFICATION DE LA BALANCE ----
  // -------------------------------------
  Result := CLO_PASERREUR ;

  // Requête cumul des écritures sur l'exercice
  QBal := OpenSQL('SELECT SUM(E_DEBIT),SUM(E_CREDIT) FROM ECRITURE WHERE E_EXERCICE="'+Exo1.Code+'" and E_QUALifPIECE="N"',true);

  // Est-ce que l'exercice est équilibré en mode normal ?
  if (not QBal.Eof) then
    if (Arrondi( QBal.Fields[0].AsFloat - QBal.Fields[1].AsFloat, FOkDecV ) <> 0)
      then Result := CLO_ERRBALEXO1 ;
  Ferme(QBal) ;

  // Est-ce que l'exercice est équilibré en mode IFRS ?
  if EstComptaIFRS and (CloDef) then
    begin
    QBal := OpenSQL('SELECT SUM(E_DEBIT),SUM(E_CREDIT) FROM ECRITURE WHERE E_EXERCICE="'+Exo1.Code+'" and E_QUALifPIECE="I"',true);
    if (not QBal.Eof) then
      if (Arrondi( QBal.Fields[0].AsFloat - QBal.Fields[1].AsFloat, FOkDecV ) <> 0)
        then Result := CLO_ERRBALEXO1IFRS ;
    Ferme(QBal) ;
    end ;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Function TTraitementCloture.JalExtraComptableOk : Integer;
var lQuery : TQuery;
    lStMessage : string;
begin
  Result := CLO_PASERREUR ;

  if (not Clodef) then Exit; // Pas de contrôle en provisoire

  try
    lQuery := OpenSql('SELECT DISTINCT J_JOURNAL, J_LIBELLE FROM JOURNAL, ECRITURE WHERE ' +
                      'J_JOURNAL = E_JOURNAL AND J_NATUREJAL = "EXT" AND ' +
                      'E_CREERPAR <> "DET" AND E_EXERCICE = "' + Exo1.Code + '" ORDER BY J_JOURNAL', True);

    if not lQuery.Eof then
    begin
      lStMessage := 'Les journaux extra-comptables suivants sont mouvementés : #13#10';
      Result := CLO_WNJALEXTRACOMPTABLE;
      while not lQuery.Eof do
      begin
        lStMessage := lStMessage + ' ' +
                      lQuery.FindField('J_JOURNAL').AsString + ' : ' +
                      lQuery.FindField('J_LIBELLE').AsString + #13#10;
        lQuery.Next;
      end;

      FStLastErrorMsg := lStMessage;
    end;

  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
Function TTraitementCloture.GeneExtraComptableOk : Integer;
begin
  Result := CLO_PASERREUR;
  if (not Clodef) then Exit; // Pas de contrôle en provisoire

  if ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE = "EXT" AND ' +
               '(G_TOTDEBE <> G_TOTCREE)') then
  begin
    Result := CLO_WNGENEEXTRACOMPTABLE;
    FStLastErrorMsg := 'Il existe des comptes généraux extra-comptables non soldés.';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/12/2006
Modifié le ... :   /  /    
Description .. : GCO - 04/12/2006 - FQ 18842
Mots clefs ... : 
*****************************************************************}
function TTraitementCloture.CompteBilanOk : Integer;
var i : integer;
    lStCptMin : string;
    lStCptMax : string;
    lStSql    : string;
begin
  Result := CLO_PASERREUR;
  for i := 1 to 5 do
  begin
    lStCptMin := GetParamSocSecur('SO_BILDEB' + IntToStr(i), '');
    lStCptMax := GetParamSocSecur('SO_BILFIN' + IntToStr(i), '');
    if (lStCptMin <> '') or  (lStCptMax <> '') then
    begin
      lStSql := 'SELECT G_GENERAL FROM GENERAUX WHERE ' +
                IIF( lStCptMin <> '', 'G_GENERAL >= "' + lStCptMin + '" AND ', '') +
                IIF( lStCptMax <> '', 'G_GENERAL <= "' + lStCptMax + '" AND ', '') +
                '(G_NATUREGENE = "CHA" OR G_NATUREGENE = "PRO") AND ' +
                '(G_TOTDEBE - G_TOTCREE) <> 0';

      if ExisteSQl(lStSql) then
      begin
        Result := CLO_WNCPTBILAN;
        FStLastErrorMsg := 'Il existe des comptes de bilans de nature charge ou produit non soldés.';
        Break;
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/12/2006
Modifié le ... :   /  /
Description .. : GCO - 04/12/2006 - FQ 18174
Mots clefs ... :
*****************************************************************}
function TTraitementCloture.CompteChargeOk: Integer;
var i : integer;
    lStCptMin : string;
    lStCptMax : string;
begin
  Result := CLO_PASERREUR ;
  for i := 1 to 5 do
  begin
    lStCptMin := GetParamSocSecur('SO_CHADEB' + IntToStr(i), '');
    lStCptMax := GetParamSocSecur('SO_CHAFIN' + IntToStr(i), '');
    if (lStCptMin <> '') or  (lStCptMax <> '') then
    begin
      if ExisteSQl('SELECT G_GENERAL FROM GENERAUX WHERE ' +
         IIF( lStCptMin <> '', 'G_GENERAL >= "' + lStCptMin + '" AND ', '') +
         IIF( lStCptMax <> '', 'G_GENERAL <= "' + lStCptMax + '" AND ', '') +
         'G_NATUREGENE <> "CHA" AND (G_TOTDEBE - G_TOTCREE) <> 0') then
      begin
        Result := CLO_WNGNATCPT6;
        FStLastErrorMsg := 'Il existe des comptes de charges non soldés qui ne sont pas de bonne nature.';
        Break;
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/12/2006
Modifié le ... :   /  /    
Description .. : GCO - 04/12/2006 - FQ 18174
Mots clefs ... : 
*****************************************************************}
function TTraitementCloture.CompteProduitOk: Integer;
var i : integer;
    lStCptMin : string;
    lStCptMax : string;
begin
  Result := CLO_PASERREUR ;
  for i := 1 to 5 do
  begin
    lStCptMin := GetParamSocSecur('SO_PRODEB' + IntToStr(i), '');
    lStCptMax := GetParamSocSecur('SO_PROFIN' + IntToStr(i), '');
    if (lStCptMin <> '') or  (lStCptMax <> '') then
    begin
      if ExisteSQl('SELECT G_GENERAL FROM GENERAUX WHERE ' +
         IIF( lStCptMin <> '', 'G_GENERAL >= "' + lStCptMin + '" AND ', '') +
         IIF( lStCptMax <> '', 'G_GENERAL <= "' + lStCptMax + '" AND ', '') +
         'G_NATUREGENE <> "PRO" AND (G_TOTDEBE - G_TOTCREE) <> 0') then
      begin
        Result := CLO_WNGNATCPT7;
        FStLastErrorMsg := 'Il existe des comptes de produits non soldés qui ne sont pas de bonne nature.';
        Break;
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/01/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TTraitementCloture.CalculANCLO(var F, F2: TFiltreClo; Recopie: Boolean; QEcr: TQuery; PourPieceCloture, TraitePointage: Boolean);
var NumPiece, NumLig        : LongInt;  // Numéro pièce
    NumPiece2, NumLig2      : LongInt;  // Numéro recopie
    Deb, Cre                : double;   // Montant pièce
    DebDev, CreDev          : double;   // Montant Dev pièce
    totDeb, totCre          : double;   // Total Montant pièce
    totDebDev, totCreDev    : double;   // Total Montant Dev pièce
    totDeb2, totCre2        : double;   // Total Montant recopie
    totDebDev2, totCreDev2  : double;   // Total Montant Dev recopie
    OkInsert, OkInsert2     : Boolean;  // indicateur pièce et recopie
    OldColl,OldColl2        : String;   // indicateur coll pièce et recopie
    Gen, Aux                : String;
    Rien,RefP               : String;
begin

  // Init F pièce
  F.SAJAL := TSaJal.Create(F.JalGenere,false);
  ChangeF(F);
  OldColl := '';

  // Init F recopie
  OldColl2 := '';
  if (Recopie) then
    begin
    ChangeF(F2);
    F2.SAJAL := TSaJal.Create(F2.JalGenere,false);
    end;

  // MAJ num pièce
  if (F.ModeGenerePiece = UneSeule) then
    SetIncNum(F.TypGenere,F.SAJAL.COMPTEURNORMAL,NumPiece,F.DateGenere);
  // MAJ num recopie
  if (Recopie) and (F2.ModeGenerePiece = UneSeule) then
      SetIncNum(F2.TypGenere,F2.SAJAL.COMPTEURNorMAL,NumPiece2,F.DateGenere);

  // Init var pièce
  NumLig      := 0;
  totDeb      := 0;
  totCre      := 0;
  totDebDev   := 0;
  totCreDev   := 0;
  OkInsert    := false;
  // init var recopie
  NumLig2     := 0;
  totDeb2     := 0;
  totCre2     := 0;
  totDebDev2  := 0;
  totCreDev2  := 0;
  OkInsert2   := false;

  if (QEcr <> nil) then
    begin
    while (not QEcr.Eof) do
      begin

      MoveCur(false);

      RefP    := '';
      Deb     := QEcr.FindField('D').AsFloat;
      Cre     := QEcr.FindField('C').AsFloat;
      DebDev  := QEcr.FindField('DD').AsFloat;
      CreDev  := QEcr.FindField('DC').AsFloat;
      Gen     := QEcr.FindField('G').AsString;
      Aux     := Trim(QEcr.FindField('A').AsString);

      {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
      if EstSpecif('51216') then begin
        if (TraitePointage and CloDef) then
          RefP := QEcr.FindField('REFP').AsString;
      end;

      if (PourPieceCloture)
        then TraiteInsertAnClo(F,Gen,Aux,Cre,Deb,CreDev,DebDev,NumPiece,OkInsert,NumLig,totDeb,totCre,totDebDev,totCreDev,PourPieceCloture,OldColl,RefP)
        else TraiteInsertAnClo(F,Gen,Aux,Deb,Cre,DebDev,CreDev,NumPiece,OkInsert,NumLig,totDeb,totCre,totDebDev,totCreDev,PourPieceCloture,OldColl,RefP);

      if Recopie then
        begin
        Deb     := QEcr.FindField('C').AsFloat;
        Cre     := QEcr.FindField('D').AsFloat;
        DebDev  := QEcr.FindField('DC').AsFloat;
        CreDev  := QEcr.FindField('DD').AsFloat;
        Gen     := QEcr.FindField('G').AsString;
        Aux     := Trim(QEcr.FindField('A').AsString);

        TraiteInsertAnClo(F2,Gen,Aux,Deb,Cre,DebDev,CreDev,NumPiece2,OkInsert2,NumLig2,totDeb2,totCre2,totDebDev2,totCreDev2,false,OldColl2,RefP);
        end;

      QEcr.Next;
      end;
    end
  else
    begin
    Deb         := F.tot.D;
    Cre         := F.tot.C;
    DebDev      := F.totDev.D;
    CreDev      := F.totDev.C;
    F.tot.D     := 0;
    F.tot.C     := 0;
    F.totDev.D  := 0;
    F.totDev.C  := 0;
    Gen         := F.CptInit;
    Aux         := '';

    TraiteInsertAnClo(F,Gen,Aux,Deb,Cre,DebDev,CreDev,NumPiece,OkInsert,NumLig,totDeb,totCre,totDebDev,totCreDev,PourPieceCloture,Rien,'');
    end;

  if (OkInsert) then
    Case F.Contrepartie Of
      EnPiedDC : begin
                 AddMvt( 1, F, F.CptContrepGenere, '',
                         totDeb, totCre, totDebDev, totCreDev,
                         F.CloDev, 'RI', '-', '-', '', NumLig, NumPiece, 0, '', '', false, [] );
                 AddMvt( 2, F, F.CptContrepGenere, '',
                         totDeb, totCre, totDebDev, totCreDev,
                         F.CloDev, 'RI', '-', '-', '', NumLig, NumPiece, 0, '', '', false, [] );
                if (Recopie and OkInsert2) then
                  begin
                  AddMvt( 1, F2, F2.CptContrepGenere, '',
                          totDeb2, totCre2, totDebDev2, totCreDev2,
                          F2.CloDev, 'RI', '-', '-', '', NumLig2, NumPiece2, 0, '', '', false, [] );
                  AddMvt( 2, F2, F2.CptContrepGenere, '',
                          totDeb2, totCre2, totDebDev2, totCreDev2,
                          F2.CloDev, 'RI', '-', '-', '', NumLig2, NumPiece2, 0, '', '', false, [] );
                  end;
                end;
      EnPiedSolde : begin
                    AddMvt( 3, F, F.CptContrepGenere, '',
                            totDeb, totCre, totDebDev, totCreDev,
                            F.CloDev, 'RI', '-', '-', '', NumLig, NumPiece, 0, '', '', false, [] );
                    if (Recopie and OkInsert2) then
                      AddMvt( 3, F2, F2.CptContrepGenere, '',
                              totDeb2, totCre2, totDebDev2, totCreDev2,
                              F2.CloDev, 'RI', '-', '-', '', NumLig2, NumPiece2, 0, '', '', false, [] );
        end;
    end;

  // MAJ des numéro de pièce
  if ((F.ModeGenerePiece = UneSeule) and (not OkInsert)) then
    SetDecNum(F.TypGenere,F.SAJAL.COMPTEURNORMAL,F.DateGenere);
  if (Recopie) and ((F2.ModeGenerePiece = UneSeule) and (not OkInsert2)) then
      SetDecNum(F2.TypGenere,F2.SAJAL.COMPTEURNORMAL,F.DateGenere);

  // Libération données du journal
  FreeAndNil(F.SAJAL) ;
  if (Recopie) then
    FreeAndNil(F2.SAJAL) ;
end;

procedure TTraitementCloture.CalculANCLOANA(F: TFiltreClo; Gen, Aux: String; Sens: Byte; NumPiece, NumLig: Integer; totMnt, totMntDev: double);
var NumVentil          : LongInt;
    Deb , Cre          : double ;
    DebDev, CreDev     : double ;
    Sect, Axe          : String;
    SDeb, SCre         : double;
    SDebDev, SCreDev   : double;
    st                 : string;
    Qana               : TQuery;
    splan1, splan2, splan3, splan4, splan5 : string;
begin
  NumVentil := 0;

  st := 'SELECT Y_GENERAL AS G,Y_SECTION AS S,Y_AXE AS A,SUM(Y_DEBIT) AS D, SUM(Y_CREDIT) AS C, SUM(Y_DEBITDEV) AS DD, SUM(Y_CREDITDEV) AS DC, Y_SOUSPLAN1 as YS1, ';
  st := st + 'Y_SOUSPLAN2 as YS2, Y_SOUSPLAN3 as YS3, Y_SOUSPLAN4 as YS4, Y_SOUSPLAN5 as YS5 FROM ANALYTIQ WHERE Y_GENERAL="'+gen+'" AND';
// GP le 16/07/2008 : je me marre...  if (V_PGI.Driver in [dboracle7,dboracle8])
  If isOracle
    then st := st + ' (Y_AUXILIAIRE LIKE "'+Aux+'%" OR Y_AUXILIAIRE IS NULL) AND '
    else if (V_PGI.Driver = dbMSACCESS)
      then st := st + ' Y_AUXILIAIRE LIKE "'+Aux+'*" AND '
      else st := st + ' Y_AUXILIAIRE LIKE "'+Aux+'%" AND ';
  st := st + 'Y_EXERCICE="'+F.CloExo.Code+'" AND Y_ETABLISSEMENT="'+F.CloEtab+'" AND Y_DEVISE="'+F.CloDev+'"' ;
  if modeIFRS
    then st := st + ' AND Y_QUALIFPIECE="I" AND Y_TYPEANALYTIQUE<>"X" GROUP BY Y_GENERAL,Y_SECTION,Y_AXE, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
    else st := st + ' AND Y_QUALIFPIECE="N" AND Y_TYPEANALYTIQUE<>"X" GROUP BY Y_GENERAL,Y_SECTION,Y_AXE, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5';

  QAna := OpenSQL(st,true);

  while not QAna.Eof do
    begin

    MoveCur(false);

    Deb       := QAna.FindField('D').AsFloat;
    Cre       := QAna.FindField('C').AsFloat;
    DebDev    := QAna.FindField('DD').AsFloat;
    CreDev    := QAna.FindField('DC').AsFloat;
    SDeb      := Deb;
    SCre      := Cre;
    SDebDev   := DebDev;
    SCreDev   := CreDev;
    Deb       := SDeb;
    DebDev    := SDebDev;

    Sect      := Trim(QAna.FindField('S').AsString);
    Axe       := Trim(QAna.FindField('A').AsString);
    //SG6 07.03.03 Gestion croisaxe
    splan1    := Trim(QAna.FindField('YS1').AsString);
    splan2    := Trim(QAna.FindField('YS2').AsString);
    splan3    := Trim(QAna.FindField('YS3').AsString);
    splan4    := Trim(QAna.FindField('YS4').AsString);
    splan5    := Trim(QAna.FindField('YS5').AsString);

    if ((Arrondi(Deb,4) <> 0) and (Arrondi(DebDev,4) <> 0) and ((Sens = 1) or (Sens = 3))) then
      begin
      Cre     := 0;
      CreDev  := 0;
      AddMvtAna(3,F,Gen,Sect,Deb,Cre,DebDev,CreDev,NumLig,NumVentil,NumPiece,Sens,Axe,totMnt,totMntDev,splan1, splan2, splan3, splan4, splan5);
      end;

    Cre     := SCre;
    CreDev  := SCreDev;

    if ((Arrondi(Cre,4) <> 0) and (Arrondi(CreDev,4) <> 0) and ((Sens = 2) or (Sens = 3))) then
      begin
      Deb     := 0;
      DebDev  := 0;
      AddMvtAna(3,F,Gen,Sect,Deb,Cre,DebDev,CreDev,NumLig,NumVentil,NumPiece,Sens,Axe,totMnt,totMntDev, splan1, splan2, splan3, splan4, splan5);
      end;

    QAna.Next;

    end;

    Ferme(QAna) ;

end;

function TTraitementCloture.ChangeEcartDeChange(Sens: integer): Boolean;
var St            : String;
    StChampUpdate : String;
begin
  Result := false;
  Case Sens Of
    0 : StChampUpdate := 'E_QUALIFORIGINE=E_DEVISE, E_DEVISE="'+FDevisePivot+'" ';
    1 : StChampUpdate := 'E_DEVISE=E_QUALIFORIGINE, E_QUALIFORIGINE="" ';
    end;

  St:='UPDATE ECRITURE SET '+StChampUpdate+' WHERE E_EXERCICE="'+Exo1.Code+'" AND E_DATECOMPTABLE>="'+UsDatetime(Exo1.Deb)+'" AND E_DATECOMPTABLE<="'+USDatetime(Exo1.Fin)+'"' ;
  // Gestion cloture IFRS
  if modeIFRS
    then St := St + ' AND E_QUALIFPIECE="I" AND E_NATUREPIECE="ECC"'
    else St := St + ' AND E_QUALIFPIECE="N" AND E_NATUREPIECE="ECC"';

  if (ExecuteSQL(St) > 0) then
    Result := true;

  EcranCursorSynchr ;
end;

procedure TTraitementCloture.ChangeRefPointage(Sens: integer; StCompteUpdate: String);
var StChampUpdate : String ;
    lStReq        : String ;
begin
  if (StCompteUpdate = '') then Exit;
  Case Sens Of
    0 : StChampUpdate := 'E_REFLETTRAGE=E_REFPOINTAGE, E_REFPOINTAGE="POINTAGE" ';
    1 : StChampUpdate := 'E_REFPOINTAGE=E_REFLETTRAGE, E_REFLETTRAGE="" ';
    end;

  lStReq := 'UPDATE ECRITURE SET '+StChampUpdate+' WHERE '+StCompteUpdate+' AND E_EXERCICE="'+Exo1.Code+'" AND E_DATECOMPTABLE>="'+UsDatetime(Exo1.Deb)+'" AND E_DATECOMPTABLE<="'+USDatetime(Exo1.Fin)+'"' ;
  // gestion mode IFRS
  if EstComptaIFRS
    then lStReq := lStReq + ' AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="I") AND E_REFPOINTAGE<>""'
    else lStReq := lStReq + ' AND E_QUALIFPIECE="N" AND E_REFPOINTAGE<>""' ;
  ExecuteSQL( lStReq );
  EcranCursorSynchr ;
end;

procedure TTraitementCloture.ChercheCpt(Gen, Aux: String; var EtatLettrage, Ana, Eche: String; var NumEche: Integer; var Coll: Boolean; var PointVent: TPointVent; var Setfb: TsetFichierBase);
var
  Ventilable,lettrable,PointableBqe,Pointable : Boolean;
  i    : Byte;
  fb   : tfichierbase;
  QGen : TQuery;
  QAux : TQuery;

begin
  // Init des résultat
  EtatLettrage := '';
  Ana := '-';
  Eche := '-';
  NumEche := 0;
  Coll := false;
  PointVent := PVRien;

  // Cas des généraux
  if (Aux = '') then
    begin
    QGen := OpenSQL('SELECT G_GENERAL, G_NATUREGENE, G_LETTRABLE, G_POINTABLE, G_VENTILABLE, G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, G_COLLECTIF FROM GENERAUX WHERE G_GENERAL="'+Gen+'"',true);

    if (not QGen.Eof) then
      begin
      Ventilable := (QGen.FindField('G_VENTILABLE').AsString = 'X');
      for i:=1 to 5 do
        begin
        fb := TFichierBase(ord(fbaxe1)+i-1);
        // Analyse des champs G_VENTILABLE1 à G_VENTILABLE5
        if (QGen.Fields[i+4].AsString = 'X') then
          Setfb := Setfb+[fb];
        end;

      Lettrable     := (QGen.FindField('G_LETTRABLE').AsString = 'X');
      PointableBqe  := ((QGen.FindField('G_POINTABLE').AsString = 'X') and (QGen.FindField('G_NATUREGENE').AsString = 'BQE'));
      Pointable     := (QGen.FindField('G_POINTABLE').AsString = 'X');
      Coll          := (QGen.FindField('G_COLLECTIF').AsString = 'X');

      // Param de l'analytique
      if (Ventilable) then
        Ana := 'X';
      // Param du lettrage et échéances
      if (Lettrable or PointableBqe) then
        begin
        Eche := 'X';
        NumEche := 1;
        end;
      // Mise en place de l'état du lettrage
      if (Lettrable) then
        EtatLettrage := 'AL';
      if (PointableBqe) then
        EtatLettrage := 'RI';
      if EtatLettrage = '' then
        EtatLettrage:='RI';  // CA - 22/04/2003 - FQ 12159
      // Param type de pointage
      if (Pointable) then
        if (Ventilable) then
          if (PointableBQE)
            then PointVent := BQPointableVentilable
            else PointVent := AutrePointableVentilable
        else
          if (PointableBQE)
            then PointVent := BQPointable
            else PointVent := AutrePointable
      end;
    Ferme(QGen) ;
    end
  else // Cas des auxiliaire
    begin
    // L'auxiliaire est-il lettrable ?
    QAux := OpenSQL('SELECT T_AUXILIAIRE, T_LETTRABLE FROM TIERS WHERE T_AUXILIAIRE="'+Aux+'"',true);
    if (not QAux.Eof) then
      begin
      Lettrable := (QAux.FindField('T_LETTRABLE').AsString = 'X');
      // Param lettrage / Echéances
      if (Lettrable) then
        begin
        Eche := 'X' ;
        NumEche := 1;
        EtatLettrage := 'AL';
        end
      else
        EtatLettrage:='RI';  // CA - 22/04/2003 - FQ 12159
      end;
    Ferme(QAux) ;

    // TESTS sur le général
    QGen := OpenSQL('SELECT G_GENERAL, G_NATUREGENE, G_LETTRABLE, G_POINTABLE, G_VENTILABLE, G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, G_COLLECTIF FROM GENERAUX WHERE G_GENERAL="'+Gen+'"',true);
    if (not QGen.Eof) then
      begin
      // Test du champ G_VENTILABLE
      Ventilable := (QGen.FindField('G_VENTILABLE').AsString = 'X');
      for i := 1 to 5 do
        begin
        fb := TFichierBase(ord(fbaxe1)+i-1);
        // Test des champs G_VENTILABLE1 à G_VENTILABLE5
        if (QGen.Fields[i+4].AsString = 'X') then
          Setfb := Setfb+[fb];
        end;
      if (Ventilable) then
        Ana := 'X';
      Coll := (QGen.FindField('G_COLLECTIF').AsString = 'X');
      end;
    Ferme(QGen) ;
    end;
end;

constructor TTraitementCloture.Create( vEcran: TForm ; vStCodeExo1, vStCodeExo2 : String ; vBoCloDef, vBoAuto : Boolean ; vAnoDyna : boolean = false );
begin
  // sortie Ecran
  Ecran   := vEcran ;
  Auto    := vBoAuto ;
  // Cloture définitive ?
  CloDef  := vBoCloDef ;
  // Exo à cloturer
  if vStCodeExo1 <> '' then
    begin
    Exo1.Code := vStCodeExo1 ;
    RempliExoDate(Exo1) ;
    end
  else
    Exo1 := VH^.EnCours;
  // Exo suivant
  if vStCodeExo2 <> '' then
    begin
    Exo2.Code := vStCodeExo2 ;
    RempliExoDate(Exo2) ;
    end
  else
    Exo2 := VH^.Suivant;

  FStLastErrorMsg := '';
  modeAnoDyna     := vAnoDyna ;

  // Init Autres attributs
  Initialisation ;
end;

procedure TTraitementCloture.DeleteMouvementResultat(var F: TFiltreClo);
var lStSQL  : String ;
begin

  lStSql := 'DELETE FROM ECRITURE WHERE E_JOURNAL="'+F.JalGenere+'" AND E_EXERCICE="'+F.ExoGenere+'" AND E_DATECOMPTABLE="'+UsDatetime(F.DateGenere)+'" AND E_NUMEROPIECE=123456789 AND E_NUMLIGNE=123456' ;
  // Gestion IFRS
  if modeIFRS
    then lStSQL := lStSQL + ' AND E_QUALIFPIECE="I" AND E_ECRANOUVEAU="GGG"'
    else lStSQL := lStSQL + ' AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="GGG"' ;

  ExecuteSQL( lStSql );
end;

destructor TTraitementCloture.Destroy;
begin
  // TStringList
  LLEtabDev.Free;
  if (CloDef) then
    begin
    LCptBQE.Free;
    LCptPV.Free;
    end;

  inherited Destroy ;

end;

procedure TTraitementCloture.DetruitANO;
begin
  if (Exo2.Code = '') then Exit;
  if ModeIFRS then Exit ;               // Modif SBO ne pas exécuter lors du traitement des IFRS
  ExecuteSQL('DELETE FROM ECRITURE WHERE E_EXERCICE="'+Exo2.Code+'" and E_ECRANOUVEAU="OAN"');
  ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_EXERCICE="'+Exo2.Code+'" and Y_ECRANOUVEAU="OAN"');
  ExecuteSQL('DELETE FROM ECRITURE WHERE E_EXERCICE="'+Exo1.Code+'" and E_ECRANOUVEAU="C"');
  ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_EXERCICE="'+Exo1.Code+'" and Y_ECRANOUVEAU="C"');

  // GCO - FQ 13089 - Suppression des ecritures typées "V" en Cloture Définitive
  if CloDef then
    ExecuteSQL('DELETE FROM ECRITURE WHERE E_EXERCICE="'+Exo1.Code+'" AND E_QUALIFPIECE="V"');
  // FIN GCO
end;

procedure TTraitementCloture.EcranCursorSynchr;
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFClos ) then
    TFClos(Ecran).CursorSynchr ;
{$ENDIF}
end;

procedure TTraitementCloture.EcranInitGeneral ;
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFClos ) then
    TFClos(Ecran).InitGeneral ;
{$ENDIF}
end;

procedure TTraitementCloture.EcranInitMove(vInQte: Integer; vStTitre: String);
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  InitMove ( vInQte , vStTitre ) ;
{$ENDIF}
end;

procedure TTraitementCloture.EcranMessage1(i: Integer);
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFClos ) then
    TFClos(Ecran).Mess1(i) ;
{$ENDIF}
end;

procedure TTraitementCloture.EcranMessage2(i: Integer; St1, St2: String);
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFClos ) then
    TFClos(Ecran).Mess2(i,St1,St2) ;
{$ENDIF}
end;

procedure TTraitementCloture.EcranMoveCur;
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  MoveCur( false );
{$ENDIF}
end;

procedure TTraitementCloture.EcranFiniMove;
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  FiniMove ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/07/2003
Modifié le ... :   /  /
Description .. : Appelé par le process server pour executer le traitement de
Suite ........ : la cloture :
Suite ........ :  - TobParam contient tout le paramètrage (cf
Suite ........ : CreerTobParam)
Suite ........ :  - retourne une tob avec indicateurs de réussite des étapes
Mots clefs ... :
*****************************************************************}
function TTraitementCloture.ExecuteCloture ( TobParam : TOB ) : Tob;
var OkTraite  : Boolean ;
    ClotureOK : Boolean ;
    errID     : Integer ;
    TobResult : TOB ;
    onSort    : Boolean ;
begin

  // Préparation de la tob retournée
  TOBResult := TOB.Create('CLOTUREPROCESS', nil, -1) ;
  TOBResult.AddChampSupValeur('RESULT',    CLO_PASERREUR ) ;
  TOBResult.AddChampSupValeur('ONSORT',    TobParam.GetValue('ONSORT') ) ;
  TOBResult.AddChampSupValeur('OKTRAITE',  False ) ;
  TOBResult.AddChampSupValeur('CLOTUREOK', False ) ;
  result := TobResult ;

  if not modeServeur then
    Exit ;

  // ----------------------------------
  // ----- AFFECTATION DES PARAMS -----
  // ----------------------------------
  AffectParamDepuisTOB ( TobParam ) ;

  // --------------------------------
  // ----- TRAITEMENT DE CLOTURE ----
  // --------------------------------
  errID := Cloture ( ClotureOk, OkTraite ) ;
  if (errID <> CLO_PASERREUR) then
    if (not Auto) then
      begin
      TOBResult.PutValue('OKTRAITE', OkTraite ) ;
      TOBResult.putValue('RESULT', errID ) ;
      TOBResult.PutValue('CLOTUREOK', ClotureOk ) ;
      Exit;
      end ;

  // --------------------------------
  // ----- VERIFICATION GENERALE ----
  // --------------------------------
  if (OkTraite and (not Auto)) then
    VerifCloture ( OnSort )
  else
    OnSort:=false;

  // -----------------------
  // ----- IT'S THE END ----
  // -----------------------
  // La cloture est ok
  TOBResult.PutValue('ONSORT', onSort ) ;
  TOBResult.PutValue('OKTRAITE', OkTraite ) ;
  TOBResult.putValue('RESULT', CLO_CLOTUREOK ) ;
  TOBResult.PutValue('CLOTUREOK', ClotureOk ) ;

end;

function TTraitementCloture.GestionEcran: Boolean;
begin
  Result := (not modeServeur) and (not Auto) and (Ecran <> nil) ;
end;

function TTraitementCloture.GetQteEtabDev: Integer;
begin
  Result := 0 ;
  if LLEtabDev = nil then
    Exit ;
  Result := LLEtabDev.Count ;
end;

procedure TTraitementCloture.InitGeneral(var F: TFiltreClo; I: Integer);
Var Q : TQuery ;
begin

  EcranInitGeneral ;

  F.tot.D       := 0;
  F.tot.C       := 0;
  F.totDev.D    := 0;
  F.totDev.C    := 0;
  F.OkGenere    := true;
  F.CptInit     := '';
  F.CloDev      := Trim(Copy(LLEtabDev.Strings[i],1,3)) ;
  F.CloEtab     :=Trim(Copy(LLEtabDev.Strings[i],4,3));
  F.CloQuotite  := 1;
  // CA - 22/04/2003
  //F.ModePaieGenere := 'DIV';
  F.ModePaieGenere  := ModePaiementDefaut;
  F.ModeGenerePiece := UneParCpt;
  F.SAJAL           := nil ;
  F.CloLibDevise    := '';
  F.CloLibEtab      := '';

  Q := OpenSQL('SELECT D_DEVISE, D_DECIMALE, D_QUOTITE, D_LIBELLE FROM DEVISE WHERE D_DEVISE="'+F.CloDev+'"',false);
  if (not Q.Eof) then
    begin
    F.CloDec        := Q.FindField('D_DECIMALE').AsInteger;
    F.CloQuotite    := Q.FindField('D_QUOTITE').AsFloat;
    F.CloLibDevise  := Q.FindField('D_LIBELLE').AsString;
    end;
  Ferme(Q) ;

  Q := OpenSQL('SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT="'+F.CloEtab+'"',false);
  if (not Q.Eof) then
    F.CloLibEtab := Q.FindField('ET_LIBELLE').AsString;
  Ferme(Q) ;

end;

procedure TTraitementCloture.Initialisation;
begin

  // Pb initialisation des variables en EAGL-Serveur FQ 15943 SBO 01/08/2005
  FOkDecV        := GetParamSocSecur('SO_DECVALEUR',0) ;
  FDevisePivot   := GetParamSocSecur('SO_DEVISEPRINC','') ;
  FCodeSociete   := GetParamSocSecur('SO_SOCIETE','') ;

  modeServeur := False ;
  modeIFRS    := False ;

  // Init Liste des couples Etablissements / devise
  LLEtabDev := TStringList.Create;
  LLEtabDev.Sorted := false;
  LLEtabDev.Duplicates := DupIgnore;
  LLEtabDev.Clear;

  // Paramètre par défaut
  ParaClo.CloContrep      := EnPiedDC;
  ParaClo.ANOContrep      := EnPiedDC;
  ParaClo.CloPiece        := UneSeule;
  ParaClo.ANOPiece        := UneSeule;
  ParaClo.ANOCompteBanque := BQUneParRef;
  ParaClo.ANOComptePV     := SurAnalytique;

  // Pour la clôture définitive, on construit une liste des comptes de banque
  //  et des comptes ventilables et pointables
  ExisteCptPV := False ;
  if (CloDef) then
    begin
    LCptBQE := TStringList.Create;
    LCptPV := TStringList.Create;
    AlimListeCpt(true);
    AlimListeCpt(false);
    end;

  AlimParamDefaut ;

end;

procedure TTraitementCloture.InitQCHAPRO(Nat: String3; var F: TFiltreClo);
begin
  F.tot.D             := 0;
  F.tot.C             := 0;
  F.totDev.D          := 0;
  F.totDev.C          := 0;
  F.EnSolde           := true ;
  F.CptContrepGenere  := StFermRes ;
  F.Contrepartie      := EnPiedSolde ;
  F.CreerAna          := false ;
  F.ModeGenerePiece   := UneSeule ;
end;

procedure TTraitementCloture.InitQResultatAN(var F: TFiltreClo; totResultat, totResultatDev: ttDCClo; Cpt: String);
begin
  F.tot.D       := totResultat.D;
  F.tot.C       := totResultat.C;
  F.totDev.D    := totResultatDev.D;
  F.totDev.C    := totResultatDev.C;
  F.EnSolde     := true;
  F.CptContrepGenere := Cpt;
  F.Contrepartie := EnPiedSolde;
  F.CreerAna    := false;
  F.CptInit     := StFermRes ;
end;

procedure TTraitementCloture.InitQResultatClo(var F: TFiltreClo; totResultat, totResultatDev: ttDCClo; Cpt: String);
begin
  F.tot.D     := totResultat.C;
  F.tot.C     := totResultat.D;
  F.totDev.D  := totResultatDev.C;
  F.totDev.C  := totResultatDev.D;
  F.EnSolde   := true;
  F.CptContrepGenere := Cpt;
  F.Contrepartie := EnPiedSolde;
  F.CreerAna  := false;
  F.CptInit   := StFermRes;
  F.ModeGenerePiece := UneSeule;
end;

procedure TTraitementCloture.InitSupp(var F: TFiltreClo; What: Byte);
begin
  Case What Of
    // Pour Cloture
    0 : begin
        F.CloExo        := Exo1;
        F.ExoGenere     := Exo1.Code;
        F.JalGenere     := stFermJal;
        F.QualifGenere  := 'C';
        F.TypGenere     := EcrClo;
        F.DateGenere    := Exo1.Fin;
        F.RefGenere     := stFermRef;
        F.LibGenere     := stFermLib;
        F.EnSolde       := true;
        F.EcrANouveauGenere := 'C';
        F.CptInit       := '';
        F.SAJAL         := nil;
        end;
    // Pour AN
    1 : begin
        F.CloExo        := Exo1;
        F.ExoGenere     := Exo2.Code;
        F.JalGenere     := stOuvJal;
        // Gestion IFRS
        if ModeIFRS
          then F.QualifGenere  := 'I'
          else F.QualifGenere  := 'N';
        F.TypGenere     := EcrGen;
        F.DateGenere    := Exo2.Deb;
        F.RefGenere     := stOuvRef;
        F.LibGenere     := stOuvLib;
        F.EnSolde       := true;
        F.EcrANouveauGenere := 'OAN';
        F.CptInit       := '';
        F.SAJAL         := nil;
        end;
    // Pour écriture provisoire de résultat sur N
    2 : begin
        F.CloExo        := Exo1;
        F.ExoGenere     := Exo1.Code;
        F.JalGenere     := stOuvJal;
        // Gestion IFRS
        if ModeIFRS
          then F.QualifGenere  := 'I'
          else F.QualifGenere  := 'N';
        F.TypGenere     := EcrGen;
        F.DateGenere    := Exo1.Fin;
        F.RefGenere     := '';
        F.LibGenere     := '';
        F.EnSolde       := true;
        F.EcrANouveauGenere := 'N';
        F.CptInit       := '';
        F.SAJAL         := nil ;
        end;
    end;
end;

procedure TTraitementCloture.InsertMvtResultat(var F: TFiltreClo; Cpt: String);
var Qi : TQuery ;
begin
  Qi := OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL="###"',false);
  Qi.Insert;
  InitNew(Qi);
  Qi.FindField('E_EXERCICE').AsString     := F.ExoGenere;
  Qi.FindField('E_JOURNAL').AsString      := F.JalGenere;
  Qi.FindField('E_DATECOMPTABLE').AsDateTime := F.DateGenere;
  Qi.FindField('E_NUMEROPIECE').AsInteger := 123456789;
  Qi.FindField('E_NUMLIGNE').AsInteger    := 123456;
  Qi.FindField('E_GENERAL').AsString      := Cpt;
  Qi.FindField('E_AUXILIAIRE').AsString   := '';
  Qi.FindField('E_DEBIT').AsFloat         := F.tot.D;
  Qi.FindField('E_CREDIT').AsFloat        := F.tot.C;
  Qi.FindField('E_ENCAISSEMENT').asString := SENSENC(F.tot.D,F.tot.C) ;
//  Qi.FindField('E_QUALifPIECE').AsString  := 'N';
  Qi.FindField('E_QUALIFPIECE').AsString  := F.QualifGenere ; // Modif IFRS
  Qi.FindField('E_ETABLISSEMENT').AsString := F.CloEtab;
  Qi.FindField('E_DEVISE').AsString       := F.CloDev;
  Qi.FindField('E_DEBITDEV').AsFloat      := F.totDev.D;
  Qi.FindField('E_CREDITDEV').AsFloat     := F.totDev.C;
  Qi.FindField('E_TAUXDEV').AsFloat       := 1;
  Qi.FindField('E_ECRANOUVEAU').AsString  := 'GGG';
  Qi.FindField('E_NUMECHE').AsInteger     := 0;
{$IFNDEF SPEC302}
  Qi.FindField('E_PERIODE').AsInteger     := GetPeriode(F.DateGenere);
  Qi.FindField('E_SEMAINE').AsInteger     := NumSemaine(F.DateGenere);
{$ENDIF}
  Qi.Post;
  Ferme( Qi );
end;

function TTraitementCloture.IsPointableATraiterCommeVentilable( Cpt: String ): Boolean;
var i : Integer ;
begin
  Result := false;
  {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
  if not EstSpecif('51216') or not CloDef then Exit;

  if (ParaClo.AnoComptePV <> SurAnalytique) then Exit ;
  for i := 0 to LCptPV.Count-1 do
    if (Cpt = LCPTPV.Strings[i]) then
      begin
      result := true;
      Exit;
      end;
end;

procedure TTraitementCloture.majExo ;
begin
  if (Auto) then Exit;
  if (CloDef) then
    begin
    ExecuteSQL('UPDATE EXERCICE SET EX_ETATCPTA="CDE" WHERE EX_EXERCICE="'+Exo1.Code+'"') ;
    end
  else
    begin
    SetParamSoc('SO_DATECLOTUREPRO',Now);
    ExecuteSQL('UPDATE EXERCICE SET EX_ETATCPTA="CPR" WHERE EX_EXERCICE="'+VH^.EnCours.Code+'"');
    end;

end;

function TTraitementCloture.ModePaiementDefaut: string;
var TMP : TOB;
begin
  if not ExisteSQL ('SELECT * FROM MODEPAIE WHERE MP_MODEPAIE="DIV"') then
    begin
    TMP := TOB.Create ('MODEPAIE',nil,-1);
    TMP.PutValue('MP_MODEPAIE',     'DIV');
    TMP.PutValue('MP_LIBELLE',      'Divers');
    TMP.PutValue('MP_ABREGE',       'Divers');
    TMP.PutValue('MP_GENERAL',      GetParamSocSecur('SO_DEFCOLFOU',''));
    TMP.PutValue('MP_ENCAISSEMENT', 'MIX');
    TMP.PutValue('MP_CATEGORIE',    'CHQ');
    TMP.InsertDB(nil);
    TMP.Free;
    end;
  Result := 'DIV';
end ;

function TTraitementCloture.NbEtabDev: Boolean;
var st         : String ;
    AvecDevise : Boolean ;
    Q          : TQuery ;
    lStReq     : String ;
begin
  Result := false;
  AvecDevise := false;
  // Récupération des couples devise / établissement
  if ModeIFRS
    then lStReq := 'SELECT E_DEVISE,E_ETABLISSEMENT FROM ECRITURE WHERE E_EXERCICE="'+Exo1.Code+'" AND E_QUALIFPIECE="I" GROUP BY E_DEVISE, E_ETABLISSEMENT'
    else lStReq := 'SELECT E_DEVISE,E_ETABLISSEMENT FROM ECRITURE WHERE E_EXERCICE="'+Exo1.Code+'" GROUP BY E_DEVISE, E_ETABLISSEMENT' ;
  Q := OpenSQL( lStReq ,true);
  // Parcours de la query
  while (not Q.Eof) do
    begin
    if (Q.Fields[0].AsString <> FDevisePivot) then
      AvecDevise := true;
    St := format_String(Q.Fields[0].AsString,3) + format_String(Q.Fields[1].AsString,3);
    // Ajout des couples devise / établissement à la TList
    LLEtabDev.Add(St);
    Q.Next;
    end;

  Ferme(Q) ;

  if (AvecDevise) then
    if (ChangeEcartDechange(0)) then
      Result := true;
end;

function TTraitementCloture.PbSurMontantDevise(What: Integer; Deb, Cre, DebDev, CreDev: double; CloDec: Integer): Boolean;
var
  Mnt,MntDev : double;
  SommeAbs,AbsSomme : double;
begin
  Result := false;

  Case What Of
    // Pièce en devise et pas pour les écritures de cloture
    0 : begin
        Mnt     := Arrondi(Deb-Cre,FOkDecV);
        MntDev  := Arrondi(DebDev-CreDev,CloDec);
        if ((Mnt = 0) and (MntDev = 0)) then Exit;
        SommeAbs := Abs(Mnt+MntDev);
        AbsSomme := Abs(Mnt)+Abs(MntDev);
        Result := ((Mnt = 0) or (MntDev = 0) or (Arrondi(SommeAbs-AbsSomme,4) <> 0));
        end;
    // Pièce En Pivot
    1 : begin
        Mnt     := Arrondi(Deb-Cre,FOkDecV);
        MntDev  := Arrondi(DebDev-CreDev,CloDec);
        if ((Mnt = 0) and (MntDev = 0)) then Exit;
        if ((Mnt = 0) and (MntDev = 0)) then Result := true;
    end;
  end;
end;

function TTraitementCloture.PositionneStUpdate(L: TStringList): String;
var i     : Integer;
    Prem  : Boolean;
begin
  Result := '';
  Prem := true;

  {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
  if not EstSpecif('51216') or not CloDef then Exit;

  for i := 0 to L.Count-1 do
    begin
    if (prem)
      then Result := Result+' E_GENERAL="'    + L.Strings[i] + '" '
      else Result := Result+' or E_GENERAL="' + L.Strings[i] + '" ' ;
    Prem := false;
    end;

  if (Result <> '') then
    Result := '('+Result+')';

end;

procedure TTraitementCloture.QuelParam(var F, F2: tFiltreClo);
begin
  F2.Contrepartie     := ParaClo.CloContrep;
  F.Contrepartie      := ParaClo.ANOContrep;
  F2.ModeGenerePiece  := ParaClo.CloPiece;
  F.ModeGenerePiece   := ParaClo.ANOPiece;
end;

procedure TTraitementCloture.TraiteInsertAnClo(var F: TFiltreClo; Gen, Aux: String; Deb, Cre, DebDev, CreDev: double; var NumPiece: Integer; var OkInsert: Boolean; var NumLig: Integer; var totDeb, totCre, totDebDev, totCreDev: double; PourPieceCloture: Boolean; var SOldColl: String; RefP: String);
var EtatLettrage, ModePaie  : String;
    Eche, Ana               : String;
    NumEche                 : Integer;
    OkEnSolde, Coll         : Boolean;
    forcerendebitCredit     : Boolean;
    CreerAnaSurAttente      : Boolean;
    StotDeb, StotCre        : double;
    StotDebDev, StotCreDev  : double;
    forceRef, forceLib      : String;
    PointVent               : TPointVent;
    Setfb                   : TsetFichierBase;
begin
  forceRef := '';
  forceLib := '';
  forcerendebitCredit := false;
  Setfb := [];
  CreerAnaSurAttente := false;

  if (F.ModeGenerePiece = UneParCpt) then
    begin
    SetIncNum(F.TypGenere,F.SAJAL.COMPTEURNORMAL,NumPiece,F.DateGenere);
    NumLig := 0;
    end;

  ChercheCpt(Gen,Aux,EtatLettrage,Ana,Eche,NumEche,Coll,PointVent,Setfb);

  {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
  if EstSpecif('51216') then begin
    if (CloDef) then
      Case PointVent Of
        //-------------
        BQPointable,BQPointableVentilable,AutrePointable :
          begin
          if (RefP <> '') then
            begin
            forceLib := TraduireMemoire('Total des écritures pointées'); //HMess.Mess[49];
            forceRef := RefP;
            end
          else
            begin
            forceLib := TraduireMemoire('Total des écritures non pointées'); //HMess.Mess[50];
            forceRef := '';
            forcerendebitCredit := true;
            end;
          if (PointVent = BQPointableVentilable) then
            CreerAnaSurAttente := true;
          end;
        //-------------
        AutrePointableVentilable :
          begin
          if (not IsPointableATraiterCommeVentilable(Gen)) then
            begin
            CreerAnaSurAttente := true;
            if (RefP <> '') then
              begin
              forceLib := TraduireMemoire('Total des écritures pointées'); //HMess.Mess[49];
              forceRef := RefP;
              end
            else
              begin
              forceLib := TraduireMemoire('Total des écritures non pointées'); //HMess.Mess[50];
              forceRef := '';
              forcerendebitCredit := true;
              end;
            end;
          end;
        //-------------
        PVRien : ;
      end; // FIN CASE
  end;

  if (not F.CreerAna) then
    Ana := '-';
  if (Eche = 'X')
    then ModePaie := F.ModePaieGenere
    else ModePaie := '';
  if (F.CloDev = FDevisePivot) then
    begin
    DebDev := Deb;
    CreDev := Cre;
    end;

  OkEnSolde := F.EnSolde;

  if (OkEnSolde) then
    if (Ana = 'X') then
      if (not CreerAnaSurAttente) then
        OkEnSolde := false;
  if (OkEnSolde) then
    if (forcerendebitCredit) then
      OkEnSolde := false;
  if (OkEnSolde and (F.CloDev <> FDevisePivot) and (not PourPieceCloture)) then
    begin
    if (PbSurMontantDevise(0,Deb,Cre,DebDev,CreDev,F.CloDec)) then
      OKEnSolde := false;
    end
  else
    if (OkEnSolde and (F.CloDev=FDevisePivot)) then
      if (PbSurMontantDevise(1,Deb,Cre,DebDev,CreDev,F.CloDec)) then
        OKEnSolde := false;
  if (OkEnSolde) then
    begin
    if (AddMvt(3,F,Gen,Aux,Deb,Cre,DebDev,CreDev,F.CloDev,EtatLettrage,Ana,Eche,ModePaie,NumLig,NumPiece,NumEche,forceRef,forceLib,CreerAnaSurAttente,Setfb)) then
      OkInsert := true;
    end
  else
    begin
    if (AddMvt(1,F,Gen,Aux,Deb,Cre,DebDev,CreDev,F.CloDev,EtatLettrage,Ana,Eche,ModePaie,NumLig,NumPiece,NumEche,forceRef,forceLib,CreerAnaSurAttente,Setfb)) then
      OkInsert := true;
    if (AddMvt(2,F,Gen,Aux,Deb,Cre,DebDev,CreDev,F.CloDev,EtatLettrage,Ana,Eche,ModePaie,NumLig,NumPiece,NumEche,forceRef,forceLib,CreerAnaSurAttente,Setfb)) then
      OkInsert := true;
    end;

  Case F.Contrepartie Of
    //-------------
    ChaqueLigne :
      begin
      totDeb      := Cre;
      totCre      := Deb;
      totDebDev   := CreDev;
      totCreDev   := DebDev;
      StotDeb     := totDeb;
      StotCre     := totCre;
      StotDebDev  := totDebDev;
      StotCreDev  := totCreDev;

      if ((Arrondi(totDeb-totCre,4) = 0) or (Arrondi(totDeb-totCre,4) = 0)) then
        begin
        totDeb      := StotDeb;
        totDebDev   := StotDebDev;
        totCre      := 0;
        totCreDev   := 0 ;
        AddMvt(1,F,F.CptContrepGenere,'',totDeb,totCre,totDebDev,totCreDev,F.CloDev,EtatLettrage,Ana,Eche,ModePaie,NumLig,NumPiece,NumEche,'','',false,[]);
        totDeb      := 0;
        totDebDev   := 0;
        totCre      := StotCre;
        totCreDev   := StotCreDev;
        AddMvt(2,F,F.CptContrepGenere,'',totDeb,totCre,totDebDev,totCreDev,F.CloDev,EtatLettrage,Ana,Eche,ModePaie,NumLig,NumPiece,NumEche,'','',false,[]);
        totDeb      := StotDeb;
        totDebDev   := StotDebDev;
        totCre      := StotCre;
        totCreDev   := StotCreDev;
        end
      else
        AddMvt(3,F,F.CptContrepGenere,'',totDeb,totCre,totDebDev,totCreDev,F.CloDev,EtatLettrage,Ana,Eche,ModePaie,NumLig,NumPiece,NumEche,'','',false,[]);
      end;

    //-------------
    EnPiedDC,EnPiedSolde :
      begin
      totDeb      := totDeb+Cre;
      totCre      := totCre+Deb;
      totDebDev   := totDebDev+CreDev;
      totCreDev   := totCreDev+DebDev;
      end;

    end; // FIN CASE

  if (PourPieceCloture) then
    begin
    F.tot.D     := Arrondi(F.tot.D+Cre,FOkDecV);
    F.tot.C     := Arrondi(F.tot.C+Deb,FOkDecV);
    F.totDev.D  := Arrondi(F.totDev.D+CreDev,F.CloDec);
    F.totDev.C  := Arrondi(F.totDev.C+DebDev,F.CloDec);
    end
  else
    begin
    F.tot.D     := Arrondi(F.tot.D+Cre,FOkDecV);
    F.tot.C     := Arrondi(F.tot.C+Deb,FOkDecV);
    F.totDev.D  := Arrondi(F.totDev.D+CreDev,F.CloDec);
    F.totDev.C  := Arrondi(F.totDev.C+DebDev,F.CloDec);
    end;

  if (Coll and (SOldColl = '')) then
    SOldColl := Gen;
  if ((F.ModeGenerePiece=UneParCpt) and (not OkInsert) and (NumLig=0)) then
    SetDecNum(F.TypGenere,F.SAJAL.COMPTEURNORMAL,F.DateGenere);

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2006
Modifié le ... :   /  /
Description .. : FQ 18662
Mots clefs ... :
*****************************************************************}
procedure MajValidationJournaux( vBoCloture : Boolean );
var i : integer;
    lMois, lAnnee, lNombreMois : Word;
    lStPeriode : string[24];
begin
  lStPeriode := '------------------------';
  if vBoCloture then
  begin
    ExecuteSQL('UPDATE JOURNAL SET J_VALIDEEN = J_VALIDEEN1');
    ExecuteSQL('UPDATE JOURNAL SET J_VALIDEEN1 = "' + lStPeriode + '"');
  end
  else
  begin
    NombrePerExo(VH^.EnCours, lMois, lAnnee, lNombreMois);
    for i := 0 to lNombreMois do
     lStPeriode[i] := 'X';

    ExecuteSQL('UPDATE JOURNAL SET J_VALIDEEN1 = J_VALIDEEN');
    ExecuteSQL('UPDATE JOURNAL SET J_VALIDEEN = "' + lStPeriode + '"');
  end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TTraitementCloture.TraitementPVBQE(i: Integer);
var st : String ;
begin
  {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
  if not EstSpecif('51216') or not CloDef then Exit;

  if (ParaClo.AnoCompteBanque = BQUneSeuleRef) then begin
    St := PositionneStUpdate(LCptBQE);
    ChangeRefPointage(i,St);
  end;

  if (ExisteCptPV) then begin
    if (ParaClo.AnoComptePV = SurAnalytique) then
    begin
      St := PositionneStUpdate(LCptPV);
      AnnuleRefPointageAnalytique(i,St);
    end;

    if (ParaClo.AnoComptePV = UneSeuleRef) then
    begin
      St := PositionneStUpdate(LCptPV);
      ChangeRefPointage(i,St);
    end;
  end;
end;

procedure TTraitementCloture.UpdateTiersSurAna ;
Var QS1 : TQuery ;
    lStReq : String ;
begin
  try
    beginTrans;
    lStReq := 'SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_QUALifPIECE,E_AUXILIAIRE FROM ECRITURE Where E_EXERCICE="'+Exo1.Code+'" and E_DATECOMPTABLE>="'+UsDateTime(Exo1.Deb)+'" and E_DATECOMPTABLE<="'+UsDateTime(Exo1.Fin)+'"' ;
    // gestion du mode IFRS
    if modeIFRS
      then lStReq := lStReq + ' and E_QUALifPIECE="I" and E_GENERAL in (SELECT G_GENERAL FROM GENERAUX WHERE G_COLLECTif="X" and G_VENTILABLE="X")'
      else lStReq := lStReq + ' and E_QUALifPIECE="N" and E_GENERAL in (SELECT G_GENERAL FROM GENERAUX WHERE G_COLLECTif="X" and G_VENTILABLE="X")' ;
    QS1 := OpenSql(lStReq,true);
    EcranInitMove( RecordsCount(QS1), '' ) ;
    while (not QS1.Eof) do
      begin
      EcranMoveCur ;
      ExecuteSQL('UPDATE ANALYTIQ SET Y_AUXILIAIRE="'+QS1.FindField('E_AUXILIAIRE').AsString +
                  '" WHERE Y_EXERCICE="' + QS1.FindField('E_EXERCICE').AsString +
                  '" and Y_DATECOMPTABLE="' + UsDateTime(QS1.FindField('E_DATECOMPTABLE').AsDateTime) +
                  '" and Y_JOURNAL="' + QS1.FindField('E_JOURNAL').AsString +
                  '" and Y_NUMEROPIECE=' + IntToStr(QS1.FindField('E_NUMEROPIECE').AsInteger) +
                  ' and Y_NUMLIGNE=' + IntToStr(QS1.FindField('E_NUMLIGNE').AsInteger) +
                  ' and Y_QUALIFPIECE="' + QS1.FindField('E_QUALIFPIECE').AsString + '"') ;
      QS1.Next;
      end;
     Ferme( QS1 );
     EcranFiniMove;
     CommitTrans;
  except
     Rollback;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/07/2003
Modifié le ... : 18/07/2003
Description .. : Etape de traitement de la cloture.
Suite ........ : Utilisable en mode 2-tiers et en mode process server
Mots clefs ... :
*****************************************************************}
function TTraitementCloture.Cloture ( var ClotureOk, OkTraite : Boolean ) : Integer;
var F, F1, F2, Fnil : TFiltreClo ;
    i               : Integer ;
    totCha , totPro : ttDCClo ;
    totResultat     : ttDCClo ;
    tot89Etab       : ttDCClo ;
    totChaDev       : ttDCClo ;
    totProDev       : ttDCClo ;
    totResultatDev  : ttDCClo ;
    Cpt1, Cpt2, st  : String ;
    RetablirEcartDeChange : Boolean ;
    Q1              : TQuery ;
    Q2              : TQuery ;
    Dev             : String3;
    Etab            : String3;
    Exo             : String3;
begin

  Result := CLO_PASERREUR ;

  // 1ère étape --> PREPARATION CLOTURE
  EcranMessage1(0);

  // remplacement de la fonction 'GenereAuxSurAna'
  UpdateTiersSurAna ;
  // fin du remplacement !

  try
    // Gestion Modif IFRS: ne pas refaire traitement effectué lors du 1er passage en mode normal
    if not ModeIFRS then
      begin
      beginTrans;
      DetruitANO;
      end ;

    RetablirEcartDeChange := NbEtabDev;
    EcranUpdateYY( InttoStr(LLEtabDev.Count) ) ;

    {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
    if EstSpecif('51216') then begin
      if (CloDef) and (not modeIFRS) then // Modif SBO IFRS : Deja traité en passage normal
        TraitementPVBQE(0);
    end;

    EcranInitMove( LLEtabDev.Count*5, '' );

    tot89Etab.D := 0;
    tot89Etab.C := 0;

    // Si aucune ligne d'écriture à traiter, le traitement est considéré ok ! 
    if LLEtabDev.Count = 0 then
      OkTraite := True
    else for i := 0 to LLEtabDev.Count-1 do
      begin
      // Initialisation Etablissement et devise
      InitGeneral( F, i ) ;
      EcranMessage2( 0, F.CloLibEtab, F.CloLibDevise );
      InitSupp(F,0);
      if (not CloDef) then
        F.OkGenere := false;
      EcranUpdateXX( InttoStr(i+1) ) ;

      // 2° étape --> CALCULS DES CHARGES
      EcranMessage1(1);
      EcranMoveCur ;
      // *****************************************
      // Calcul et ecritures de cloture sur charge
      // *****************************************
      // Init F
      InitQChaPro( 'CHA', F );
      // Init Query
      Dev   := F.CloDev;
      Etab  := F.CloEtab;
      Exo   := F.CloExo.Code;
      st    := 'SELECT E_GENERAL AS G, " " AS A, SUM(E_DEBIT) AS D, SUM(E_CREDIT) AS C, SUM(E_DEBITDEV) AS DD, SUM(E_CREDITDEV) AS DC FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL' ;
      st    := st + ' WHERE G_NATUREGENE="CHA" AND  E_EXERCICE="'+Exo+'" AND E_ETABLISSEMENT="'+Etab+'" AND E_DEVISE="'+Dev+'"' ;
      // Gestion cloture IFRS
      if modeIFRS
        then st := st + ' AND E_QUALIFPIECE="I" GROUP BY E_GENERAL'
        else st := st + ' AND E_QUALIFPIECE="N" GROUP BY E_GENERAL';
      Q1 := OpenSql(st,true);
      // Calcul et écritures
      CalculANCLO( F, Fnil, false, Q1, true, false );
      Ferme(Q1) ;
      // MAJ cumuls
      Alimtot( totCha, totChaDev, F, false);

      // 3° étape --> CALCULS DES PRODUITS
      EcranEtapeSuivante( 1 ) ;
      EcranMessage1(2);
      EcranMoveCur ;

      // ******************************************
      // Calcul et ecritures de cloture sur produit
      // ******************************************
      // InitF
      InitQChaPro('PRO',F);
      // Init Query
      Dev   := F.CloDev;
      Etab  := F.CloEtab;
      Exo   := F.CloExo.Code;
      st    := 'SELECT E_GENERAL AS G, " " AS A, SUM(E_DEBIT) AS D, SUM(E_CREDIT) AS C, SUM(E_DEBITDEV) AS DD, SUM(E_CREDITDEV) AS DC FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL' ;
      st    := st + ' WHERE G_NATUREGENE="PRO" AND  E_EXERCICE="'+Exo+'" AND E_ETABLISSEMENT="'+Etab+'" AND E_DEVISE="'+Dev+'"' ;
      // Gestion cloture IFRS
      if modeIFRS
        then st := st + ' AND E_QUALIFPIECE="I" GROUP BY E_GENERAL'
        else st := st + ' AND E_QUALIFPIECE="N" GROUP BY E_GENERAL';
      Q1 := OpenSql(st,true);
      // Calcul et écritures
      CalculANCLO(F,Fnil,false,Q1,true,false);
      Ferme(Q1) ;
      // MAJ cumul
      Alimtot(totPro,totProDev,F,false);

      // *****************************************************
      // Calcul et écriture solde résultat + bénéfice ou perte
      // *****************************************************
      // Etape suivante (4)
      EcranEtapeSuivante( 2 ) ;
      EcranMessage1(3);
      EcranMoveCur ;
      // Bénéfice ou perte ?
      if (CalculResultat(totCha,totChaDev,totPro,totProDev,totResultat,totResultatDev,false,F)) then
        begin
        Cpt1 := stFermBen;
        Cpt2 := stOuvBen;
        end
      else
        begin
        Cpt1 := stFermPer;
        Cpt2 := stOuvPer;
        end;
      // Init F
      F1 := F;
      InitSupp(F1,2);
      InitQResultatAN(F1,totResultat,totResultatDev,Cpt1);
      // Insert mouvement provisoire de resultat
      InsertMvtResultat(F1,Cpt1);

      if (not CloDef) then
        F.OkGenere := true;

      // *********************************************************
      // Calcul et écriture solde résultat par (bénéfice ou perte)
      // *********************************************************
      // Init F
      InitSupp(F,0);
      InitQResultatClo(F,totResultat,totResultatDev,Cpt1);
      // Calcul et écritures
      CalculANCLO(F,Fnil,false,nil,false,false);

      // **********************************************************
      // Creation des écriture de cloture et AN sur compte de bilan
      // **********************************************************
      // Etape suivante
      EcranEtapeSuivante( 3 ) ;
      // init F
      InitSupp(F,1);
      F.tot.D := 0;
      F.tot.C := 0;
      F.totDev.D := 0;
      F.totDev.C := 0;
      F.EnSolde := true;
      F.CptContrepGenere := stOuvBil;
      F.Contrepartie := EnPiedSolde;
      F.CreerAna := true;
      // Init Query
      st := 'SELECT E_GENERAL AS G, E_AUXILIAIRE AS A ," " AS X, ';
      {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
      if EstSpecif('51216') then begin
        if (CloDef)
          then st := st + 'SUM(E_DEBIT) as D, SUM(E_CREDIT) as C, SUM(E_DEBITDEV) as DD , SUM(E_CREDITDEV) as DC, E_REFPOINTAGE as REFP '
          else st := st + 'SUM(E_DEBIT) as D, SUM(E_CREDIT) as C, SUM(E_DEBITDEV) as DD , SUM(E_CREDITDEV) as DC ';
      end
      else
        st := st + 'SUM(E_DEBIT) as D, SUM(E_CREDIT) as C, SUM(E_DEBITDEV) as DD , SUM(E_CREDITDEV) as DC ';

      st := st + 'FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL WHERE G_NATUREGENE<>"CHA" and G_NATUREGENE<>"PRO" and G_NATUREGENE<>"EXT" and E_EXERCICE="'+F.CloExo.Code+'" and E_DATECOMPTABLE>="'+UsDatetime(F.CloExo.Deb)+'" and ';
      st := st + 'E_DATECOMPTABLE<="'+USDatetime(F.CloExo.Fin)+'" and E_DEVISE="'+F.CloDev+'" and E_ETABLISSEMENT="'+F.CloEtab+'"' ;
      // Gestion cloture IFRS
      if modeIFRS
        then st := st + ' and E_QUALifPIECE="I" '
        else st := st + ' and E_QUALifPIECE="N" ' ;

      {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
      if EstSpecif('51216') then begin
        if (CloDef)
          then st := st + 'GROUP BY E_GENERAL, E_AUXILIAIRE,E_REFPOINTAGE '
          else st :=  st + 'GROUP BY E_GENERAL, E_AUXILIAIRE ';
      end
      else
        st :=  st + 'GROUP BY E_GENERAL, E_AUXILIAIRE ';

      Q2 := OpenSQL(st,true);

      // Init F pour la recopie
      F2 := F;
      InitSupp(F2,0);
      if (not CloDef) then
        F2.OkGenere := false;
      F2.CreerAna := false;

      // 5° étape
      if (CloDef)
        then EcranMessage1(4)   // cloture définitive --> CREATION DES ECRITURES DE CLOTURE
        else EcranMessage1(5);  // sinon              --> CREATION DES ECRITURES D'A-NOUVEAUX
      EcranMoveCur ;
      QuelParam(F,F2);
      if (not CloDef) then
        F.OkGenere := true;
      // Calcul et écritures
      CalculANCLO(F,F2,CloDef,Q2,false,true);
      Ferme(Q2) ;

      // ***************************************
      // Delete mouvement provisoire de resultat
      // ***************************************
      // Etape suivante
      EcranEtapeSuivante( 4 ) ;
      EcranMessage1(5);
      EcranMoveCur ;
      // Delete base
      DeleteMouvementResultat(F1);

      OkTraite := true;
      EcranEtapeSuivante( 5 ) ;
      end;

    if (RetablirEcartDeChange) then
      ChangeEcartDeCHange(1);
    EcranFiniMove ;

    // ****************
    // CLOTURE DES IFRS
    // ****************
    if EstComptaIFRS and ( not modeIFRS ) and (AuMoinsUneIFRS) then
      begin
      ReinitPourIFRS ;
      if Cloture ( ClotureOk, OkTraite ) <> CLO_PASERREUR then
        raise EAbort.Create('Erreur lors de la cloture des écritures IFRS') ;
      end ;

    { Mise à jour du VISA sur les comptes }
    if CloDef then ExecuteSQL('UPDATE GENERAUX SET G_CYCLEREVISION = "", G_VISAREVISION = "-"');

    // Mise à jour du stock initial en GC
    if ((CloDef) and (not EstBasePclAllegee)) then UpdateStockInitial(VH^.Suivant.Deb);

    // *********************
    // VALIDATION TRAITEMENT
    // *********************
    if modeIFRS
      then modeIFRS := False // On zappe la validation, elle sera faite au retour de la récursivité
      else
        begin
        // Commit
        CommitTrans;
        if not modeAnoDyna then
         begin
          SetParamSoc('SO_CPEXOREF',VH^.Suivant.Code) ;
          SetParamsoc('SO_CPDERNEXOCLO',VH^.EnCours.Code) ;
         end ;
        ClotureOk := true ;
        end ;


  Except
    // ATTENTION : Un incident s'est produit pendant le traitement.La clôture va être annulée.
    if modeIFRS then
      begin
      result := CLO_ERRCLOTUREIFRS ;
      modeIFRS := False ;
      end
    else
      if result <> CLO_ERRCLOTUREIFRS then
        Result := CLO_ERRCLOTURE ;

    Rollback;

    OkTraite := false;
    ClotureOk := false;
  end;

 if ClotureOk and CloDef then
  Transactions(ClotureGCD, 1 ) ; 

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/07/2003
Modifié le ... : 09/06/2006
Description .. : Etape de vérification de la cloture.
Suite ........ : Utilisable en mode 2-tiers et en mode process server
Suite ........ : - LG - 09/06/2006 - FB 18350 - on laisse l'exercice en cpr 
Suite ........ : pour les anao dyna
Mots clefs ... : 
*****************************************************************}
procedure TTraitementCloture.VerifCloture ( var OnSort : Boolean ) ;
var pbRecalculSoldeCompte : Boolean ;
begin

  pbRecalculSoldeCompte := false ;

  // 7° Etape --> VERIFICATION GENERALE
  EcranMessage1(9);

  if (CloDef) then  // Cas cloture définitive
  begin
    {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
    if EstSpecif('51216') then TraitementPVBQE(1);
    MajExo;

  {$IFDEF EAGLCLIENT}
    AvertirCacheServer( 'EXERCICE' ) ; // GCO - 14/06/2007
  {$ENDIF}

    ChargeMagExo(false);
    RecalculSouchePourCloture;

    // GCO - 04/09/2006 - FQ 18662
    MajValidationJournaux( True );

    // Génération des balances de situations dynamiques pour le nouvel exercice
    if GetParamSocSecur('SO_CPBDSDYNA', False ) then
      TraitementSituationDynamique;

    try
      beginTrans;
      MajTotTousComptes(false,'');
      commitTrans;
      OnSort := true;
    except
      OnSort := false ;
      pbRecalculSoldeCompte := true;
      Rollback;
    end;
  end
  else
  begin   // Cas autre cloture
    MajtottousComptes(false,'');
    OnSort := true;
    // if not modeAnoDyna then // *ANODYNA
    MajExo;
  end;

  if (pbRecalculSoldeCompte) then
  begin
    MajtottousComptes(false,'');
    OnSort := true;
  end;
end;

procedure TTraitementCloture.EcranUpdateXX(vStCaption: String);
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFClos ) then
    TFClos(Ecran).UpdateXX( vStCaption ) ;
{$ENDIF}
end;

procedure TTraitementCloture.EcranUpdateYY(vStCaption: String);
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFClos ) then
    TFClos(Ecran).UpdateYY( vStCaption ) ;
{$ENDIF}
end;

procedure TTraitementCloture.EcranEtapeSuivante( vInEtape: Integer );
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFClos ) then
    TFClos(Ecran).EtapeSuivante( vInEtape ) ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/07/2003
Modifié le ... :   /  /    
Description .. : Affectation des attributs de paramètrage de la fermeture 
Suite ........ : d'exercice
Mots clefs ... :
*****************************************************************}
procedure TTraitementCloture.SetParamFerm(vStRef, vStLib, vStJal, vStBil, vStPer, vStBen, vStRes: String; vDtDeb, vDtFin: TDateTime);
begin
  StFermRef := vStRef ;
  StFermLib := vStLib ;
  StFermJal := vStJal;
  StFermBil := vStBil;
  StFermPer := vStPer;
  StFermBen := vStBen;
  StFermRes := vStRes;
  DtFermDeb := vDtDeb;
  DtFermFin := vDtFin;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/07/2003
Modifié le ... :   /  /    
Description .. : Affectation des attributs de paramètrage d'ouverture du 
Suite ........ : nouvel exercice et des ANO
Mots clefs ... :
*****************************************************************}
procedure TTraitementCloture.SetParamOuv(vStRef, vStLib, vStJal, vStBil, vStPer, vStBen: String; vDtDeb, vDtFin: TDateTime);
begin
  StOuvRef := vStRef ;
  StOuvLib := vStLib ;
  StOuvJal := vStJal;
  StOuvBil := vStBil;
  StOuvPer := vStPer;
  StOuvBen := vStBen;
  DtOuvDeb := vDtDeb;
  DtOuvFin := vDtFin;
end;

function TTraitementCloture.VerifParamOk : Integer ;
var i       : Integer ;
    FactClo : String ;
    FactAN  : String ;
begin
    Result := CLO_PASERREUR ;

    i := VerifCpt(StOuvBil);
    if (i > 0) then
      begin
      Result := CLO_ERRCPTBILOUV ;
      Exit;
      end;

    i := VerifCpt(StOuvBen);
    if (i > 0) then
      begin
      REsult := CLO_ERRCPTBENOUV ;
      Exit;
      end;

    i := VerifCpt(StOuvPer);
    if (i > 0) then
      begin
      Result := CLO_ERRCPTPEROUV ;
      Exit;
      end;

    i := VerifJal(StOuvJal,'ANO',FactAN);
    if (i > 0) then
      begin
      Result := CLO_ERRJALOUV ;
      Exit;
      end;

    if (CloDef) then
      begin
      i := VerifCpt(StFermBil);
      if (i > 0) then
        begin
        Result := CLO_ERRCPTBILFER ;
        Exit;
        end;

      i := VerifCpt(StFermBen);
      if (i > 0) then
        begin
        Result := CLO_ERRCPTBENFER ;
        Exit;
        end;

      i := VerifCpt(StFermPer);
      if (i > 0) then
        begin
        Result := CLO_ERRCPTPERFER ;
        Exit;
        end;

      i := VerifJal(StFermJal,'CLO',FactClo);
      if (i > 0) then
        begin
        Result := CLO_ERRJALFER ;
        Exit;
        end;

      i := VerifCpt(StFermRes);
      if (i > 0) then
        begin
        Result := CLO_ERRCPTRES ;
        Exit;
        end;
    end;

    // --------------------- WARNING normalement...
    if (not Auto) then
      if (Multifacturier('ANO',factAN)) then
        begin
        Result := CLO_WNGCOMPTEURJALANO ;
        Exit ;
        end;

    if (not Auto) then
      if (CloDef) then
        if (Multifacturier('CLO',factClo)) then
          begin
          result := CLO_WNGCOMPTEURJALCLO ;
          Exit ;
          end;

end;

procedure TTraitementCloture.SetParamCloture(vParamClo: TParamCloture);
begin
  ParaClo := vParamClo ;
end;

function TTraitementCloture.GetExisteCptPV: boolean;
begin
  Result := ExisteCptPV ;
end;

procedure TTraitementCloture.SetModeServeur(vBoMode: Boolean);
begin
  modeServeur := vBoMode ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/07/2003
Modifié le ... :   /  /
Description .. : Récupération des paramètres de cloture depuis la tob reçu
Suite ........ : par le process server
Mots clefs ... :
*****************************************************************}
procedure TTraitementCloture.AffectParamDepuisTOB(TobParam: TOB);
begin
  // --- Paramètres de de fermeture
  StFermRef := TOBParam.GetValue('FERMREF') ;
  StFermLib := TOBParam.GetValue('FERMLIB') ;
  StFermJal := TOBParam.GetValue('FERMJAL') ;
  StFermBil := TOBParam.GetValue('FERMBIL') ;
  StFermPer := TOBParam.GetValue('FERMPER') ;
  StFermBen := TOBParam.GetValue('FERMBEN') ;
  StFermRes := TOBParam.GetValue('FERMRES') ;
  DtFermDeb := TOBParam.GetValue('FERMDEB') ;
  DtFermFin := TOBParam.GetValue('FERMFIN') ;
  // --- Paramètres de de fermeture
  StOuvRef := TOBParam.GetValue('OUVREF') ;
  StOuvLib := TOBParam.GetValue('OUVLIB') ;
  StOuvJal := TOBParam.GetValue('OUVJAL') ;
  StOuvBil := TOBParam.GetValue('OUVBIL') ;
  StOuvPer := TOBParam.GetValue('OUVPER') ;
  StOuvBen := TOBParam.GetValue('OUVBEN') ;
  DtOuvDeb := TOBParam.GetValue('OUVDEB') ;
  DtOuvFin := TOBParam.GetValue('OUVFIN') ;
  // Paramètres de génération
  ParaClo.CloContrep      := ttContrePartie ( TOBParam.GetValue('CloContrep') ) ;
  ParaClo.ANOContrep      := ttContrePartie ( TOBParam.GetValue('ANOContrep') ) ;
  ParaClo.CloPiece        := ttModeGenerePiece ( TOBParam.GetValue('CloPiece') ) ;
  ParaClo.ANOPiece        := ttModeGenerePiece ( TOBParam.GetValue('ANOPiece') ) ;
  ParaClo.ANOCompteBanque := ttCompteBanque ( TOBParam.GetValue('ANOCompteBanque') ) ;
  ParaClo.ANOComptePV     := ttComptePointableVentilable ( TOBParam.GetValue('ANOComptePV') ) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/07/2003
Modifié le ... :   /  /
Description .. : Création de la tob contenant les paramètres d'appels du 
Suite ........ : process server
Mots clefs ... :
*****************************************************************}
function TTraitementCloture.CreerTobParam: TOB;
Var TobParam : TOB ;
begin
  TOBParam := TOB.Create('ParamCloture', nil, -1) ;

  // Codes des exercice
  TOBParam.AddChampSupValeur('EXO1', Exo1.Code ) ;
  TOBParam.AddChampSupValeur('EXO2', Exo2.Code ) ;

  // CloDef et Auto
  TOBParam.AddChampSupValeur('CloDef', CloDef ) ;
  TOBParam.AddChampSupValeur('Auto', Auto ) ;

  // Contexte
  if ctxPCL in V_PGI.PGIContexte then
    TOBParam.AddChampSupValeur('ctxPCL', True )
  else TOBParam.AddChampSupValeur('ctxPCL', False );

  // Paramètres pour l'appel du process server :
(*
 CA - 17/07/2007 - Suppression car renseigné par CBP (dixit XP)
Avec ce code, on avait des soucis de connexion en authentification NT

{$IFNDEF EAGLSERVER}
  TOBParam.AddChampSupValeur('USERLOGIN' , V_PGI.UserLogin ) ;
  TOBParam.AddChampSupValeur('INIFILE' , HalSocIni ) ;
  TOBParam.AddChampSupValeur('PASSWORD' , V_PGI.Password ) ;
  TOBParam.AddChampSupValeur('DOMAINNAME' , '' ) ;
  TOBParam.AddChampSupValeur('DATEENTREE' , V_PGI.DateEntree ) ;
  TOBParam.AddChampSupValeur('DOSSIER' , V_PGI.CurrentAlias ) ;
{$ENDIF}
*)

  Result := TOBParam ;

end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 17/11/2003
Modifié le ... : 17/11/2003
Description .. : Supprime les écritures de type ANO "A-Nouveaux"
Suite ........ : Retourne le nombre d'écritures supprimées
Mots clefs ... :
*****************************************************************}
function TTraitementCloture.DeleteANO: Integer;
begin
  // Suppression des écritures générales
  Result := ExecuteSQL('DELETE FROM ECRITURE WHERE E_EXERCICE="'+EXO2.Code+'" And E_ECRANOUVEAU="OAN"') ;
  // Suppression des écritures analytiques
  ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_EXERCICE="'+EXO2.Code+'" And Y_ECRANOUVEAU="OAN"') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 17/11/2003
Modifié le ... : 17/11/2003
Description .. : Supprime les écritures d'A-Nouveaux typées "H"
Suite ........ : Retourne le nombre d'écritures supprimées
Mots clefs ... :
*****************************************************************}
function TTraitementCloture.DeleteANOH: Integer;
begin
  // Suppression des écritures générales
  Result := ExecuteSQL('DELETE FROM ECRITURE WHERE E_EXERCICE="'+EXO2.Code+'" And E_ECRANOUVEAU="H"') ;
  // Suppression des écritures analytiques
  ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_EXERCICE="'+EXO2.Code+'" And Y_ECRANOUVEAU="H"') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 17/11/2003
Modifié le ... : 17/11/2003
Description .. : Supprime les écritures d'A-Nouveaux typées "C"
Suite ........ : Retourne le nombre d'écritures supprimées
Mots clefs ... :
*****************************************************************}
function TTraitementCloture.DeleteCLO: Integer;
begin
  // Suppression des écritures générales
  Result := ExecuteSQL('DELETE FROM ECRITURE WHERE E_EXERCICE="'+EXO1.Code+'" And E_ECRANOUVEAU="C"') ;
  // Suppression des écritures analytiques
  ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_EXERCICE="'+EXO1.Code+'" And Y_ECRANOUVEAU="C"') ;
end;

procedure TTraitementCloture.MajExoV8;
var QExo : TQuery;
    Exo , ExoV8 : string;
    bTrouve : boolean;
begin
  ExoV8 := '';  bTrouve := False;
  QExo := OpenSQL ( 'SELECT * FROM EXERCICE WHERE EX_DATEDEBUT<"'+USDateTime(VH^.EnCours.Deb)+'" ORDER BY EX_DATEDEBUT DESC',True );
  try
    while not QExo.Eof do
    begin
      Exo := QExo.FindField('EX_EXERCICE').AsString;
      { L'exercice a t'il des à nouveaux ? }
      if not ExisteSQL('SELECT * FROM ECRITURE WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="H" OR E_ECRANOUVEAU="OAN"') then
      begin
        ExoV8 := Exo;
        bTrouve := True;
      end
      { L'exercice a t'il des à nouveaux H? }
      else if ExisteSQL ( 'SELECT * FROM ECRITURE WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="H"') then
      begin
        ExoV8 := Exo;
        bTrouve := True;
      end
      { Une écriture est-elle lettrée sur les à-nouveaux OAN : suite récup. sisco II }
      else if ExisteSQL  ( 'SELECT * FROM ECRITURE WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="OAN" AND E_LETTRAGE <>""') then
      begin
        { L' exercice à des à-nouveau OAN lettrés ==> on passe ces à-nouveaux à H }
        ExecuteSQL ( 'UPDATE ECRITURE SET E_ECRANOUVEAU="H" WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="OAN"');
        ExecuteSQL ( 'UPDATE ANALYTIQ SET Y_ECRANOUVEAU="H" WHERE Y_EXERCICE="'+Exo+'" AND Y_ECRANOUVEAU="OAN"');
        ExoV8 := Exo;
        bTrouve := True;
      end;
      if bTrouve then break;
      QExo.Next;
    end;
  finally
    Ferme ( QExo );
  end;
  { si exoV8 n'est pas renseigné, dans le cas où le premier exercice est à OAN, on force les à-nouveaux à H }
  if ExoV8 = '' then
  begin
    QExo := OpenSQL ( 'SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT',True);
    if not QExo.Eof then
    begin
      Exo := QExo.FindField('EX_EXERCICE').AsString;
      if ExisteSQL ('SELECT * FROM ECRITURE WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="OAN"') then
      begin
        ExecuteSQL ('UPDATE ECRITURE SET E_ECRANOUVEAU="H" WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="OAN"');
        ExecuteSQL ('UPDATE ANALYTIQ SET Y_ECRANOUVEAU="H" WHERE Y_EXERCICE="'+Exo+'" AND Y_ECRANOUVEAU="OAN"');
        ExoV8 := Exo;
      end;
    end;
    Ferme ( QExo );
  end;
  SetParamSoc ( 'SO_EXOV8', ExoV8 );
end;

procedure TTraitementCloture.MajExoAnnulClo ;
begin
  if cloDef then
    begin
    ExecuteSQL ('UPDATE EXERCICE SET EX_ETATCPTA="OUV" WHERE EX_EXERCICE="'+VH^.Precedent.Code+'"');
    If VH^.Suivant.Code<>'' Then
      ExecuteSQL ('UPDATE EXERCICE SET EX_ETATCPTA="NON" WHERE EX_EXERCICE="'+VH^.Suivant.Code+'"') ;
    end
  else
    ExecuteSQL ('UPDATE EXERCICE SET EX_ETATCPTA="OUV" WHERE EX_EXERCICE="'+VH^.EnCours.Code+'"');
end;

function TTraitementCloture.ExecuteAnnuleCloture( TobParam : TOB ) : TOB ;
var errID     : Integer ;
    onSort    : Boolean ;
begin

  // Préparation de la tob retournée
  Result := TOB.Create('CLOTUREPROCESS', nil, -1) ;
  Result.AddChampSupValeur('RESULT',    CLO_PASERREUR ) ;
  Result.AddChampSupValeur('ONSORT',    TobParam.GetValue('ONSORT') ) ;

  if not modeServeur then
    Exit ;

  // --------------------------------
  // ----- TRAITEMENT DE CLOTURE ----
  // --------------------------------
  errID := AnnuleCloture ( onSort ) ;

  // -----------------------
  // ----- IT'S THE END ----
  // -----------------------
  Result.putValue('RESULT', errID ) ;
  Result.PutValue('ONSORT', onSort ) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 22/06/2007
Modifié le ... :   /  /    
Description .. : - LG - FB 20805 - la decloture ne validait pas ecritures
Mots clefs ... : 
*****************************************************************}
function TTraitementCloture.AnnuleCloture ( var OnSort : Boolean ) : Integer ;
var NbDeleteANO : Integer ;
    OnArrete : Boolean ;
begin

  // on y croit, tou va bien se passer ;-)
  Result   := CLO_PASERREUR ;
  OnArrete := FALSE ;

  Try
    BeginTrans ;

    // ---------------------------------------
    // ----> Suppression des A-Nouveaux <-----
    // ---------------------------------------
    NbDeleteANO := DeleteANO ;

    // -----------------------------------------------------
    // ----> PCL : Suppression des A-Nouveaux typés H <-----
    // -----------------------------------------------------
    // Ajout CA le 09/03/2001
    if ((NbDeleteANO=0)) then
      begin
      if not PresenceLettragePointage
        then NbDeleteANO := DeleteANOH  // Suppression des à-nouveaux typés 'H'
        else
          begin
          Result   := CLO_ERRLETTANO ;
          OnArrete := True;
          end;
      end; // Fin CA le 09/03/2001

    // ---------------------------------------------------
    // ----> PGE : SI AUCUN ANO SUPPRIME, ON ARRETE <-----
    // ---------------------------------------------------
    // Modif CA le 09/03/2001
    if ( not (ctxPCL in V_PGI.PGIContexte) and (NbDeleteANO<=0) ) then
      begin
{$IFNDEF ENTREPRISE}
      Result   := CLO_ERREXOPASCLO ;
{$ENDIF}
      OnArrete := TRUE ;
      end ; // Fin CA le 09/03/2001

    // ---------------------------
    // ----> FIN TRAITEMENT <-----
    // ---------------------------
    if not OnArrete then
      begin
      // == Suppression des ANO typé "C" ==
      DeleteCLO ;
      // == uniquement pour l'annulation de cloture définitive ==
      if CloDef Then
        begin
        // -> Mise à jour de SO_EXOV8 // CA - 22/04/2003
        MajExoV8;
        MajExoAnnulClo;
        AvertirCacheServer( 'EXERCICE' ) ; // SBO : 09/03/2005
        ChargeMagExo(False) ;

        // GCO - 04/09/2006 - FQ 18662
        MajValidationJournaux( False );

        // -> Mise à jour de SO_CPEXOREF // CA - 22/04/2003
        SetParamSoc('SO_CPEXOREF',VH^.EnCours.Code);
        // GCO - 20/07/2005 - FQ 15025
        SetParamSoc('SO_CPDERNEXOCLO', VH^.Precedent.Code);
        // GCO - 19/09/2006 - Validations automatiques des écritures concernées par la décloture
        ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + VH^.EnCours.Code + '" ' +
                 'AND E_DATECOMPTABLE>="' + USDateTime(VH^.EnCours.Deb) + '" ' +
                 'AND E_DATECOMPTABLE<="' + USDateTime(VH^.EnCours.Fin) + '" ' +
                 'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N"');
        { Mise à jour des infos de clôture immo - CA - 20/01/2004 - FQ 13202 }
        if ((ExisteSQL('SELECT I_IMMO FROM IMMO')) and (GetParamSocSecur ('SO_EXOCLOIMMO','')='')) then
         begin
          SetParamSoc('SO_DATECLOTUREIMMO',Date);
          SetParamSoc('SO_EXOCLOIMMO',VH^.Encours.Code);
         end;
        RecalculSouchePourCloture ;
        AvertirCacheServer( 'PARAMSOC' ) ;
        OnSort:=TRUE ;
        end
      else
      // == sinon pour l'annulation de cloture provisoire ==
        BEGIN
        MajExoAnnulClo;
        AvertirCacheServer( 'EXERCICE' ) ; // SBO : 09/03/2005
        END ;

      // == Maj des comptes ==
      If Not OnArrete Then MajTotTousComptes(TRUE,'') ;
      end;

    CommitTrans ;

  Except
    OnSort := FALSE ;
    Result := CLO_ERRANNULCLO ;
    Rollback ;
  end ;

end;

function TTraitementCloture.PresenceLettragePointage: boolean;
begin
  Result := ExisteSQL('SELECT * FROM ECRITURE WHERE E_ECRANOUVEAU="H" AND E_EXERCICE="' + EXO2.Code
                      + '" AND ((E_ETATLETTRAGE<>"RI" AND E_ETATLETTRAGE<>"AL") OR (E_REFPOINTAGE<>""))');
end;

function TTraitementCloture.CreerTobParamCloture: TOB;
Var TobParam : TOB ;
begin
  TOBParam := CreerTobParam ;

  // --- Paramètres de de fermeture
  TOBParam.AddChampSupValeur('FERMREF', StFermRef ) ;
  TOBParam.AddChampSupValeur('FERMLIB', StFermLib ) ;
  TOBParam.AddChampSupValeur('FERMJAL', StFermJal ) ;
  TOBParam.AddChampSupValeur('FERMBIL', StFermBil ) ;
  TOBParam.AddChampSupValeur('FERMPER', StFermPer ) ;
  TOBParam.AddChampSupValeur('FERMBEN', StFermBen ) ;
  TOBParam.AddChampSupValeur('FERMRES', StFermRes ) ;
  TOBParam.AddChampSupValeur('FERMDEB', DtFermDeb ) ;
  TOBParam.AddChampSupValeur('FERMFIN', DtFermFin ) ;

  // --- Paramètres de de fermeture
  TOBParam.AddChampSupValeur('OUVREF', StOuvRef ) ;
  TOBParam.AddChampSupValeur('OUVLIB', StOuvLib ) ;
  TOBParam.AddChampSupValeur('OUVJAL', StOuvJal ) ;
  TOBParam.AddChampSupValeur('OUVBIL', StOuvBil ) ;
  TOBParam.AddChampSupValeur('OUVPER', StOuvPer ) ;
  TOBParam.AddChampSupValeur('OUVBEN', StOuvBen ) ;
  TOBParam.AddChampSupValeur('OUVDEB', DtOuvDeb ) ;
  TOBParam.AddChampSupValeur('OUVFIN', DtOuvFin ) ;

  // Paramètres de génération
  TOBParam.AddChampSupValeur('CloContrep' ,      Integer(ParaClo.CloContrep) ) ;
  TOBParam.AddChampSupValeur('ANOContrep' ,      Integer(ParaClo.ANOContrep) ) ;
  TOBParam.AddChampSupValeur('CloPiece' ,        Integer(ParaClo.CloPiece) ) ;
  TOBParam.AddChampSupValeur('ANOPiece' ,        Integer(ParaClo.ANOPiece) ) ;
  TOBParam.AddChampSupValeur('ANOCompteBanque' , Integer(ParaClo.ANOCompteBanque) ) ;
  TOBParam.AddChampSupValeur('ANOComptePV' ,     Integer(ParaClo.ANOComptePV) ) ;

  Result := TOBParam ;

end;

procedure TTraitementCloture.ReinitPourIFRS;
begin

  modeIFRS    := True ;

  // Init Liste des couples Etablissements / devise
  FreeAndNil( LLEtabDev ) ;
  LLEtabDev := TStringList.Create;
  LLEtabDev.Sorted := false;
  LLEtabDev.Duplicates := DupIgnore;
  LLEtabDev.Clear;

end;

function TTraitementCloture.EstModeIFRS: boolean;
begin
  Result := ModeIFRS ;
end;

function TTraitementCloture.AuMoinsUneIFRS: Boolean;
begin
  Result := ExisteSQL('SELECT E_NUMEROPIECE FROM ECRITURE WHERE E_QUALIFPIECE="I"');
end;

function TTraitementCloture.CanCloseExoGC: boolean;
begin
  if not EstBasePclAllegee then
    result := CanCloseExoFromGC( Exo1.Code )
  else Result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 11/08/2006
Modifié le ... :   /  /
Description .. : Verification de présence d'écritures non validées sur
Suite ........ : l'exercice (NormesNF/BOI)
Mots clefs ... : VALIDE NF BOI
*****************************************************************}
function TTraitementCloture.EcritValidOk: Integer ;
Var
  Qdates : TQuery ;
  DateExo : TExoDate;
begin
    Result := CLO_PASERREUR ;

    Qdates := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE' +
            ' WHERE EX_EXERCICE="' + Exo1.Code + '"', TRUE);
    DateExo.Deb := Qdates.FindField('EX_DATEDEBUT').asDateTime;
    DateExo.Fin := Qdates.FindField('EX_DATEFIN').asDateTime;

    Ferme(QDates);


    If ExisteSQL('SELECT E_VALIDE FROM ECRITURE WHERE E_VALIDE="-"'
          +  ' AND E_EXERCICE="'       + Exo1.Code
          + '" AND E_DATECOMPTABLE>="' + USDateTime(DateExo.Deb)
          + '" AND E_DATECOMPTABLE<="' + USDateTime(DateExo.Fin)
          + '" AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ') then
       Result := CLO_ECRNOVALIDE;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/08/2007
Modifié le ... :   /  /
Description .. : - LG - FB 21271 - on ne supprime plus la table des 
Suite ........ : operation, pour faire une decloture
Mots clefs ... : 
*****************************************************************}
procedure TTraitementCloture.ClotureGCD ;
begin
 //ExecuteSQL('delete CPGCDOPERATION where GOP_NUMERO IN (SELECT CND_NUMEROSEQ FROM CNTVDETAIL WHERE CND_PERM="2" ) ') ;
 //ExecuteSQL('delete CNTVDETAIL where CND_PERM="2" ') ;
 ExecuteSQL('update CNTVDETAIL set CND_PERM="C" WHERE CND_PERM="1" ') ; 
 EXecuteSQL('update CPGCDCUMULS set CGC_SOLDE= CGC_SOLDE+CGC_CREANCE-CGC_REGLEMENT-CGC_PERTE , ' +
            ' CGC_CREANCE=0 , CGC_REGLEMENT = 0 , CGC_PERTE = 0 ,' +
            ' CGC_PROVISION= ( CGC_HT*CGC_TAUX)/100 ' ) ;
//            'cgc_reglement=0,cgc_perte=0,cgc_solde=cgc_provision ,cgc_provision=0') ;
end ;

end.
