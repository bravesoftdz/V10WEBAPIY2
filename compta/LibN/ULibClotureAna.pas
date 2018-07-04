{***********UNITE*************************************************
Auteur  ...... : Compta
Créé le ...... : 05/10/2004
Modifié le ... :   /  /
Description .. : SBO : 05/10/2004 : Regroupe les procédures de clotures
Suite ........ : / déclotures analytiques pour utilisation en process server
Mots clefs ... :
*****************************************************************}
unit ULibClotureAna;

interface

uses
  Windows, Classes, Hctrls,
  Forms,   // TForm
  Ent1,   // TExoDate, TSetFichierBase, TFichierBase, AxeTofb, jbJal, VH, ExoToDates
  HEnt1,  // String3, SyncrDefault, BeginTrans, CommitTrans, RollBack, NumSemaine, V_PGI, EnableControls
{$IFDEF EAGLCLIENT}
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  UTOB,
  SaisUtil,    // SetIncNum
  HStatus,     // InitMove, MoveCur, FinitMove
  SoldeCpt,    // MajTotComptes
  Sysutils, //IntToStr
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  hmsgbox ;


// ------------------------------------------
// - Code erreurs utilisés par les méthodes -
// ------------------------------------------
// Rem :
// - ERR = Erreur
// - WNG = Warning (pas bloquant)
const CLOANA_PASERREUR          =  -1 ;
      CLOANA_CLOTUREOK          =  0 ;
      CLOANA_ERRCLOTURE         =  1 ;
      CLOANA_ERRPROCESSSERVER   =  2 ;

//==================================================
// Definition de Type
//==================================================
Type tTypeCloAna = (EnDetail,EnSolde) ;

Type TGenereANOAna = Record
                     Jal,Ref,Lib,Axe : String ;
                     AvecANO : Boolean ;
                     TypeCloAna : tTypeCloAna ;
                     CloExo : TExoDate ;
                     NewExo : TExoDate ;
                     Etab : String3 ;
                     Souche : String3 ;
                     End ;

Type TFiltreCloAna = Record
                     GeneODA,GeneChaPro : TGenereANOAna ;
                     end ;

//==================================================
// Definition de la classe TTraitementClotureAna
//==================================================
type
    TTraitementClotureAna = class
//----------
      private
//----------
        Ecran         : TForm ;
        modeServeur   : Boolean;          // mode serveur -> déroulement autonome...
//        modeIFRS      : Boolean;          // mode IFRS -> Action sur les écritures "I"...

        // Traitements
        Procedure GenereEcritures  (EnDetail, ChaPro : Boolean ; Var FG : tGenereANOAna ; St : String ) ;
        Function  GenereANODetail  (ChaPro : Boolean ; Var FG : tGenereANOAna ) : Boolean ;
        Function  GenereANOSolde   (ChaPro : Boolean ; Var FG : tGenereANOAna ) : Boolean ;
        Function  GenereANO        (ChaPro : Boolean ; FG : tGenereANOAna ) : Boolean ;

        // Initialisation
        Procedure AffectParamCloDepuisTOB      ( TobParam : TOB ) ;
        Procedure AffectParamAnnulCloDepuisTOB ( TobParam : TOB ) ;

        // Messages Ecran
        function  GestionEcran : Boolean ;
        procedure EcranMessage ( St1, St2, St3 : String ) ;
        procedure EcranInitMove ( vInQte : Integer ; vStTitre : String ) ;
        procedure EcranMoveCur ;
        procedure EcranFiniMove ;
        procedure EcranEnCours ;

        // Tests
//        Function  EstModeIFRS : boolean ;

//----------
      PUBLIC
//----------

        FParamClo        : TFiltreCloAna ;
        FParamAnnulClo   : tGenereANOAna ;
        OkODA            : Boolean ;
        OkChaPro         : Boolean ;
        OkSaisie         : Boolean ;

        // Constructeur / Destructeur
        Constructor Create ( vEcran : TForm; vBoModeServeur : Boolean ) ;

        // Affectation
        Procedure SetParamCloture   ( vParamClo : TFiltreCloAna ) ;
        Procedure SetParamAnnulClo  ( vParamClo : tGenereANOAna ; vOkODA, vOkChaPro, vOkSaisie : Boolean  ) ;

        // Traitements
        function  ClotureAna       : Integer;
        function  AnnuleClotureAna : Integer ;

        // Procedure pour le process server
        Function  CreerTobParamClo        ( vParam : TFiltreCloAna ) : Tob ;
        Function  CreerTobParamAnnulClo   ( vParam : tGenereANOAna ; OkODA, OkChaPro, OkSaisie : Boolean ) : Tob ;
        Function  ExecuteClotureAna       ( TobParam : TOB ) : Tob;
        function  ExecuteAnnuleClotureAna ( TobParam : TOB ) : Tob ;


        end;


//================================================
//========= Fonctions externes ===================
//================================================
Procedure MajTotClotureAna(OkODA,OKChaPro,OkSaisie : Boolean ; AxeODA,AxeChaPro : String) ;
Function  DetruitANOODA(ChaPro : Integer ; Var FG : tGenereANOAna) : Boolean ;


implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  {$IFNDEF EAGLSERVER}
  ClotAna, // pour TFCloAna
  {$ENDIF}
  ULibAnalytique; //RecherchePremDerAxeVentil //SG6 07.03.05

//================================================
//========= Fonctions externes ===================
//================================================
Function FaitStWhere(ChaPro : Integer ; Detruit : Boolean ; Var FG : tGenereANOAna) : String ;
var St : String ;
begin
  Result:='' ;

  if Detruit
    then St := 'Y_EXERCICE ="' + FG.NewExo.Code + '" '
    else St := 'Y_EXERCICE ="' + FG.CloExo.Code + '" ' ;

  if Detruit
    then St := St + ' AND Y_QUALIFPIECE="N" AND (Y_ECRANOUVEAU="OAN") '
    else if FG.AvecAno
           then St := St + ' AND Y_QUALIFPIECE="N" AND (Y_ECRANOUVEAU="N" OR Y_ECRANOUVEAU="OAN") '
           else St := St + ' AND Y_QUALIFPIECE="N" AND (Y_ECRANOUVEAU="N") ' ;

  if not VH^.AnaCroisaxe then
    St := St + ' AND Y_AXE="' + FG.Axe + '" '
  else
    St := St + ' AND Y_AXE="' + 'A' + IntToStr(RecherchePremDerAxeVentil.premier_axe) + '" ';

  if Detruit then
    begin
    St := St + ' AND Y_TYPEANALYTIQUE="X" ' ;
    case Chapro of
      0 : St := St + ' AND Y_TYPEANOUVEAU="AOD" ' ; // Report ODA
      1 : St := St + ' AND Y_TYPEANOUVEAU="ACP" ' ; // Report Charge Produit
      2 : St := St + ' AND Y_TYPEANOUVEAU="ASA" ' ; // ODA anouveau saisie ;
      end ;
    end
  else if ChaPro = 0
         then St := St + 'AND Y_TYPEANALYTIQUE="X" '
         else St := St + ' AND Y_TYPEANALYTIQUE="-" '
                       + ' AND (G_NATUREGENE="CHA" OR G_NATUREGENE="PRO") ' ;

  Result := St ;

end ;


Procedure MajTotClotureAna( OkODA, OKChaPro, OkSaisie : Boolean ; AxeODA, AxeChaPro : String ) ;
Var Setfb : TSetFichierBase ;
    fbAxe : TFichierBase ;
begin

  Setfb := [] ;
  if OkODA Or OkSaisie then
    begin
    FbAxe := AxeTofb(AxeODA) ;
    Setfb := Setfb + [fbAxe] ;
    end ;

  if OkChaPro then
    begin
    FbAxe := AxeTofb(AxeChaPro) ;
    Setfb := Setfb + [fbAxe] ;
    end ;

  Setfb := Setfb + [fbJal] ;

  MajTotComptes(  Setfb,  TRUE,  FALSE,  '' ) ;

end ;

Function DetruitANOODA(ChaPro : Integer ; Var FG : tGenereANOAna) : Boolean ;
Var St      : String ;
    StWhere : String ;
begin
  Result := TRUE ;

  Try
    BeginTrans ;
    StWhere := FaitStWhere( ChaPro, TRUE, FG ) ;
    St      := 'DELETE FROM ANALYTIQ WHERE ' + StWHERE ;

    ExecuteSQL(St) ;

    CommitTrans ;

    Except
      Result:=FALSE ;
      Rollback ;
  End ;

end ;



{ TTraitementClotureAna }

procedure TTraitementClotureAna.AffectParamAnnulCloDepuisTOB(TobParam: TOB);
begin

  with FParamAnnulClo do
    begin
    NewExo.Code := TOBParam.GetValue( 'NEWEXOCODE' ) ;
    Axe         := TOBParam.GetValue( 'AXE' ) ;
    end ;

  OkODA    := TOBParam.GetValue( 'OKODA' ) ;
  OkChaPro := TOBParam.GetValue( 'OKCHAPRO' ) ;
  OkSaisie := TOBParam.GetValue( 'OKSAISIE' ) ;

end;

procedure TTraitementClotureAna.AffectParamCloDepuisTOB(TobParam: TOB);
begin

  // Paramètres de la cloture
  with FParamClo do
    begin

    // Paramètres pour les ODA
    with GeneODA do
      begin
      Jal         := TOBParam.GetValue( 'ODAJAL' ) ;
      Ref         := TOBParam.GetValue( 'ODAREF' ) ;
      Lib         := TOBParam.GetValue( 'ODALIB' ) ;
      Axe         := TOBParam.GetValue( 'ODAAXE' ) ;
      TypeCloAna  := tTypeCloAna( TOBParam.GetValue( 'ODATYPECLOANA' ) ) ;
      CloExo.Code := TOBParam.GetValue( 'ODACLOEXOCODE' ) ;
      CloExo.Deb  := TOBParam.GetValue( 'ODACLOEXODEB' ) ;
      CloExo.Fin  := TOBParam.GetValue( 'ODACLOEXOFIN' ) ;
      NewExo.Code := TOBParam.GetValue( 'ODANEWEXOCODE' ) ;
      NewExo.Deb  := TOBParam.GetValue( 'ODANEWEXODEB' ) ;
      NewExo.Fin  := TOBParam.GetValue( 'ODANEWEXOFIN' ) ;
      Etab        := TOBParam.GetValue( 'ODAETAB' ) ;
      Souche      := TOBParam.GetValue( 'ODASOUCHE' ) ;
      end ;

    // Paramètres pour les ventilations de comptes généraux
    with GeneChaPro do
      begin
      Jal         := TOBParam.GetValue( 'GENEJAL' ) ;
      Ref         := TOBParam.GetValue( 'GENEREF' ) ;
      Lib         := TOBParam.GetValue( 'GENELIB' ) ;
      Axe         := TOBParam.GetValue( 'GENEAXE' ) ;
      TypeCloAna  := tTypeCloAna( TOBParam.GetValue( 'GENETYPECLOANA' ) ) ;
      CloExo.Code := TOBParam.GetValue( 'GENECLOEXOCODE' ) ;
      CloExo.Deb  := TOBParam.GetValue( 'GENECLOEXODEB' ) ;
      CloExo.Fin  := TOBParam.GetValue( 'GENECLOEXOFIN' ) ;
      NewExo.Code := TOBParam.GetValue( 'GENENEWEXOCODE' ) ;
      NewExo.Deb  := TOBParam.GetValue( 'GENENEWEXODEB' ) ;
      NewExo.Fin  := TOBParam.GetValue( 'GENENEWEXOFIN' ) ;
      Etab        := TOBParam.GetValue( 'GENEETAB' ) ;
      Souche      := TOBParam.GetValue( 'GENESOUCHE' ) ;
      end ;

    end ;

end;




function TTraitementClotureAna.AnnuleClotureAna : Integer ;
begin

  result := CLOANA_PASERREUR ;

  try
    BeginTrans ;

    if OkODA    then DetruitANOODA( 0, FParamAnnulClo ) ;
    if OkChaPro then DetruitANOODA( 1, FParamAnnulClo ) ;
    if OkSaisie then DetruitANOODA( 2, FParamAnnulClo ) ;

    MajTotClotureAna( OkODA,
                      OkChaPro,
                      OkSAisie,
                      FParamAnnulClo.Axe,
                      FParamAnnulClo.Axe) ;

    CommitTrans ;

    Except

      result := CLOANA_ERRCLOTURE ;
      Rollback ;

  end ;

end;

function TTraitementClotureAna.ClotureAna : Integer ;
begin

  result := CLOANA_PASERREUR ;

  // Cloture des ODA
  if FParamClo.GeneODA.Jal <> '' Then
    begin
    if not DetruitANOODA(0,FParamClo.GeneODA) then
      result := CLOANA_ERRCLOTURE ;
    if not GenereANO(FALSE,FParamClo.GeneODA) then
      result := CLOANA_ERRCLOTURE ;
    end ;

  // Cloture des comptes charges et produits
  if FParamClo.GeneChaPro.Jal <> '' Then
    begin
    if not DetruitANOODA(1,FParamClo.GeneChaPro) then
      result := CLOANA_ERRCLOTURE ;
    if not GenereANO(TRUE,FParamClo.GeneChaPro) then
      result := CLOANA_ERRCLOTURE ;
    end ;

  // MAJ des totaux
  if result = CLOANA_PASERREUR Then
    begin
    MajTotClotureAna( FParamClo.GeneODA.Jal <> '' ,
                      FParamClo.GeneChaPro.Jal <> '' ,
                      FALSE,
                      FParamClo.GeneODA.Axe,
                      FParamClo.GeneChaPro.Axe ) ;
    end ;

end;

constructor TTraitementClotureAna.Create ( vEcran : TForm; vBoModeServeur : Boolean ) ;
begin
  // sortie Ecran
  Ecran       := vEcran ;
  modeServeur := vBoModeServeur ;
//  modeIFRS    := False ;
end;


function TTraitementClotureAna.CreerTobParamAnnulClo ( vParam : tGenereANOAna ; OkODA, OkChaPro, OkSaisie : Boolean ) : Tob ;
Var TobParam : TOB ;
begin
  TOBParam := TOB.Create('ParamCloture', nil, -1) ;

  // Paramètres de l'annulation de cloture
  TOBParam.AddChampSupValeur('NEWEXOCODE' , vParam.NewExo.Code ) ;
  TOBParam.AddChampSupValeur('AXE' ,        vParam.Axe ) ;
  TOBParam.AddChampSupValeur('OKODA' ,      OkODA ) ;
  TOBParam.AddChampSupValeur('OKCHAPRO' ,   OkChaPro ) ;
  TOBParam.AddChampSupValeur('OKSAISIE' ,   OkSaisie ) ;

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

function TTraitementClotureAna.CreerTobParamClo(vParam: TFiltreCloAna): Tob;
Var TobParam : TOB ;
begin
  TOBParam := TOB.Create('ParamCloture', nil, -1) ;

  // Paramètres de la cloture
  with vPAram do
    begin

    // Paramètres pour les ODA
    with GeneODA do
      begin
      TOBParam.AddChampSupValeur('ODAJAL',        Jal ) ;
      TOBParam.AddChampSupValeur('ODAREF',        Ref ) ;
      TOBParam.AddChampSupValeur('ODALIB',        Lib ) ;
      TOBParam.AddChampSupValeur('ODAAXE',        Axe ) ;
      TOBParam.AddChampSupValeur('ODATYPECLOANA', Integer(TypeCloAna) ) ;
      TOBParam.AddChampSupValeur('ODACLOEXOCODE', CloExo.Code ) ;
      TOBParam.AddChampSupValeur('ODACLOEXODEB',  CloExo.Deb ) ;
      TOBParam.AddChampSupValeur('ODACLOEXOFIN',  CloExo.Fin ) ;
      TOBParam.AddChampSupValeur('ODANEWEXOCODE', NewExo.Code ) ;
      TOBParam.AddChampSupValeur('ODANEWEXODEB',  NewExo.Deb ) ;
      TOBParam.AddChampSupValeur('ODANEWEXOFIN',  NewExo.Fin ) ;
      TOBParam.AddChampSupValeur('ODAETAB',       Etab ) ;
      TOBParam.AddChampSupValeur('ODASOUCHE',     Souche ) ;
      end ;

    // Paramètres pour les ventilations de comptes généraux
    with GeneChaPro do
      begin
      TOBParam.AddChampSupValeur('GENEJAL',        Jal ) ;
      TOBParam.AddChampSupValeur('GENEREF',        Ref ) ;
      TOBParam.AddChampSupValeur('GENELIB',        Lib ) ;
      TOBParam.AddChampSupValeur('GENEAXE',        Axe ) ;
      TOBParam.AddChampSupValeur('GENETYPECLOANA', Integer(TypeCloAna) ) ;
      TOBParam.AddChampSupValeur('GENECLOEXOCODE', CloExo.Code ) ;
      TOBParam.AddChampSupValeur('GENECLOEXODEB',  CloExo.Deb ) ;
      TOBParam.AddChampSupValeur('GENECLOEXOFIN',  CloExo.Fin ) ;
      TOBParam.AddChampSupValeur('GENENEWEXOCODE', NewExo.Code ) ;
      TOBParam.AddChampSupValeur('GENENEWEXODEB',  NewExo.Deb ) ;
      TOBParam.AddChampSupValeur('GENENEWEXOFIN',  NewExo.Fin ) ;
      TOBParam.AddChampSupValeur('GENEETAB',       Etab ) ;
      TOBParam.AddChampSupValeur('GENESOUCHE',     Souche ) ;
      end ;


    end ;

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

procedure TTraitementClotureAna.EcranEnCours;
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFCloAna ) then
    TFCloAna(Ecran).EcranEnCours ;
{$ENDIF}
end;

procedure TTraitementClotureAna.EcranFiniMove;
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  FiniMove ;
{$ENDIF}
end;

procedure TTraitementClotureAna.EcranInitMove(vInQte: Integer; vStTitre: String);
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  InitMove ( vInQte , vStTitre ) ;
{$ENDIF}
end;

procedure TTraitementClotureAna.EcranMessage(St1, St2, St3 : String);
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFCloAna ) then
    TFCloAna(Ecran).Mess( St1, St2, St3 ) ;
{$ENDIF}
end;

procedure TTraitementClotureAna.EcranMoveCur;
begin
  if not GestionEcran then Exit ;
{$IFNDEF EAGLSERVER}
  MoveCur( false );
{$ENDIF}
end;
{
function TTraitementClotureAna.EstModeIFRS: boolean;
begin
  result := modeIFRS ;
end;
}
function TTraitementClotureAna.ExecuteAnnuleClotureAna(TobParam: TOB): Tob;
var errID     : Integer ;
begin

  // Préparation de la tob retournée
  Result := TOB.Create('CLOTUREPROCESS', nil, -1) ;
  Result.AddChampSupValeur('RESULT',    CLOANA_PASERREUR ) ;

  if not modeServeur then
    Exit ;

  // ----------------------------------
  // ----- AFFECTATION DES PARAMS -----
  // ----------------------------------
  AffectParamAnnulCloDepuisTOB ( TobParam ) ;

  // --------------------------------
  // ----- TRAITEMENT DE CLOTURE ----
  // --------------------------------
  errID := AnnuleClotureAna ;
  Result.putValue('RESULT', errID ) ;

end;

function TTraitementClotureAna.ExecuteClotureAna(TobParam: TOB): Tob;
var errID     : Integer ;
begin

  // Préparation de la tob retournée
  Result := TOB.Create('CLOTUREPROCESS', nil, -1) ;
  Result.AddChampSupValeur('RESULT',    CLOANA_PASERREUR ) ;

  if not modeServeur then
    Exit ;

  // ----------------------------------
  // ----- AFFECTATION DES PARAMS -----
  // ----------------------------------
  AffectParamCloDepuisTOB ( TobParam ) ;

  // --------------------------------
  // ----- TRAITEMENT DE CLOTURE ----
  // --------------------------------
  errID := ClotureAna ;
  Result.putValue('RESULT', errID ) ;

end;

function TTraitementClotureAna.GenereANO( ChaPro : Boolean ; FG : tGenereANOAna ) : Boolean ;
begin

  Result := TRUE ;
  if FG.Jal = '' Then Exit ;

  EcranEnCours ;

  if Fg.TypeCloAna = EnDetail
    then Result := GenereANODetail ( ChaPro, FG )
    else Result := GenereANOSolde  ( ChaPro, FG ) ;

end;

function TTraitementClotureAna.GenereANODetail ( ChaPro : Boolean ; var FG : tGenereANOAna ) : Boolean ;
Var StWhere : String  ;
    St      : String  ;
    i       : Integer ;
begin
  Result:=TRUE ;

  Try
    BeginTrans ;

    i:=0 ;
    If Chapro Then i:=1 ;

    StWhere := FaitStWhere(i,FALSE,FG) ;
    St      := 'SELECT Y_ETABLISSEMENT, Y_GENERAL, Y_SECTION, Y_DEBIT D, Y_CREDIT C, Y_REFINTERNE, Y_LIBELLE, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5  FROM ANALYTIQ ' ;
    If ChaPro Then
      St := St + ' LEFT JOIN GENERAUX ON Y_GENERAL=G_GENERAL ' ;
    St := St + ' WHERE ' + StWHERE + ' ORDER BY Y_ETABLISSEMENT,Y_GENERAL,Y_SECTION ' ;

    // Génération des écritures en détail
    GenereEcritures( True, ChaPro, Fg, St ) ;

    CommitTrans ;

    Except
      Result := FALSE ;
      Rollback ;
    End ;

end;

function TTraitementClotureAna.GenereANOSolde ( ChaPro : Boolean ; var FG : tGenereANOAna ) : Boolean ;
Var StWhere : String  ;
    St      : String  ;
    i       : Integer ;
begin

  Result := TRUE ;

  Try
    BeginTrans ;

    i := 0 ;
    if Chapro then i := 1 ;

    StWhere := FaitStWhere( i, FALSE, FG ) ;
    St      := 'SELECT Y_ETABLISSEMENT, Y_GENERAL, Y_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5, SUM(Y_DEBIT) D, SUM(Y_CREDIT) C  FROM ANALYTIQ ' ;
    if ChaPro then
      St := St + ' LEFT JOIN GENERAUX ON Y_GENERAL=G_GENERAL ' ;
    St := St + ' WHERE ' + StWHERE + ' GROUP BY Y_ETABLISSEMENT,Y_GENERAL,Y_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5';

    // Génération des écritures en solde
    GenereEcritures( False, ChaPro, Fg, St ) ;

    CommitTrans ;

    Except
      Result:=FALSE ;
      Rollback ;
  End ;

end;

procedure TTraitementClotureAna.GenereEcritures ( EnDetail, ChaPro : Boolean ; var FG : tGenereANOAna ; St : String ) ;
Var Q       : TQuery ;
    Q1      : TQuery ;
    OldGen  : String ;
    Gen     : String ;
    OldEtab : String ;
    Etab    : String ;
    NumP    : Integer ;
    NumL    : Integer ;
    Sect    : String ;
    sPlan1  : string ;
    sPlan2  : string ;
    sPlan3  : string ;
    sPlan4  : string ;
    sPlan5  : string ;
    DD      : TDateTime ;
begin

  OldGen:='' ;
  OldEtab:='' ;
  NumP:=0 ;
  NumL:=0 ;

  // Query de parcours de l'analytiq
  Q := OpenSQL( St, TRUE ) ;

  // Query pour la création
  Q1 := OpenSQL('SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="' + W_W + '" ', FALSE ) ;

  EcranInitMove( RecordsCount(Q), '' ) ;

  // Parcours de l'analytiq
  while Not Q.Eof do
    begin

    EcranMoveCur ;

    Gen  := Q.FindField('Y_GENERAL').AsString ;
    Etab := Q.FindField('Y_ETABLISSEMENT').AsString ;
    Sect := Q.FindField('Y_SECTION').AsString ;
    sPlan1 := Q.FindField('Y_SOUSPLAN1').AsString;
    sPlan2 := Q.FindField('Y_SOUSPLAN2').AsString;
    sPlan3 := Q.FindField('Y_SOUSPLAN3').AsString;
    sPlan4 := Q.FindField('Y_SOUSPLAN4').AsString;
    sPlan5 := Q.FindField('Y_SOUSPLAN5').AsString;

    EcranMessage(Etab,Gen,Sect) ;

    // Doit-on faire une rupture (cad nouvelle pièce)
    if ( (OldGen<>Gen) or (OldEtab<>Etab) or (NumP=0) ) then
      begin
      DD := 0 ;
      if (FG.NEwExo.Code=VH^.EnCours.Code) or (FG.NEwExo.Code=VH^.Suivant.Code)
        then DD := FG.NEwExo.Deb ;
      // Nouveau n° l pièce
      SetIncNum( EcrAna, FG.Souche, NumP, DD ) ;
      NumL    := 0 ;
      OldGen  := Gen ;
      OldEtab := Etab ;
      end ;

    // Nouvel enregistrement
    Q1.Insert ;
    InitNew( Q1 ) ;

    Inc( NumL ) ; /// N° ligne

    Q1.FindField('Y_GENERAL').AsString         := Gen ;
    Q1.FindField('Y_AXE').AsString             := FG.Axe ;
    Q1.FindField('Y_DATECOMPTABLE').AsDateTime := FG.NEwExo.Deb ;
    Q1.FindField('Y_PERIODE').AsInteger        := GetPeriode(FG.NEwExo.Deb) ;
    Q1.FindField('Y_SEMAINE').AsInteger        := NumSemaine(FG.NEwExo.Deb) ;
    Q1.FindField('Y_NUMEROPIECE').AsInteger    := NumP ;
    Q1.FindField('Y_NUMLIGNE').AsInteger       := 0 ;
    Q1.FindField('Y_SECTION').AsString         := Sect ;
    Q1.FindField('Y_EXERCICE').AsString        := FG.NewExo.Code ;
    Q1.FindField('Y_NATUREPIECE').AsString     := 'OD' ;
    Q1.FindField('Y_QUALIFPIECE').AsString     := 'N' ;
    Q1.FindField('Y_TYPEANALYTIQUE').AsString  := 'X' ;
    Q1.FindField('Y_VALIDE').AsString          := 'X' ;
    Q1.FindField('Y_ETABLISSEMENT').AsString   := Etab ;
    Q1.FindField('Y_DEVISE').AsString          := V_PGI.DevisePivot ;
    Q1.FindField('Y_TAUXDEV').AsFloat          := 1 ;
    Q1.FindField('Y_JOURNAL').AsString         := FG.Jal ;
    Q1.FindField('Y_NUMVENTIL').AsInteger      := NumL ;
    if ChaPro
      then Q1.FindField('Y_TYPEANOUVEAU').AsString := 'ACP'
      else Q1.FindField('Y_TYPEANOUVEAU').AsString := 'AOD' ;
    Q1.FindField('Y_ECRANOUVEAU').AsString     := 'OAN' ;
    Q1.FindField('Y_DEBIT').AsFloat            := Q.FindField('D').AsFloat ;
    Q1.FindField('Y_CREDIT').AsFloat           := Q.FindField('C').AsFloat ;
    Q1.FindField('Y_DEBITDEV').AsFloat         := Q.FindField('D').AsFloat ;
    Q1.FindField('Y_CREDITDEV').AsFloat        := Q.FindField('C').AsFloat ;

    if EnDetail then
      begin
      Q1.FindField('Y_REFINTERNE').AsString      := Q.FindField('Y_REFINTERNE').AsString ;
      Q1.FindField('Y_LIBELLE').AsString         := Q.FindField('Y_LIBELLE').AsString ;
      end
    else
      begin
      Q1.FindField('Y_REFINTERNE').AsString      := FG.Ref ;
      Q1.FindField('Y_LIBELLE').AsString         := FG.Lib ;
      end ;

    //SG6 07.03.05 Mode Croisaxe
    Q1.FindField('Y_SOUSPLAN1').AsString         := sPlan1 ;
    Q1.FindField('Y_SOUSPLAN2').AsString         := sPlan2 ;
    Q1.FindField('Y_SOUSPLAN3').AsString         := sPlan3 ;
    Q1.FindField('Y_SOUSPLAN4').AsString         := sPlan4 ;
    Q1.FindField('Y_SOUSPLAN5').AsString         := sPlan5 ;

    // INSERT en base
    Q1.Post ;

    // Au suivant
    Q.Next ;

    end ;

  EcranFiniMove ;

  Ferme(Q1) ;
  Ferme(Q) ;

end;


function TTraitementClotureAna.GestionEcran: Boolean;
begin
  Result := (not modeServeur) and (Ecran <> nil) ;
end;

procedure TTraitementClotureAna.SetParamAnnulClo  ( vParamClo : tGenereANOAna ; vOkODA, vOkChaPro, vOkSaisie : Boolean  ) ;
begin
  FParamAnnulClo := vParamClo ;
  OkODA          := vOkODA ;
  OkChaPro       := vOkChaPro ;
  OkSaisie       := vOkSaisie ;
end;

procedure TTraitementClotureAna.SetParamCloture ( vParamClo : TFiltreCloAna ) ;
begin
  FParamClo := vParamClo ;
end;

end.
