unit ZEtebac;

interface

Uses StdCtrls, Controls, Classes,forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox,
     // VCL
     Dialogs,                                           
     Windows,
     // Lib
     {$IFNDEF EAGLCLIENT}
     HCompte, // pour le TGGeneral ( utilise par l'analytique manuel )
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}

     {$IFDEF VER150}
     Variants,
     {$ENDIF}

     Ent1,
     ParamSoc, // pour le GetParamSocSecur
     Lookup,
     SaisUtil, // ObjAna
     // Lib
     ULibEcriture,
     UCstEtebac,
     // AGL
     UTOB ;

const

 cPasImput                = '0';
 cImputIncorrecte         = '1';
 cImputCorrectNonValide   = '2';
 cImputValide             = '3';


type


 TZEtebac = Class(TObjetCompta)
  private
   FTOBReleve                : TOB;
   FTOBLigneReleve           : TOB;
   FTOBImput                 : TOB;
   FTOBLigneImputation       : TOB;
   FTOBAnalytiq              : TOB;
   FTOBLigneAnalytiq         : TOB;
//   FTOBLettrage              : TOB ;
   FRechCompte               : TRechercheGrille;  // objet gerant les recherche des compte et des auxiliaires
   FTOBEcr                   : TTOBEcriture; // objet de controle des ecritures
   FMessCompta               : TMessageCompta ;   // objet de gestion des messages
   FStatut                   : TActionFiche;
   FStatutImput              : TActionFiche;
   FStJournalContrepartie    : string; // code du journal de la contrepartie
   FStCompteContrepartie     : string; // Compte bancaire de contrepartie
   FStCompteTVA              : string; // Compte de TVA
   FRdSoldeDebit             : double; // solde des debits des releves
   FRdSoldeCredit            : double; // solde des credits des releves
   FRdSoldeDebitImput        : double; // solde des debits d'imputations
   FRdSoldeCreditImput       : double; // solde des credit d'imputations
   FRdSoldeImput             : double; // solde des imputations
   FStCodeTVA                : string; // code TVA du compte d'imputation
   FCurrentLigneImput        : integer; // ligne en cours des imputations de la ligne de releves
   FListObjAnalytiq          : TStringList;
   FCodeGuide                : string; // code du guide courant
   FStCurrentGuide           : string;
   FTypeContexte             : TTypeContexte;
   FStModeSaisie             : string;
   FStCompteurNormal         : string;
   FDtDateDefaut             : TDateTime;

   procedure TransactionDeleteReleve;
   procedure TransactionSaveLigneReleve;

   function  GetDeviseCompte( value : string) : string;
   procedure SetJournalContrepartie ( Value : string );
   procedure CreationLigne ( vQ : TQuery ; vTOBLettrage : TOB  ) ;
   procedure ReporteEtat ( vBoResult : boolean ) ;

  protected
   function  GetFieldPropertyReleve ( vIndex : integer ) : variant;
   procedure SetFieldPropertyReleve ( vIndex : integer ; Value : variant );
   function  GetFieldPropertyImput  ( vIndex : integer ) : variant;
   procedure SetFieldPropertyImput  ( vIndex : integer ; Value : variant );
   procedure SetFieldAnneeMoisImput ( vIndex : integer ; Value : variant );
   function  GetCodeTVA : string; // recupere le code TVA du compte d'imputation ou le codeTVA du dossier
   function  GetLibelleEtat : string; // libelle de l'etat d'imputation tablette CPETATIMPUTATION
   procedure CalculeSoldeReleve;
   procedure SetGeneralImput( Value : string);
   procedure SetCompteContrepartie( Value : string);
   function  GetDevise : RDevise ;
  public

   constructor Create( vInfoEcr : TInfoEcriture ); override ;
   destructor  Destroy; override;

   procedure InitializeLigneReleve;
   procedure InitializeLigneImput;
   procedure Initialize; override ;
   procedure InitializeObject;

   function  IsValideGeneral ( vStCompte : string ) : boolean; // validation du couple general + auxi
   function  IsValideGeneralImput ( var vStCompte : string ) : boolean; // validation du compte general

   function  IsValideLigneReleve : boolean;
   function  IsValideReleve      : boolean;
   function  IsValideLigneImput  : boolean;
   function  IsValideImputation  : boolean;

   function  SupprimerLigneImputVide : boolean;

   function  AddLigneReleve : TOB;
   procedure AddLigneImput ( vInIndex : integer = -1 );
   function  SaveLigneReleve : boolean;
   function  SaveReleve : boolean;
   function  DeleteLigneReleve : boolean;
   function  DeleteCurrentImputation : boolean;
   function  DeleteCurrentLigneImputation: boolean;
   function  LoadReleve ( Q : TQuery ) : boolean;
   function  LoadImputation : boolean;
   function  LibelleImputation ( Value : string ) : string;
   function  IsCollectif( Value : string ) : boolean;
   function  IsTvaAutorise ( value : string ; vInIndex : integer = 0 ) : boolean;
   procedure AssignInfoTVA;
   procedure AssignInfoCompte;
   function  RechercheImputation : boolean;
   procedure NextImputation;
   procedure FirstImputation;
   procedure CalculSoldeImput;
   procedure RecopierLigneReleve( Value : TOB );
   procedure CreerImputation(vTOBEcr : TOB);
   function  ProchainCompteur : integer; // retourne le prochain RLB_COMPTEUR disponible
   procedure CellExitGenImput( Sender : TObject; var ACol,ARow : integer; var Cancel : boolean ) ;
   procedure CellExitAuxImput( Sender : TObject; var ACol,ARow : integer; var Cancel : boolean ) ;
   procedure ElipsisClick( Sender : TObject ; Key : Word = 0) ;
   procedure CreationAvecCompteAttente ;
   function  CreationDepuisEcriture( vTOBLettrage : TOB ) : boolean ;

   // propriete des lignes de releves
   property InCOMPTEUR_R            : variant index 0  read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StJOURNAL_R             : variant index 1  read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property DtDATECOMPTABLE_R       : variant index 2  read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StREFINTERNE_R          : variant index 3  read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StREFEXTERNE_R          : variant index 4  read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StETAT_R                : variant index 5  read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StGENERAL_R             : variant index 6  read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property RdTAUXTVA_R             : variant index 7  read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property RdDEBIT_R               : variant index 8  read GetFieldPropertyReleve; //        write SetFieldPropertyReleve;
   property RdCREDIT_R              : variant index 9  read GetFieldPropertyReleve; //        write SetFieldPropertyReleve;
   property StLIBELLE_R             : variant index 10 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property InCOMPTEURPERE_R        : variant index 11 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property InTYPE_R                : variant index 12 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StIMPORT_R              : variant index 13 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property DtDATECREATION_R        : variant index 14 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StDEVISE_R              : variant index 15 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StNATUREPIECE_R         : variant index 16 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StETABLISSEMENT_R       : variant index 17 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StAUXILIAIRE_R          : variant index 18 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StGENERALBQE_R          : variant index 19 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StTVASAISIE_R           : variant index 20 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property DtDATEVALEUR_R          : variant index 21 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StCODEAFB_R             : variant index 22 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property StREFPIECE_R            : variant index 23 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   property RdCOTATION_R            : variant index 24 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   {JP 19/01/07 : Pour stocker le Libellé 1}
   property StLibEnrichi1_R         : variant index 25 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   {JP 19/01/07 : Pour stocker le Libellé 2}
   property StLibEnrichi2_R         : variant index 26 read GetFieldPropertyReleve        write SetFieldPropertyReleve;
   {JP 19/01/07 : Pour stocker le Libellé 3}
   property StLibEnrichi3_R         : variant index 27 read GetFieldPropertyReleve        write SetFieldPropertyReleve;

   // propriete des lignes d'imputation
   property InCOMPTEUR_I            : variant index 0  read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StJOURNAL_I             : variant index 1  read GetFieldPropertyImput        write SetFieldPropertyImput;
   property DtDATECOMPTABLE_I       : variant index 2  read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StREFINTERNE_I          : variant index 3  read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StREFEXTERNE_I          : variant index 4  read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StETAT_I                : variant index 5  read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StGENERAL_I             : variant index 6  read GetFieldPropertyImput        write SetFieldPropertyImput;
   property RdTAUXTVA_I             : variant index 7  read GetFieldPropertyImput        write SetFieldPropertyImput;
   property RdDEBIT_I               : variant index 8  read GetFieldPropertyImput; //        write SetFieldPropertyImput;
   property RdCREDIT_I              : variant index 9  read GetFieldPropertyImput; //        write SetFieldPropertyImput;
   property StLIBELLE_I             : variant index 10 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property InCOMPTEURPERE_I        : variant index 11 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property InTYPE_I                : variant index 12 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StIMPORT_I              : variant index 13 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property DtDATECREATION_I        : variant index 14 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StDEVISE_I              : variant index 15 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StNATUREPIECE_I         : variant index 16 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StETABLISSEMENT_I       : variant index 17 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property InNUMLIGNE_I            : variant index 18 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StAUXILIAIRE_I          : variant index 19 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StVENTILABLE_I          : variant index 20 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StCOLLECTIF_I           : variant index 21 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property DtDATEVALEUR_I          : variant index 22 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StGUIDE_I               : variant index 23 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StANA_I                 : variant index 24 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StPERIODE_I             : variant index 25 read GetFieldPropertyImput        write SetFieldAnneeMoisImput;
   property StCONTREPARTIEGEN_I     : variant index 26 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StCONTREPARTIEAUX_I     : variant index 27 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StTVA_I                 : variant index 28 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StCOTATION_I            : variant index 29 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StALettrer_I            : variant index 30 read GetFieldPropertyImput        write SetFieldPropertyImput;
   {JP 17/01/07 : Pour stocker le code pointage}
   property StCodePointage_I        : variant index 31 read GetFieldPropertyImput        write SetFieldPropertyImput;
   {JP 28/02/07 : Pour Reprendre dans l'écriture le cib}
   property StCodeAfb_i             : variant index 32 read GetFieldPropertyImput        write SetFieldPropertyImput;
   {JP 09/03/07 : Ajout des libellés enrichis}
   property StLibEnrichi1_I         : variant index 33 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StLibEnrichi2_I         : variant index 34 read GetFieldPropertyImput        write SetFieldPropertyImput;
   property StLibEnrichi3_I         : variant index 35 read GetFieldPropertyImput        write SetFieldPropertyImput;

   property StJournalContrepartie : string           read FStJournalContrepartie         write SetJournalContrepartie;
   property StCompteContrepartie  : string           read FStCompteContrepartie          write SetCompteContrepartie;
   property RdSoldeDebit          : double           read FRdSoldeDebit;
   property RdSoldeCredit         : double           read FRdSoldeCredit;
   property RdSoldeDebitImput     : double           read FRdSoldeDebitImput             write FRdSoldeDebitImput;
   property RdSoldeCreditImput    : double           read FRdSoldeCreditImput            write FRdSoldeCreditImput;
   property RdSoldeImput          : double           read FRdSoldeImput                  write FRdSoldeImput;
   property StCodeTVA             : string           read GetCodeTVA                     write FStCodeTVA;
   property StLibelleEtat         : string           read GetLibelleEtat;
   property StCompteTVA           : string           read FStCompteTVA;
   property Devise                : RDevise          read GetDevise;

   property Statut                : TActionFiche     read FStatut                        write FStatut;
   property StatutImput           : TActionFiche     read FStatutImput                   write FStatutImput;
   property TOBLigneReleve        : TOB              read FTOBLigneReleve                write FTOBLigneReleve;
   property TOBReleve             : TOB              read FTOBReleve;
   property TOBLigneImputation    : TOB              read FTOBLigneImputation            write FTOBLigneImputation;
   property TOBImput              : TOB              read FTOBImput;
//   property TOBLettrage           : TOB              read FTOBLettrage ;
   property TypeContexte          : TTypeContexte    read FTypeContexte                  write FTypeContexte;
   property StCurrentGuide        : string           read FStCurrentGuide;
   property StModeSaisie          : string           read FStModeSaisie;
   property StCompteurNormal      : string           read FStCompteurNormal;
   property DtDateDefaut          : TDatetime        read FDtDateDefaut;
   {JP 19/01/07 : Gestion des guides sur les libellés enrichis}
//   property StLibEnrichi1         : string           read FStLibEnrichi1                 write FStLibEnrichi1;
  // property StLibEnrichi2         : string           read FStLibEnrichi2                 write FStLibEnrichi2;
   //property StLibEnrichi3         : string           read FStLibEnrichi3                 write FStLibEnrichi3;


  end; // class

procedure TOBEtebacVersTOBEcr ( vTOBEtebac , vTOBECr : TOB ; vInfo : TInfoEcriture ) ;
procedure CInitialiseDateEtebac( var vDateDebut,vDateFin : TDateTime ) ;
procedure CZEteAddChampSupp ( vTOB : TOB ; PourToutesLesFilles : boolean = false ) ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPTypeCons,
  {$ENDIF MODENT1}
  ULibExercice ,
  SaisComm , // pour le teEnCours
  LettUtil ;

const
 cStManqueEtablissement    = 'Il manque le code établissement';
 cStNumeroPieceIncorrecte  = 'Le numéro de folio est incorrecte';

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 05/12/2002
Modifié le ... :   /  /
Description .. : -05/12/2002 - reecriture du code
Mots clefs ... :
*****************************************************************}
procedure CInitialiseDateEtebac( var vDateDebut,vDateFin : TDateTime ) ;
begin

 vDateDebut  := GetParamSocSecur('SO_DATECLOTUREPER',iDate1900) + 1 ;

 if (VH^.DateCloturePer + 1) > VH^.CPExoRef.Deb then
  vDateDebut := VH^.DateCloturePer + 1
   else
    vDateDebut := VH^.CPExoRef.Deb ;

 vDateFin := VH^.CPExoRef.Fin ;

end;

procedure CZEteAddChampSupp ( vTOB : TOB ; PourToutesLesFilles : boolean = false ) ;
begin
 vTOB.AddChampSup('CRL_VENTILABLE'          , PourToutesLesFilles);
 vTOB.AddChampSup('CRL_COLLECTIF'           , PourToutesLesFilles);
 vTOB.AddChampSup('GUIDE'                   , PourToutesLesFilles);
 vTOB.AddChampSup('CRL_NUMGROUPEECR'        , PourToutesLesFilles);
 vTOB.AddChampSupValeur('CRL_ALETTRER'      ,'-', PourToutesLesFilles);
 vTOB.AddChampSupValeur('CRL_LETTRAGE'      ,'',PourToutesLesFilles);
 vTOB.AddChampSupValeur('CRL_ETATLETTRAGE'  ,'',PourToutesLesFilles);
 vTOB.AddChampSupValeur('CRL_COUVERTURE'    ,0 ,PourToutesLesFilles);
 vTOB.AddChampSupValeur('CRL_COUVERTUREDEV' ,0 ,PourToutesLesFilles);
 vTOB.AddChampSupValeur('CRL_DATEPAQUETMAX' ,iDate1900,PourToutesLesFilles);
 vTOB.AddChampSupValeur('CRL_DATEPAQUETMIN' ,iDate1900,PourToutesLesFilles);
end ;



constructor TZEtebac.Create( vInfoEcr : TInfoEcriture );
begin

 inherited ;

 FTOBReleve     := TOB.Create('',nil,-1);        // TOB des lignes de relevés bancaires
// FTOBLettrage   := TOB.Create('',nil,-1);
 FRechCompte    := TRechercheGrille.Create(vInfoEcr) ;
 FTOBEcr        := TTOBEcriture.Create('ECRITURE',nil,-1) ;
 FMessCompta    := TMessageCompta.Create('Saisie de trésorerie') ;
end;

destructor TZEtebac.Destroy;
var
 i              : integer;
begin

 try

  if assigned(FListObjAnalytiq )   then
   begin
    for i := 0 to ( FListObjAnalytiq.Count - 1 ) do
     if assigned(FListObjAnalytiq.Objects[i]) then
      begin
      FListObjAnalytiq.Objects[i].Free;
      FListObjAnalytiq.Objects[i] := nil;
     end; // if

    FListObjAnalytiq.Free;
   end; // if
 if assigned(FTOBReleve)          then FTOBReleve.Free;
 if assigned(FTOBEcr)             then FTOBEcr.Free;
 if assigned(FMessCompta)         then FMessCompta.Free;
 if assigned(FRechCompte)         then FRechCompte.Free;
// if assigned(FTOBLettrage)        then FTOBLettrage.Free;


 FMessCompta                      := nil;
 FTOBReleve                       := nil;
 FListObjAnalytiq                 := nil;
 FRechCompte                      := nil;
 FTOBEcr                          := nil;

 inherited destroy;

 except
  on E : Exception do
   begin
    NotifyError( 'Erreur lors de la destruction !',
                  E.Message ,
                 'TZEtebac.Destroy');
   end;
 end;

end;

function TZEtebac.GetFieldPropertyReleve ( vIndex : integer ) : variant;
begin

 result := '';

 if not assigned(FTOBLigneReleve) then exit;

 case vIndex of
   0 : result := FTOBLigneReleve.GetValue('CRL_COMPTEUR');
   1 : result := FTOBLigneReleve.GetValue('CRL_JOURNAL');
   2 : result := FTOBLigneReleve.GetValue('CRL_DATECOMPTABLE');
   3 : result := FTOBLigneReleve.GetValue('CRL_REFINTERNE');
   4 : result := FTOBLigneReleve.GetValue('CRL_REFEXTERNE');
   5 : result := FTOBLigneReleve.GetValue('CRL_ETAT');
   6 : result := FTOBLigneReleve.GetValue('CRL_GENERAL');
   7 : result := FTOBLigneReleve.GetValue('CRL_TAUXTVA');
   8 : result := FTOBLigneReleve.GetValue('CRL_DEBIT');
   9 : result := FTOBLigneReleve.GetValue('CRL_CREDIT');
  10 : result := FTOBLigneReleve.GetValue('CRL_LIBELLE');
  11 : result := FTOBLigneReleve.GetValue('CRL_COMPTEURPERE');
  12 : result := FTOBLigneReleve.GetValue('CRL_TYPE');
  13 : result := FTOBLigneReleve.GetValue('CRL_IMPORT');
  14 : result := FTOBLigneReleve.GetValue('CRL_DATECREATION');
  15 : result := FTOBLigneReleve.GetValue('CRL_DEVISE');
  16 : result := FTOBLigneReleve.GetValue('CRL_NATUREPIECE');
  17 : result := FTOBLigneReleve.GetValue('CRL_ETABLISSEMENT');
  18 : result := FTOBLigneReleve.GetValue('CRL_AUXILIAIRE');
  19 : result := FTOBLigneReleve.GetValue('CRL_GENERALBQE');
  20 : result := FTOBLigneReleve.GetValue('CRL_TVASAISIE');
  21 : result := FTOBLigneReleve.GetValue('CRL_DATEVALEUR');
  22 : result := FTOBLigneReleve.GetValue('CRL_CODEAFB');
  23 : result := FTOBLigneReleve.GetValue('CRL_REFPIECE');
  24 : result := FTOBLigneReleve.GetValue('CRL_COTATION');
  {JP 19/01/07 : gestion des libellés enrichis lors de la recherche des guides}
  25 : result := FTOBLigneReleve.GetValue('CRL_LIBELLE1');
  26 : result := FTOBLigneReleve.GetValue('CRL_LIBELLE2');
  27 : result := FTOBLigneReleve.GetValue('CRL_LIBELLE3');
 end; // case

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 08/06/2006
Modifié le ... :   /  /
Description .. : - LG - 08/06/2006 - FB 18281 - voir fiche
Mots clefs ... :
*****************************************************************}
function TZEtebac.GetFieldPropertyImput ( vIndex : integer ) : variant;
begin

 result := '';

 if ( not assigned(FTOBLigneImputation) ) or ( FTOBImput = nil ) then exit;

 case vIndex of
  0  : result := FTOBLigneImputation.GetValue('CRL_COMPTEUR');
  1  : result := FTOBLigneImputation.GetValue('CRL_JOURNAL');
  2  : result := FTOBLigneImputation.GetValue('CRL_DATECOMPTABLE');
  3  : result := FTOBLigneImputation.GetValue('CRL_REFINTERNE');
  4  : result := FTOBLigneImputation.GetValue('CRL_REFEXTERNE');
  5  : result := FTOBLigneImputation.GetValue('CRL_ETAT');
  6  : result := FTOBLigneImputation.GetValue('CRL_GENERAL');
 // 7  : result := FTOBLigneImputation.GetValue('CRL_TAUXTVA');
  8  : result := FTOBLigneImputation.GetValue('CRL_DEBIT');
  9  : result := FTOBLigneImputation.GetValue('CRL_CREDIT');
  10 : result := FTOBLigneImputation.GetValue('CRL_LIBELLE');
  11 : result := FTOBLigneImputation.GetValue('CRL_COMPTEURPERE');
  12 : result := FTOBLigneImputation.GetValue('CRL_TYPE');
  13 : result := FTOBLigneImputation.GetValue('CRL_IMPORT');
  14 : result := FTOBLigneImputation.GetValue('CRL_DATECREATION');
  15 : result := FTOBLigneImputation.GetValue('CRL_DEVISE');
  16 : result := FTOBLigneImputation.GetValue('CRL_NATUREPIECE');
  17 : result := FTOBLigneImputation.GetValue('CRL_ETABLISSEMENT');
  18 : result := FTOBLigneImputation.GetValue('CRL_NUMLIGNE');
  19 : result := FTOBLigneImputation.GetValue('CRL_AUXILIAIRE');
  20 : result := FTOBLigneImputation.GetValue('CRL_VENTILABLE');
  21 : result := FTOBLigneImputation.GetValue('CRL_COLLECTIF');
  22 : result := FTOBLigneImputation.GetValue('CRL_DATEVALEUR');
  23 : result := FTOBLigneImputation.GetValue('CRL_GUIDE');
  24 : result := FTOBLigneImputation.GetValue('CRL_ANA');
  25 : result := FTOBLigneImputation.GetValue('CRL_PERIODE');
  26 : result := FTOBLigneImputation.GetValue('CRL_CONTREPARTIEGEN');
  27 : result := FTOBLigneImputation.GetValue('CRL_CONTREPARTIEAUX');
  28 : result := FTOBLigneImputation.GetValue('CRL_TVA');
  29 : result := FTOBLigneImputation.GetValue('CRL_COTATION');
  30 : result := FTOBLigneImputation.GetValue('CRL_ALETTRER');
  {JP 17/01/07 : Pour stocker le code pointage}
  31 : result := FTOBLigneImputation.GetValue('CRL_CODEPOINTAGE');
  {JP 28/02/07 : Pour Reprendre dans l'écriture le cib}
  32 : result := FTOBLigneImputation.GetValue('CRL_CODEAFB');
  {JP 09/03/07 : Ajout des libellés enrichis}
  33 : result := FTOBLigneImputation.GetValue('CRL_LIBELLE1');
  34 : result := FTOBLigneImputation.GetValue('CRL_LIBELLE2');
  35 : result := FTOBLigneImputation.GetValue('CRL_LIBELLE3');
 end; // case

end;


{---------------------------------------------------------------------------------------}
procedure TZEtebac.SetFieldPropertyReleve ( vIndex : integer ; Value : variant );
{---------------------------------------------------------------------------------------}

    {JP 19/01/06 : Gestion des libellés enrichis dans la recheerche des guides
    {----------------------------------------------------------}
    procedure _PutLibelleEnrchi(Ind : Byte);
    {----------------------------------------------------------}
    var
      Chp : string;
    begin
      {Comme il s'agit de champs supplémentaires, il ne sont pas créés lors du chargement de la Tob}
      Chp := 'CRL_LIBELLE' + IntToStr(Ind);
      if FTOBLigneReleve.GetNumChamp(Chp) >= 0 then FTOBLigneReleve.SetString(Chp, Value)
                                               else FTOBLigneReleve.AddChampSupValeur(Chp, Value);
    end;

begin

 if not assigned(FTOBLigneReleve) then
  exit;

 case vIndex of
   0 : FTOBLigneReleve.PutValue('CRL_COMPTEUR' , Value);
   1 : FTOBLigneReleve.PutValue('CRL_JOURNAL', Value);
   2 : if IsValidDate(Value) then
        FTOBLigneReleve.PutValue('CRL_DATECOMPTABLE' , StrToDate(Value));
   3 : FTOBLigneReleve.PutValue('CRL_REFINTERNE', Value);
   4 : FTOBLigneReleve.PutValue('CRL_REFEXTERNE' , Value);
   5 : FTOBLigneReleve.PutValue('CRL_ETAT', Value);
   6 : FTOBLigneReleve.PutValue('CRL_GENERAL', Value);
   7 : if IsNumeric(Value) then
        FTOBLigneReleve.PutValue('CRL_TAUXTVA', Valeur(Value));
  10 : FTOBLigneReleve.PutValue('CRL_LIBELLE', Value);
  11 : FTOBLigneReleve.PutValue('CRL_COMPTEURPERE', Value);
  12 : FTOBLigneReleve.PutValue('CRL_TYPE', Value);
  13 : FTOBLigneReleve.PutValue('CRL_IMPORT', Value);
  14 : if IsValidDate(Value) then
        FTOBLigneReleve.PutValue('CRL_DATECREATION', StrToDate(Value));
  15 : FTOBLigneReleve.PutValue('CRL_DEVISE', Value);
  16 : FTOBLigneReleve.PutValue('CRL_NATUREPIECE', Value);
  17 : FTOBLigneReleve.PutValue('CRL_ETABLISSEMENT', Value);
  18 : FTOBLigneReleve.PutValue('CRL_AUXILIAIRE', Value);
  19 : FTOBLigneReleve.PutValue('CRL_GENERALBQE', Value);
  20 : FTOBLigneReleve.PutValue('CRL_TVASAISIE', Value);
  21 : if IsValidDate(Value) then
        FTOBLigneReleve.PutValue('CRL_DATEVALEUR' , StrToDate(Value));
  22 : FTOBLigneReleve.PutValue('CRL_CODEAFB', Value);
  23 : FTOBLigneReleve.PutValue('CRL_REFPIECE', Value);
  24 : FTOBLigneReleve.PutValue('CRL_COTATION', Value);
  {JP 19/01/07 : gestion des libellés enrichis lors de la recherche des guides}
  25 : _PutLibelleEnrchi(1);
  26 : _PutLibelleEnrchi(2);
  27 : _PutLibelleEnrchi(3);
 end; // case

 if FStatut <> taCreat then
  FStatut := taModif;

end;

{---------------------------------------------------------------------------------------}
procedure TZEtebac.SetFieldPropertyImput ( vIndex : integer ; Value : variant );
{---------------------------------------------------------------------------------------}
begin

 if not assigned(FTOBLigneImputation) then
  exit;

 case vIndex of
  0 : FTOBLigneImputation.PutValue('CRL_COMPTEUR' , Value);
  1 : FTOBLigneImputation.PutValue('CRL_JOURNAL', Value);
  2 : if IsValidDate(Value) then
       FTOBLigneImputation.PutValue('CRL_DATECOMPTABLE' , StrToDate(Value));
  3 : FTOBLigneImputation.PutValue('CRL_REFINTERNE', Value);
  4 : FTOBLigneImputation.PutValue('CRL_REFEXTERNE' , Value);
  5 : FTOBLigneImputation.PutValue('CRL_ETAT', Value);
  6 : SetGeneralImput(Value); // FTOBLigneImput.PutValue('CRL_GENERAL', Value);
  7 : if IsNumeric(Value) then
       FTOBLigneImputation.PutValue('CRL_TAUXTVA', Value)
        else
         FTOBLigneImputation.PutValue('CRL_TAUXTVA', 0);
 10 : FTOBLigneImputation.PutValue('CRL_LIBELLE', Value);
 11 : FTOBLigneImputation.PutValue('CRL_COMPTEURPERE', Value);
 12 : FTOBLigneImputation.PutValue('CRL_TYPE', Value);
 13 : FTOBLigneImputation.PutValue('CRL_IMPORT', Value);
 14 : if IsValidDate(Value) then
       FTOBLigneImputation.PutValue('CRL_DATECREATION', StrToDate(Value));
 15 : FTOBLigneImputation.PutValue('CRL_DEVISE', Value);
 16 : FTOBLigneImputation.PutValue('CRL_NATUREPIECE', Value);
 17 : FTOBLigneImputation.PutValue('CRL_ETABLISSEMENT', Value);
 18 : FTOBLigneImputation.PutValue('CRL_NUMLIGNE', Value);
 19 : FTOBLigneImputation.PutValue('CRL_AUXILIAIRE', Value);
 20 : FTOBLigneImputation.PutValue('CRL_VENTILABLE', Value);
 21 : FTOBLigneImputation.PutValue('CRL_COLLECTIF', Value);
 22 : if IsValidDate(Value) then
       FTOBLigneImputation.PutValue('CRL_DATEVALEUR' , StrToDate(Value));
 23 : FTOBLigneImputation.PutValue('CRL_GUIDE', Value);
 24 : FTOBLigneImputation.PutValue('CRL_ANA', Value);
 26 : FTOBLigneImputation.PutValue('CRL_CONTREPARTIEGEN', Value);
 27 : FTOBLigneImputation.PutValue('CRL_CONTREPARTIEAUX', Value);
 28 : FTOBLigneImputation.PutValue('CRL_TVA', Value);
 29 : FTOBLigneImputation.PutValue('CRL_COTATION', Value);
 30 : FTOBLigneImputation.PutValue('CRL_ALETTRER', Value);
 {JP 17/01/07 : Pour stocker le code pointage}
 31 : FTOBLigneImputation.PutValue('CRL_CODEPOINTAGE', Value);
 {JP 28/02/07 : Pour Reprendre dans l'écriture le cib}
 32 : FTOBLigneImputation.PutValue('CRL_CODEAFB', Value);
 {JP 09/03/07 : Ajout des libellés enrichis}
 33 : FTOBLigneImputation.PutValue('CRL_LIBELLE1', Value);
 34 : FTOBLigneImputation.PutValue('CRL_LIBELLE2', Value);
 35 : FTOBLigneImputation.PutValue('CRL_LIBELLE3', Value);
 end; // case

 if FStatutImput <> taCreat then
  FStatutImput := taModif;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 14/11/2007
Modifié le ... :   /  /    
Description .. : - FB 21717 - LG - 14/11/2007 - affection de l'exo en cours 
Suite ........ : pour le controle sur le visa
Mots clefs ... : 
*****************************************************************}
procedure TZEtebac.InitializeObject;
var
 lDtDateCompable2 : TDateTime ;
 lStCodeExo       : string ;
begin
  // initialisation de l'objet de recherche des comptes
 FRechCompte.OnError    := OnError;
 // init de l'objet ecriture
 FTOBEcr.Info           := Info;
 FTOBEcr.OnError        := OnError;
 CInitialiseDateEtebac(FDtDateDefaut,lDtDateCompable2) ;

 lStCodeExo := ctxExercice.QuelExoDT(FDtDateDefaut) ;

 if lStCodeExo = GetEncours.Code then
  Info.TypeExo := teEncours
   else
    if lStCodeExo = GetSuivant.Code then
     Info.TypeExo := teSuivant
      else
       Info.TypeExo := tePrecedent;
end;

procedure TZEtebac.InitializeLigneReleve;
begin
 InCOMPTEUR_R             := 0;
 StIMPORT_R               := '-';
 StETAT_R                 := cPasImput; // ligne non importé par etabac par defaut
 InTYPE_R                 := integer(TTypeReleve); // releve bancaire par défaut
 StTVASAISIE_R            := '-';
 RdTAUXTVA_R              := 0 ;
 StDEVISE_R               := Info.Devise.Dev.Code ;
 DtDATECREATION_R         := DateToStr(FDtDateDefaut);
 StGENERALBQE_R           := StCompteContrepartie; // compte de contrepartie du journal de banque
end;


procedure TZEtebac.Initialize;
begin

 inherited ;

 FRdSoldeDebit  := 0;
 FRdSoldeCredit := 0;

 if FTOBReleve.Detail.Count > 0 then
  FTOBReleve.ClearDetail;

 FTOBImput := nil;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 02/09/2004
Modifié le ... :   /  /    
Description .. : - LG - 02/09/2004 - FB 14492 - utilisation de usDateTime 
Suite ........ : pour initialiser les dates
Mots clefs ... : 
*****************************************************************}
function TZEtebac.ProchainCompteur : integer;
begin
 // on calcul le prochain numero et on le met à jour en meme temps
 result := CCalculProchainNumeroSouche ('TRE','NOR');
 if result = -1 then
  begin // creation de la souche
   ExecuteSQL('insert into SOUCHE( '                 +
              'SH_TYPE, '                            +
              'SH_SOUCHE, '                          +
              'SH_LIBELLE, '                         +
              'SH_ABREGE, '                          +
              'SH_NATUREPIECE, '                     +
              'SH_NUMDEPART, '                       +
              'SH_JOURNAL, '                         +
              'SH_MASQUENUM, '                       +
              'SH_SOCIETE, '                         +
              'SH_DATEDEBUT, '                       +
              'SH_DATEFIN, '                         +
              'SH_FERME, '                           +
              'SH_ANALYTIQUE, '                      +
              'SH_NATUREPIECEG, '                    +
              'SH_SIMULATION, '                      +
              'SH_NUMDEPARTS, '                      +
              'SH_NUMDEPARTP, '                      +
              'SH_SOUCHEEXO, '                       +
              'SH_RESERVEWEB ) '                     +
              ' values ('                            +
              '"TRE", '                              +
              '"NOR", '                              +
              '"Compteur tréso", '                   +
              '"Tréso", '                            +
              'NULL, '                               +
              '1, '                                  +
              'NULL, '                               +
              'NULL, '                               +
              '001, '                                +
              '"'+UsDateTime(iDate1900)+'", '        +
              '"'+UsDateTime(iDate2099)+'", '        +
              '"-", '                                +
              '"-", '                                +
              '"", '                                 +
              '"-", '                                +
              '1, '                                  +
              '1, '                                  +
              '"-", '                                +
              '"-" ) ');
   result := 1;
  end; // if

end;


function TZEtebac.AddLigneReleve : TOB;
begin

 FTOBLigneReleve                 := TOB.Create('CRELBQE',FTOBReleve,-1);
 FTOBImput                       := nil;
 InitializeLigneReleve;
 Statut                          := taCreat;
 StatutImput                     := taCreat;
 result                          := FTOBLigneReleve;

end;

procedure TZEtebac.TransactionSaveLigneReleve;
begin

  // on supprime toutes les lignes de releve
  ExecuteSQL('delete CRELBQE where CRL_COMPTEUR = ' + varToStr(InCOMPTEUR_R) );
  // on supprime toutes les ligesn d'analytiques
  ExecuteSQL('delete CRELBQEANALYTIQ where CRY_NUMEROPIECE = ' + varToStr(InCOMPTEUR_R) );

  // on sauve l'entete puis les lignes
  FTOBLigneReleve.SetAllModifie(true);
  FTOBLigneReleve.InsertDBByNivel(false);

end;


function TZEtebac.SaveLigneReleve : boolean;
begin

 result := false;

 if Blocage([cStVerrouTreso], false , cStVerrouTreso ) then
  exit;

 try

  if InCOMPTEUR_R = 0 then
   InCOMPTEUR_R := ProchainCompteur;

  {$IFDEF TEST}
  //TOBLigneReleve.SaveToXMLFile('c:\TOBLigneReleve.xml',true,true);
 {$ENDIF}

  result := Transactions(TransactionSaveLigneReleve,1) = oeOK  ;

  if not result then
   NotifyError( cStTexteErreurEnr + varToStr(InCOMPTEUR_R) ,
                V_PGI.LastSQLError,
               'TZEtebac.TransactionSaveLigneReleve' );

  if result then
   begin
    StatutImput                 := taConsult;
    Statut                      := taConsult;
     // calcul du nouveau solde
    FRdSoldeDebit               := FRdSoldeDebit  + RdDEBIT_R;
    FRdSoldeCredit              := FRdSoldeCredit + RdCREDIT_R;
   end; // if

 finally
  Bloqueur( cStVerrouTreso , false );
 end; // try

end;

function TZEtebac.SupprimerLigneImputVide : boolean;
begin

 result := false;

 // on parcourd les lignes pour supprimer les enregistrements vides
 if assigned(FTOBImput) then
  begin

  FTOBLigneImputation := FTOBImput.FindFirst(['CRL_COMPTEUR'],[InCOMPTEUR_R],false);

   while assigned(FTOBLigneImputation) do
    begin

     // on recherche si le compte est ventilable
     if ( RdDEBIT_I = 0 )  and ( RdCREDIT_I = 0 ) then
      begin
       FTOBLigneImputation.Free;
       result := true;
      end; // if

     FTOBLigneImputation := FTOBImput.FindNext(['CRL_COMPTEUR'],[InCOMPTEUR_R],false);

    end; // for

  end; // if

end;

function TZEtebac.SaveReleve : boolean;
var
 i    : integer ;
 lTOB : TOB ;
begin

 for i := 0 to FTOBReleve.Detail.Count - 1 do
  begin
   lTOB := FTOBReleve.Detail[i] ;
   if lTOB.GetValue('CRL_COMPTEUR') = 0 then // pour supprimer un pb de ligne a vide
    continue ;
   lTOB.InsertDB(nil) ;
  end ;

 result := true;

end;

procedure TZEtebac.TransactionDeleteReleve;
begin
 ExecuteSQL('delete CRELBQE where CRL_COMPTEUR = ' + varToStr(InCOMPTEUR_R) );
 ExecuteSQL('delete CRELBQEANALYTIQ where CRY_NUMEROPIECE = ' + varToStr(InCOMPTEUR_R) );
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/07/2001
Modifié le ... :   /  /
Description .. : Supprime la ligne de releve et les imputations associées
Mots clefs ... :
*****************************************************************}
function TZEtebac.DeleteLigneReleve : boolean;
begin

  result := true;

  try

  if Statut <> taCreat then
   begin
    if Blocage([cStVerrouTreso], false , cStVerrouTreso ) then exit;
    // suppression des enregistrement en base
    result := Transactions(TransactionDeleteReleve,1) = oeOK  ;
    if not result then exit;
   end ;

 // suppression de la ligne de releve courante en memoire
 // suppression des lignes d'imputation et des lignes d'analytiques
 FTOBLigneReleve.Free;

 // remise à zéro des pointeurs
 FTOBLigneReleve           := nil;
 FTOBImput                 := nil;
 FTOBLigneImputation       := nil;
 FTOBAnalytiq              := nil;
 FTOBLigneAnalytiq         := nil;

 FStatutImput              := taCreat; // on place les lignes d'imputations en creation pour ne pas envoyer d'ordre sql en base

 FCurrentLigneImput        := 0; // on se place sur la premiere ligne

 if Statut <> taCreat then
  CalculeSoldeReleve;

 FStatut         := taConsult;
 result          := true;

 finally
  if Statut <> taCreat then Bloqueur( cStVerrouTreso , false );
 end; // try


end;

function TZEtebac.IsValideReleve : boolean;
var           
 i : integer;
begin

 result := false;

 // validation de chaque ligne
 for i := 0 to FTOBReleve.Detail.Count - 1 do
  begin
   // faire une fct NextTOB(index);
   FTOBLigneReleve     := FTOBReleve.Detail[i];
   result             := IsValideLigneReleve;
   if not result then
    break;

  end // for

end;

function  TZEtebac.IsValideGeneral ( vStCompte : string ) : boolean;
begin

 TOBEtebacVersTOBEcr(FTOBLigneReleve,FTOBEcr,Info);
 FTOBEcr.PutValue('E_GENERAL',vStCompte);

 result := FTOBEcr.IsValidCompte ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/04/2007
Modifié le ... : 23/04/2007
Description .. : - LG - 19/04/2007 - on affecte le compte meme s'il y a une 
Suite ........ : erreur, pour gere le cas des 401 ss auxi
Suite ........ : - LG - 23/04/2007 - FB 19140 - on ne modifie pas le compte 
Suite ........ : s'il n'est pas correcte et on affiche pas le message
Mots clefs ... : 
*****************************************************************}
function  TZEtebac.IsValideGeneralImput ( var vStCompte : string ) : boolean;
var
 lInError : integer ;
begin
 TOBEtebacVersTOBEcr(FTOBLigneReleve,FTOBEcr,Info);
 FTOBEcr.PutValue('E_GENERAL',vStCompte);
 lInError  := CIsValidCompte(FTOBEcr,Info) ;
 if (lInError = RC_PASERREUR) or ( lInError = RC_AUXOBLIG ) then
  vStCompte := Info.StCompte ; //GetString('E_GENERAL') ;
 if ( lInError <> RC_PASERREUR ) and ( lInError <> RC_AUXOBLIG ) then
  begin
   if ( lInError <> RC_COMPTEINEXISTANT ) then
    NotifyError(lInError,'' ) ;
   result := false ;
  end
   else
     result    := true ;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... : 19/04/2007
Description .. : Validation d'un ligne dans la TOB
Suite ........ : - LG - 10/09/2004 - FB 13040 - on ne pouvait pas saisir de 
Suite ........ : montant < 0
Suite ........ : - LG - 14/09/2004 - FB 14582 - on ouvait saisir des 
Suite ........ : momntants nulls
Suite ........ : - LG - 01/09/2005 - FB 16556 - on ne puvait plus saisir de 
Suite ........ : ligne ss imputation
Suite ........ : - LG - 19/04/2007 - FB 19140 - on reaffecte le general qui a 
Suite ........ : etet complete par la vaalidation
Mots clefs ... : 
*****************************************************************}
function TZEtebac.IsValideLigneReleve : boolean;
var
 lStGeneral : string;
begin

  result       := true;  { TODO : Gestion des montants max et min sur débit/crédit }

  lStGeneral   := StGENERAL_R;
  TOBEtebacVersTOBEcr(FTOBLigneReleve,FTOBEcr,Info);

  if lStGeneral <> '' then
   result := IsValideGeneralImput(lStGeneral) ;

  if not result then exit ;

  StGENERAL_R := lStGeneral ;

  // la date doit être valide
  if not IsValidDate(DtDATECOMPTABLE_R) then
   begin
    result := false;
    NotifyError(cStTexteDateNonValide + VarToStr(DtDATECOMPTABLE_R), '' , 'TZEtebac.IsValideLigneReleve');
    exit;
   end;

  if StJOURNAL_R = '' then
   StJOURNAL_R := StJournalContrepartie;

  result := StJOURNAL_R <> '';

  if not result then exit;

  if StETAT_R = '' then
   StETAT_R := cPasImput;

  // soit le credit ou le debit sont renseigné et supérieur à zéro
  result := ( CIsValidMontant(RdDEBIT_R) = RC_PASERREUR ) and
            ( CIsValidMontant(RdCREDIT_R) = RC_PASERREUR ) ;
  if result then
   result := not ( ( RdDEBIT_R = 0 ) and ( RdCREDIT_R = 0 ) );

  if not result then
   begin
    NotifyError( cStTexteDebitEtCreditNull , '' , 'TZEtebac.IsValideLigneReleve');
    exit;
   end; // if
   
  // recupération du libellé de l'imputation s'il n'est pas définie
  if ( StGENERAL_R <> '' ) and ( StLIBELLE_R = '' ) then
   StLIBELLE_R := LibelleImputation(StGENERAL_R);

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... :   /  /
Description .. : Retourne le libellé du compte general
Mots clefs ... :
*****************************************************************}
function TZEtebac.LibelleImputation ( Value : string ) : string;
begin
 if Info.LoadCompte(Value) then
  result := Info.Compte_GetValue('G_LIBELLE') ;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/07/2001
Modifié le ... : 17/07/2001
Description .. : Chargement des lignes de releves a partir de Q ( cree dans
Suite ........ : un filtre )
Mots clefs ... :
*****************************************************************}
function TZEtebac.LoadReleve ( Q : TQuery ) : boolean;
begin

 Initialize;

 result := FTOBReleve.LoadDetailDB('CRELBQE','','',Q,true);
 if FTOBReleve.Detail.Count > 0 then
  begin
   FTOBLigneReleve     := FTOBReleve.Detail[0];
   CZEteAddChampSupp(FTOBReleve.Detail[0],true) ;
   CalculeSoldeReleve;
   Statut             := taConsult;
  end // if
   else
    FTOBLigneReleve := nil;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 04/07/2001
Modifié le ... :   /  /
Description .. : Récupère le code TVA du compte d'imputation, ou s'il
Suite ........ : n'existe pas le code TVA du dossier
Mots clefs ... :
*****************************************************************}
function TZEtebac.GetCodeTVA: string;
begin
 result := '';
 // on recherche le code tva du compte
 if Info.LoadCompte(StGENERAL_R) then
  result := Info.Compte_GetValue('G_TVA') ;

 // s'il n'existe pas on prend le param par defaut
 if result = '' then
  result := GetParamSocSecur('SO_CODETVAGENEDEFAULT','');

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/07/2001
Modifié le ... :   /  /
Description .. : Affection du taux de tva et du compte de tva en fonction du
Suite ........ : compte d'imputation et du sens du mouvement
Mots clefs ... :
*****************************************************************}
procedure TZEtebac.AssignInfoTVA;
var
 lStCompteTVA   : string ;
 lRdTauxTVA     : double ;
 lStNatureAuxi  : string ;
 lStRegimeTva   : string ;
begin
 lStRegimeTva  := '' ;
 lStNatureAuxi := '' ;
// if FTOBImput = nil then
//  exit ;
 // chargement des info sur le journal
 if Info.Journal.Load([StJournalContrepartie]) = - 1 then exit ;
 if Info.Aux.Load([StAUXILIAIRE_I]) <> - 1 then
  begin
  lStNatureAuxi :=  Info.Aux.GetValue('T_NATUREAUXI') ;
  lStRegimeTva  :=  Info.Aux.GetValue('T_REGIMETVA') ;
  end ;

 Info.Compte.RecupInfoTVA(StGENERAL_R,
                          RdDEBIT_R,
                          StNATUREPIECE_R,
                          Info.Journal.GetValue('J_NATUREJAL') ,
                          lStNatureAuxi ,
                          lStCompteTVA,
                          lRdTauxTVA,
                          lStRegimeTva);

 if lStCompteTVA = '' then
  StTVASAISIE_R := '-'
   else
    begin
     if not ( StTVASAISIE_R = 'X' ) then
      RdTAUXTVA_R   := lRdTauxTVA;
     FStCompteTVA  := lStCompteTVA;
     StTVASAISIE_R := 'X';
    end; // if
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/07/2001
Modifié le ... :   /  /
Description .. : recherche le libelle en clair de l'etat de l'imputation
Mots clefs ... :
*****************************************************************}
function TZEtebac.GetLibelleEtat : string;
begin
 result := RechDom('CPETATIMPUTATION',StETAT_R,false);
end;

procedure TZEtebac.AssignInfoCompte;
begin
 if StLIBELLE_R = '' then StLIBELLE_R := LibelleImputation (StGENERAL_R);
 AssignInfoTVA;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 16/07/2001
Modifié le ... :   /  /
Description .. : Recherche des lignes d'imputations pour la ligne de releves
Suite ........ : courante et ajout à FTOBImput
Mots clefs ... :
*****************************************************************}
function TZEtebac.LoadImputation : boolean;
var
 lQImput          : TQuery;
begin

 result := false;

 // on ne recherche pas les imputations pour les lignes à l'etat 0 : pas de ligne d'imputation
 if StETAT_R = cPasImput then
  exit;

 lQImput := OpenSql ( ' select *'                     +
                      ' from CRELBQE '                +
                      ' where CRL_COMPTEUR ='         + VarToStr( InCOMPTEUR_R ) +
                      ' and CRL_NUMLIGNE <> 0 ' , true);

 if lQImput.EOF then
  begin
   // la ligne etait marqué comme possédant des imputations mais on n'a pas trouver des lignes
   // on remet la ligne comme n'étant pas imputée
   StETAT_R  := cPasImput;
   FTOBImput := nil;
   exit;
  end;

 FTOBImput := TOB.Create( '', FTOBLigneReleve , -1);

 if not FTOBImput.LoadDetailDB('CRELBQE','','',lQImput,true) then
  exit;

 // on garde la requete preparé pour une prochaine utilisation
 Ferme ( lQImput );

 FCurrentLigneImput := 0;

 FTOBLigneImputation := FTOBImput.Detail[0];

 if assigned(FTOBLigneImputation) then
  begin
   CZEteAddChampSupp(FTOBLigneImputation,true) ;
   StatutImput          := taConsult;
  end; // if


 FTOBLigneImputation := FTOBImput.Detail[0];

 result := true;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 16/07/2001
Modifié le ... :   /  /
Description .. : Recherche des lignes d'imputations :
Suite ........ :  1- en memoire
Suite ........ :  2- puis en base
Mots clefs ... :
*****************************************************************}
function TZEtebac.RechercheImputation : boolean;
begin

 // on se place sur la premiere ligne
 FCurrentLigneImput  := 0;

 // recherche de l'imputation en mémoire


 if FTOBLigneReleve.Detail.Count > 0 then
  FTOBImput := FTOBLigneReleve.Detail[0]
   else
    FTOBImput := nil;
 if assigned(FTOBImput) then
  result             := true
   else // recherche dans la base
     result          :=  LoadImputation;


end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 10/12/2002
Modifié le ... :   /  /    
Description .. : - 10/12/2002 - correction d'un fuite memoire sur les 
Suite ........ : AddChampsSupp
Mots clefs ... :
*****************************************************************}
procedure TZEtebac.AddLigneImput ( vInIndex : integer = -1 ) ;
begin
 if FTOBImput = nil then FTOBImput := TOB.Create('',FTOBLigneReleve,-1) ;
 FTOBLigneImputation              := TOB.Create('CRELBQE',FTOBImput,vInIndex);
 CZEteAddChampSupp(FTOBLigneImputation) ;
 InitializeLigneImput;
 StatutImput                     := taCreat;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/07/2001
Modifié le ... :   /  /
Description .. : Initialise les valeurs par défaut d'une ligne de releve
Mots clefs ... :
*****************************************************************}
procedure TZEtebac.InitializeLigneImput;
begin

 InCOMPTEUR_I             := InCOMPTEUR_R;
 StPERIODE_I              := DtDATECOMPTABLE_R;
 StJOURNAL_I              := StJournalContrepartie;
 DtDATECOMPTABLE_I        := DtDATECOMPTABLE_R;
 DtDATECREATION_I         := Date;
 StREFINTERNE_I           := StREFINTERNE_R;
 StREFEXTERNE_I           := StREFEXTERNE_R;
 StLIBELLE_I              := StLIBELLE_R ;
 InTYPE_I                 := integer(TTypeImputation); // releve bancaire par défaut
 DtDateCreation_I         := Date;
 StEtat_I                 := '0'; // non validé par défaut
 InCOMPTEURPERE_I         := InCOMPTEUR_R; // la ligne d'imputation pointe sur la ligne de releve courante
 StANA_I                  := '-'; // pas d'analytique par defaut
 StETABLISSEMENT_I        := StETABLISSEMENT_R;
 StNATUREPIECE_I          := StNATUREPIECE_R;
 StDEVISE_I               := StDEVISE_I;
 StIMPORT_I               := StIMPORT_R ;
 {JP 09/03/07 : Initialisation des libellés enrichis}
 StLibEnrichi1_I          := StLibEnrichi1_R;
 StLibEnrichi2_I          := StLibEnrichi2_R;
 StLibEnrichi3_I          := StLibEnrichi3_R;
end;


function TZEtebac.IsCollectif( Value : string ) : boolean;
begin
 result := Info.Compte.IsCollectif(Value);
end;

function TZEtebac.IsTvaAutorise ( value : string ; vInIndex : integer = 0 ) : boolean;
begin
 result := Info.Compte.IsTvaAutorise(Value,vININdex);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/08/2001
Modifié le ... :   /  /
Description .. : On se place sur la premiere ligne d'imputation du releve
Mots clefs ... :
*****************************************************************}
procedure TZEtebac.FirstImputation;
begin

 // on passe à la ligne suivante
 FCurrentLigneImput := 0;
 if FTOBImput.Detail.Count > 0 then
  FTOBLigneImputation     := FTOBImput.Detail[FCurrentLigneImput]
   else
    FTOBLigneImputation   := nil;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 16/07/2001
Modifié le ... :   /  /
Description .. : Trouve la prochaine ligne d'imputation pour la ligne de
Suite ........ : releve courante
Mots clefs ... :
*****************************************************************}
procedure TZEtebac.NextImputation;
begin

 // on passe à la ligne suivante
 FCurrentLigneImput := FCurrentLigneImput + 1;
 if FCurrentLigneImput < FTOBImput.Detail.Count then
  FTOBLigneImputation     := FTOBImput.Detail[FCurrentLigneImput]
   else
    FTOBLigneImputation   := nil;

end;

procedure TZEtebac.CalculSoldeImput;
var
 ltDC : RecCalcul;
begin

 try

  if assigned(FTOBImput) and (InCOMPTEUR_R <> 0) then  // pour pas que cela plante sur une nouvelle ligne
   begin
    ltDC                 := CCalculSoldePiece ( FTOBImput );
    FRdSoldeDebitImput   := ltDC.D;
    FRdSoldeCreditImput  := ltDC.C;
    FRdSoldeImput        := FRdSoldeDebitImput - FRdSoldeCreditImput;
   end;

 except
  on E : Exception do
   NotifyError( cStTexteErreurSoldeImput, E.Message , 'TZEtebac.CalculSoldeImput' );
 end; // try

end;


procedure TZEtebac.ReporteEtat( vBoResult : boolean ) ;
var
 j : integer ;
begin
 for j := 0 to FTOBImput.Detail.Count - 1 do
  begin
   if vBoResult then
    FTOBImput.Detail[j].PutValue('CRL_ETAT',cImputCorrectNonValide)
     else
      FTOBImput.Detail[j].PutValue('CRL_ETAT',cImputIncorrecte) ;
  end ; // for 
end ;

function TZEtebac.IsValideImputation : boolean;
var
 i : integer;
begin

 result := false;

 if not assigned(FTOBImput) then
  begin
   NotifyError( cStTexteErreurImput , '' ,'TZEtebac.IsValideImputation');
   exit;
  end;

 try

 result := Arrondi(FRdSoldeImput,4) = 0;

 if not result then
  begin
   NotifyError( cStTexteSoldeNonNull + StrfMontant(FRdSoldeImput, 15 , V_PGI.OkDecV, '' , true) , '' ,'TZEtebac.IsValideImputation');
   exit;
  end; // if

 // validation de chaque ligne
 for i := 0 to FTOBImput.Detail.Count - 1 do
  begin
   // faire une fct NextTOB(index);
   FTOBLigneImputation := FTOBImput.Detail[i];
   result             := IsValideLigneImput;
   if i = 0 then
    StGENERAL_R := StGENERAL_I;
   if not result then
    break;

  end // for

 finally
  ReporteEtat(result) ;
 end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 18/07/2001
Modifié le ... : 22/09/2006
Description .. : Suppression des ligne d'imputations de la ligne de releves
Suite ........ : courantes
Suite ........ : - FB 18281 - LG - 22/09/2006 - le echap plantait
Mots clefs ... : 
*****************************************************************}
function TZEtebac.DeleteCurrentImputation : boolean;
begin

// result              := false;
 // on se place sur la premiere ligne
 FCurrentLigneImput  := 0;

 if assigned(FTOBLigneReleve.Detail) then
  FTOBLigneReleve.ClearDetail ;

 //if assigned(FTOBImput) and assigned(FTOBImput.Detail) then
 //  FTOBImput.free;

 FTOBImput           := nil;
 FStatutImput        := taConsult;
 result              := true;

end;


procedure TZEtebac.CalculeSoldeReleve;
var
 i    : integer;
 lTOB : TOB;
begin

 FRdSoldeDebit  := 0;
 FRdSoldeCredit := 0;

 for i := 0 to FTOBReleve.Detail.Count - 1 do
  begin
   // faire une fct NextTOB(index);
   lTOB := FTOBReleve.Detail[i];
   // calcul du solde des debit et credit
   FRdSoldeDebit  := FRdSoldeDebit  + lTOB.GetValue('CRL_DEBIT');
   FRdSoldeCredit := FRdSoldeCredit + lTOB.GetValue('CRL_CREDIT');
  end; // for

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2002
Modifié le ... : 09/12/2002
Description .. : - 28/06/2002 - ajout de la validation des tiers
Suite ........ : - 09/12/2002 - fiche 10718 - la generation auto des guides 
Suite ........ : validait les 
Suite ........ : racines de compte
Mots clefs ... : 
*****************************************************************}
function TZEtebac.IsValideLigneImput : boolean ;
var
 lInError : integer ;
begin
 if (FTypeContexte = TModeAuto) and ( length(StGENERAL_I) < VH^.Cpta[fbGene].Lg ) then
  begin
   NotifyError( 'cette erreur ne doit pas s''afficher', '' , 'TZEtebac.IsValideLigneImput' );
   result := false ;
   exit ;
  end;

 TOBEtebacVersTOBEcr(FTOBLigneImputation,FTOBEcr,Info);

 lInError := CIsValidCompte(FTOBEcr,Info) ;
 result   := lInError = RC_PASERREUR ;
 if not result then
  begin
   NotifyError(lInError,'') ;
   exit ;
  end ;

 lInError := CIsValidAux(FTOBEcr,Info) ;
 result   := lInError = RC_PASERREUR ;
 if not result then
  begin
   NotifyError(lInError,'') ;
   exit ;
  end ;
end;


function TZEtebac.DeleteCurrentLigneImputation : boolean;
begin

// result := false;

 FTOBLigneImputation.free;
 FTOBLigneImputation := nil;
 FStatutImput        := taModif;
 result              := true;
end;



function TZEtebac.GetDeviseCompte( value : string ) : string;
var
 Q : TQuery;
begin
  // recherche de la devise du compte de contrepartie
 Q := nil;

 try

  Q := OpenSQL('select BQ_DEVISE from BANQUECP where BQ_GENERAL = "' +  value
              +'" AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"',true); // 19/10/2006 YMO Multisociétés

  Result := Q.FindField('BQ_DEVISE').asString;

 finally
  if assigned(Q) then Ferme(Q);
 end; //try

end;

procedure TZEtebac.SetGeneralImput( Value : string );
var
 lIndex      : integer;
 lStGeneral  : string;
begin

 FTOBLigneImputation.PutValue('CRL_GENERAL', Value);

 lStGeneral := VarToStr(StGENERAL_I);

 lIndex     := Info.Compte.GetCompte(lStGeneral) ;
 // lIndex := ZComptes.GetCompte(Value);
 if ( lIndex > - 1 ) and ( lStGeneral = Value )  then
  begin

   // est ce que le compte est ventilable
   if Info.Compte.IsVentilable(0, lIndex) then
    StVENTILABLE_I := 'X'
     else
      StVENTILABLE_I := '-';

   // est ce que le compte est collectif
   if Info.Compte.IsCollectif(lIndex) then
    StCOLLECTIF_I := 'X'
     else
      StCOLLECTIF_I := '-';

   if Info.Compte.IsTvaAutorise('',lIndex) then
    StTVA_I := Info.Compte.GetValue('G_TVA', lIndex) ;

  end; //if

end;

procedure TZEtebac.SetCompteContrepartie( Value : string );
begin
 FStCompteContrepartie       := Value;
 Info.Devise.load([GetDeviseCompte(Value)]);
end;



procedure TZEtebac.RecopierLigneReleve ( Value : TOB );
begin
 FTOBLigneReleve := TOB.Create('CRELBQE',FTOBReleve,-1);
 FTOBLigneReleve.Dupliquer(Value,true,true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/07/2007
Modifié le ... :   /  /    
Description .. : - 17/07/2007 - FB 20561 - on affecte l'etablissmetn qd on 
Suite ........ : rajoute une ligne d'imput
Mots clefs ... : 
*****************************************************************}
procedure TZEtebac.CreerImputation ( vTOBEcr : TOB );
var
 lTOBLigneEcr        : TOB;
 lInNumLigne         : integer;
begin

 if vTOBEcr.Detail.Count < 0 then exit;

 FTOBImput := TOB.Create('',FTOBLigneReleve,-1);

 lInNumLigne    := 0;

 lTOBLigneEcr   := vTOBEcr.Detail[lInNumLigne];
 FCodeGuide     := lTOBLigneEcr.GetValue('GUIDE'); // on recupere le code du guide courant

 while assigned(lTOBLigneEcr) do
  begin

  //AddLigneImput;
   FTOBLigneImputation              := TOB.Create('CRELBQE',FTOBImput,-1);
   CZEteAddChampSupp ( FTOBLigneImputation , false ) ;
 {  FTOBLigneImputation.AddChampSup('CRL_VENTILABLE',false);
   FTOBLigneImputation.AddChampSup('CRL_COLLECTIF',false);
   FTOBLigneImputation.AddChampSup('GUIDE',false);
   FTOBLigneImputation.AddChampSup('CRL_NUMGROUPEECR',false);   }
   InitializeLigneImput;
   StatutImput                     := taCreat;

   InNUMLIGNE_I     := lInNumLigne + 1;

   StGENERAL_I      := lTOBLigneEcr.GetValue('CRL_GENERAL');
   StLIBELLE_I      := lTOBLigneEcr.GetValue('CRL_LIBELLE');
   StAUXILIAIRE_I   := lTOBLigneEcr.GetValue('CRL_AUXILIAIRE');
   StETABLISSEMENT_I:= lTOBLigneEcr.GetValue('CRL_ETABLISSEMENT');
   StETABLISSEMENT_I:= lTOBLigneEcr.GetValue('CRL_ETABLISSEMENT');
   CSetMontants( FTOBLigneImputation ,
                 lTOBLigneEcr.GetValue('CRL_DEBIT') ,
                 lTOBLigneEcr.GetValue('CRL_CREDIT') ,
                 Info.Devise.Dev,
               //  false,
                 true );

   StREFINTERNE_I   := lTOBLigneEcr.GetValue('CRL_REFINTERNE');
   StGUIDE_I        := lTOBLigneEcr.GetValue('GUIDE');
   FStCurrentGuide  := StGUIDE_I;

   {JP 09/03/07 : Ajout des libellés enrichis}
   StLibEnrichi1_I  := lTOBLigneEcr.GetValue('CRL_LIBELLE1');
   StLibEnrichi2_I  := lTOBLigneEcr.GetValue('CRL_LIBELLE2');
   StLibEnrichi3_I  := lTOBLigneEcr.GetValue('CRL_LIBELLE3');

   Inc(lInNumLigne);

   if lInNumLigne < vTOBEcr.Detail.Count then
    lTOBLigneEcr     := vTOBEcr.Detail[lInNumLigne]
     else
      begin
       break;
       //lTOBLigneEcr   := nil;
      end;

  end; // while

end;



procedure TZEtebac.SetJournalContrepartie( Value : string );
begin

 FStJournalContrepartie := Value;

 if Info.Journal.Load([Value])<>-1 then
  begin
   FStModeSaisie        := Info.Journal.GetValue('J_MODESAISIE');
   FStCompteurNormal    := Info.Journal.GetValue('J_COMPTEURNORMAL');
  end; // if

end;

{procedure TZEtebac.NotifyError( vMessage, vDelphiMessage , vFunctionName : string);
var
 lStMessage : string;
begin

 if vDelphiMessage = '' then
  lStMessage   := vMessage
   else
    lStMessage := vMessage + #13#10 + vDelphiMessage + #13#10 + vFunctionName;

 if ( V_PGI.LastSQLError <> '') or ( FTypeContexte <> TModeAuto ) then
  MessageAlerte( lStMessage );

end;  }

procedure TZEtebac.SetFieldAnneeMoisImput( vIndex : integer ; Value : variant);
begin
 FTOBLigneImputation.PutValue('CRL_PERIODE', intToStr( GetPeriode(Value) ) );
end;


{procedure TZEtebac.OnError (sender : TObject; Error : TRecError ) ;
begin
 if FTypeContexte = TModeAuto then exit ;
 if trim(Error.RC_Message)<>'' then PGIInfo(Error.RC_Message, 'Saisie de trésorerie' )
 else if (Error.RC_Error<>RC_PASERREUR) then FMessCompta.Execute(Error.RC_Error) ;
end;}


procedure TZEtebac.CellExitAuxImput(Sender: TObject; var ACol, ARow: integer;var Cancel: boolean);
begin
 FRechCompte.COL_GEN    := cColGeneralImput;
 FRechCompte.COL_AUX    := cColAuxiliaireImput;
 TOBEtebacVersTOBEcr(FTOBLigneImputation,FTOBEcr,Info);
 FRechCompte.Ecr        := FTOBEcr ;
 Info.AjouteErrIgnoree ([RC_MODEREGLE] ) ;
 FRechCompte.CellExitAux(Sender,ACol,ARow,Cancel) ;
end;

procedure TZEtebac.CellExitGenImput(Sender: TObject; var ACol, ARow: integer;var Cancel: boolean);
begin
 FRechCompte.COL_GEN    := cColGeneralImput;
 FRechCompte.COL_AUX    := cColAuxiliaireImput;
 TOBEtebacVersTOBEcr(FTOBLigneImputation,FTOBEcr,Info);
 FRechCompte.Ecr        := FTOBEcr ;
 FRechCompte.CellExitGen(Sender,ACol,ARow,Cancel) ;
end;


procedure TZEtebac.ElipsisClick(Sender: TObject; Key: Word=0);
begin
 FRechCompte.COL_GEN    := cColGeneralImput;
 FRechCompte.COL_AUX    := cColAuxiliaireImput;
 TOBEtebacVersTOBEcr(FTOBLigneImputation,FTOBEcr,Info);
 FRechCompte.Ecr        := FTOBEcr ;
 FRechCompte.ElipsisClick(Sender) ;
end;

procedure TOBEtebacVersTOBEcr ( vTOBEtebac , vTOBECr : TOB ; vInfo : TInfoEcriture ) ;
begin
 if ( vTOBEtebac = nil ) or ( vTOBEcr = nil ) then raise exception.Create('TOBEtebacVersTOBEcr : les deux TOB doivent être assignées' ) ;
 vTOBEcr.PutValue( 'E_JOURNAL'         , vTOBEtebac.GetValue('CRL_JOURNAL') );
 vTOBEcr.PutValue( 'E_GENERAL'         , vTOBEtebac.GetValue('CRL_GENERAL') );
 vTOBEcr.PutValue( 'E_AUXILIAIRE'      , vTOBEtebac.GetValue('CRL_AUXILIAIRE') );
 vTOBEcr.PutValue( 'E_DATECOMPTABLE'   , vTOBEtebac.GetValue('CRL_DATECOMPTABLE') ) ;
 vTOBEcr.PutValue( 'E_NATUREPIECE'     , vTOBEtebac.GetValue('CRL_NATUREPIECE') ) ;
 vTOBEcr.PutValue( 'E_DEVISE'          , vTOBEtebac.GetValue('CRL_DEVISE') ) ;
 CSetMontants( vTOBEcr ,
               vTOBEtebac.GetValue('CRL_DEBIT') ,
               vTOBEtebac.GetValue('CRL_CREDIT') ,
               vInfo.Devise.Dev,
               //  false,
               true );

end;

function TZEtebac.GetDevise : RDevise;
begin
 result := Info.Devise.Dev ;
end;



procedure TZEtebac.CreationAvecCompteAttente ;
var
 lBoResult : boolean ;
begin

 if GetParamSocSecur('SO_GENATTEND','') = '' then exit ;

 // ligne sur le compte d'attente
 AddLigneImput ;
 InNUMLIGNE_I     := 1 ;
 StGENERAL_I      := GetParamSocSecur('SO_GENATTEND','') ;
 StAUXILIAIRE_I   := '' ;
 CSetMontants( FTOBLigneImputation ,
               RdDEBIT_R ,
               RdCREDIT_R ,
               Info.Devise.Dev,
               true );

 // ligne sur le compte de bqe
 AddLigneImput ;
 InNUMLIGNE_I     := 2 ;
 StGENERAL_I      := StGENERALBQE_R ;
 StAUXILIAIRE_I   := '' ;
 CSetMontants( FTOBLigneImputation ,
               RdCREDIT_R ,
               RdDEBIT_R ,
               Info.Devise.Dev,
               true );

 CalculSoldeImput;

 lBoResult := IsValideImputation ;
 ReporteEtat(lBoResult) ;

 if lBoResult then
  StETAT_R := cImputCorrectNonValide
   else
    StETAT_R := cImputIncorrecte ;

end;

function _ChargeEcriture(  vQ : TQuery ; vTOBLettrage : TOB ) : TOB ;
var
 lQ : TQuery ;
begin

 lQ := OpenSql('SELECT * FROM ECRITURE WHERE ' +
               '(E_JOURNAL = "' + vQ.FindField('E_JOURNAL').asString + '") AND ' +
               '(E_EXERCICE = "' + vQ.FindField('E_EXERCICE').asString + '") AND ' +
               '(E_DATECOMPTABLE = "' + USDateTime(vQ.FindField('E_DATECOMPTABLE').asDateTime) + '") AND ' +
               '(E_NUMEROPIECE = ' + IntToStr(vQ.FindField('E_NUMEROPIECE').asInteger) + ') AND ' +
               '(E_NUMLIGNE = ' + IntToStr(vQ.FindField('E_NUMLIGNE').asInteger) + ') AND ' +
               '(E_NUMECHE = ' + IntToStr(vQ.FindField('E_NUMECHE').asInteger) + ') AND ' +
               '(E_QUALIFPIECE = "' + vQ.FindField('E_QUALIFPIECE').asString + '")', False);

 result := TOB.Create('ECRITURE',vTOBLettrage,-1 ) ;
 result.SelectDB('',lQ) ;
 Ferme(lQ) ;

end ;

function _GetCredit ( vDebit : double ; vDate : TDateTime ) : TQuery ;
begin
 result := OpenSql('select distinct e_exercice,e_journal,e_general,e_auxiliaire,e_datecomptable, ' +
               ' e_numeropiece,e_numligne,E_NUMECHE,E_QUALIFPIECE from ecriture, journal ' +
               ' where ' +
               ' e_journal=j_journal ' +
               ' and e_etatlettrage="AL" ' +
               ' and e_datecomptable <"' + UsDatetime(vDate) + '" ' +
               ' and e_credit= ' + VariantToSql(vDebit)  +
               ' and ( ( e_general like "40%" and j_naturejal="ACH" ) or ' +
               ' ( e_general like "42%" and ( j_naturejal="REG" or j_naturejal="OD" ) ) or ' +
               ' ( e_general like "43%" and ( j_naturejal="REG" or j_naturejal="OD" ) ) ) ' +
               ' group by e_exercice,e_journal,e_general,e_auxiliaire,e_datecomptable, ' +
               ' e_numeropiece,e_numligne,E_NUMECHE,E_QUALIFPIECE ' +
               ' having count(*) = 1 ' , true ) ;
end ;

function _GetDebit ( vCredit : double ; vDate : TDateTime ) : TQuery ;
begin
 result := OpenSql('select distinct e_exercice,e_journal,e_general,e_auxiliaire,e_datecomptable, ' +
               ' e_numeropiece,e_numligne,E_NUMECHE,E_QUALIFPIECE from ecriture, journal ' +
               ' where ' +
               ' e_journal=j_journal ' +
               ' and e_debit= ' + VariantToSql(vCredit)  +
               ' and ( e_general like "41%" and j_naturejal="VTE" ) ' +
               ' and e_etatlettrage="AL" ' +
               ' and e_datecomptable <"' + UsDatetime(vDate) + '" ' +
               ' group by e_exercice,e_journal,e_general,e_auxiliaire,e_datecomptable, ' +
               ' e_numeropiece,e_numligne,E_NUMECHE,E_QUALIFPIECE ' +
               ' having count(*) = 1 ' , true ) ;
end ;

procedure CLettrageDeuxLigne( vTOB1 , vTOB2 : TOB ) ;
var
 lStCodeL       : string ;
 lP1,lP2        : string ;
 lRdCouvCur     : double ;
 lRdCouvDev     : double ;
 lDtDateMax     : TDateTime ;
 lDtDateMin     : TDateTime ;
begin

 lP1         := TableToPrefixe(vTOB1.NomTable) ;
 lP2         := TableToPrefixe(vTOB2.NomTable) ;
 lStCodeL    := GetSetCodeLettre(vTOB1.GetValue(lP1 + '_GENERAL') ,vTOB1.GetValue(lP1 + '_AUXILIAIRE')) ;
 lRdCouvCur  := Arrondi(Abs(vTOB1.GetValue(lP1 + '_DEBIT')-vTOB1.GetValue(lP1 + '_CREDIT')),V_PGI.OkDecV) ;
 lRdCouvDev  := Arrondi(Abs(vTOB1.GetValue(lP1 + '_DEBITDEV')-vTOB1.GetValue(lP1 + '_CREDITDEV')),V_PGI.OkDecV) ;

 if ( vTOB1.GetValue(lP1 + '_DATECOMPTABLE') > vTOB2.GetValue(lP2 + '_DATECOMPTABLE') ) then
  begin
   lDtDateMax := vTOB1.GetValue(lP1 + '_DATECOMPTABLE') ;
   lDtDateMin := vTOB2.GetValue(lP2 +'_DATECOMPTABLE') ;
  end
   else
    begin
     lDtDateMax := vTOB2.GetValue(lP2 + '_DATECOMPTABLE') ;
     lDtDateMin := vTOB1.GetValue(lP1 + '_DATECOMPTABLE') ;
    end ;

 vTOB1.PutValue(lP1 + '_ETATLETTRAGE'     , 'TL' ) ;
 vTOB2.PutValue(lP2 + '_ETATLETTRAGE'     , 'TL' ) ;
 vTOB1.PutValue(lP1 + '_LETTRAGE'         , lStCodeL ) ;
 vTOB2.PutValue(lP2 + '_LETTRAGE'         , lStCodeL ) ;
 vTOB1.PutValue(lP1 + '_COUVERTURE'       , lRdCouvCur ) ;
 vTOB1.PutValue(lP1 + '_COUVERTUREDEV'    , lRdCouvDev ) ;
 vTOB2.PutValue(lP2 + '_COUVERTURE'       , lRdCouvCur ) ;
 vTOB2.PutValue(lP2 + '_COUVERTUREDEV'    , lRdCouvDev ) ;
 vTOB1.PutValue(lP1 + '_DATEPAQUETMAX'    , lDtDateMax) ;
 vTOB1.PutValue(lP1 + '_DATEPAQUETMIN'    , lDtDateMin) ;
 vTOB2.PutValue(lP2 + '_DATEPAQUETMAX'    , lDtDateMax) ;
 vTOB2.PutValue(lP2 + '_DATEPAQUETMIN'    , lDtDateMin) ;

end ;


procedure TZEtebac.CreationLigne( vQ : TQuery ; vTOBLettrage : TOB ) ;
var
 lBoResult : boolean ;
begin

  // ligne sur le compte d'attente
  AddLigneImput ;
  InNUMLIGNE_I     := 1 ;
  StALettrer_I     := 'X' ;
  StGENERAL_I      := vQ.FindField('E_GENERAL').asString ;
  StAUXILIAIRE_I   := vQ.FindField('E_AUXILIAIRE').asString ;
  CSetMontants( FTOBLigneImputation ,
                RdDEBIT_R ,
                RdCREDIT_R ,
                Info.Devise.Dev,
                true );

  if ( Info.Journal.GetValue('J_TRESOCHAINAGE') = 'X' ) then
   CLettrageDeuxLigne( _ChargeEcriture(vQ,vTOBLettrage) , FTOBLigneImputation ) ;

  // ligne sur le compte de bqe
  AddLigneImput ;
  InNUMLIGNE_I     := 2 ;
  StGENERAL_I      := StGENERALBQE_R ;
  StAUXILIAIRE_I   := '' ;
  CSetMontants( FTOBLigneImputation ,
                RdCREDIT_R ,
                RdDEBIT_R ,
                Info.Devise.Dev,
                true );

  CalculSoldeImput;

  lBoResult := IsValideImputation ;
  ReporteEtat(lBoResult) ;

  if lBoResult then
   StETAT_R := cImputCorrectNonValide
    else
     StETAT_R := cImputIncorrecte ;

end ;



function TZEtebac.CreationDepuisEcriture( vTOBLettrage : TOB ) : boolean ;
var
 lQ   : TQuery ;
begin

 result := false ;

 try

  if RdDEBIT_R > 0 then
   begin
    lQ := _GetCredit(RdDEBIT_R,DtDATECOMPTABLE_R) ;

    if lQ.RecordCount = 1 then
     begin
      CreationLigne(lQ,vTOBLettrage) ;
      result := true ;
      Ferme(lQ) ;
      exit ;
     end ;

    Ferme(lQ) ;
   end ;

  if RdCREDIT_R > 0 then
   begin
    lQ := _GetDebit(RdCREDIT_R,DtDATECOMPTABLE_R) ;

    if lQ.RecordCount = 1 then
     begin
      CreationLigne(lQ,vTOBLettrage) ;
      result := true ;
     end ;

    Ferme(lQ) ;
   end ;

 except
  On E : Exception do
   PGIError('Erreur lors de la recherche dans la table écriture !' + #10#13 + E.Message ) ;
 end ;

end ;

end.
