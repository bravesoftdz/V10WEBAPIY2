unit ZReleveBanque;

interface

Uses StdCtrls, Controls, Classes,forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox,
     // VCL
     Dialogs,
     Windows,
     // Lib
     TZ,       // pour le TZF
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF SANSCOMPTA}
     HCompte,
     SaisAnal,
     db,
     dbTables,
     {$ENDIF}
     {$ENDIF}
     ZCompte,
     Ent1,
     ParamSoc, // pour le GetParamSocSecur
     ZTiers,
     ZTva,
     ZJournal, // pout le TZListJournal
     Lookup,
     SaisUtil, // ObjAna
     // Lib
     ULibEcriture,
     UCstEcriture,
     // AGL
     UTOB ;

const

 cPasImput                = '0';
 cImputIncorrecte         = '1';
 cImputCorrectNonValide   = '2';
 cImputValide             = '3';


type


 TZReleveBanque = Class
  private
   FTOBReleve                : TOB;
   FTOBLigneReleve           : TOB;
   FTOBImput                 : TOB;
   FTOBLigneImputation       : TOB;
   FTOBAnalytiq              : TOB;
   FTOBLigneAnalytiq         : TOB;
   FZTva                     : TZTva;
   FZComptes                 : TZCompte ;  // objet de gestion des comptes
   ZTiers                    : TZTiers;    // objet de gestion des tiers
   FZListJournal             : TZListJournal; // objet de gestion des journaux
   //FTOBTauxTVA               : TOB;       // tob des taux de tva par regime fiscal
   FAowner                   : TObject;
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
   FDevise                   : RDevise;
   FStChampsPieceBqe         : string;
   FStChampsPiece            : string;
   FStChampsReference        : string;
   FListObjAnalytiq          : TStringList;
   SQL_Insert                : TQuery;
   FCodeGuide                : string; // code du guide courant
   FStCurrentGuide           : string;
   FTypeContexte             : TTypeContexte;
   FStModeSaisie             : string;
   FStCompteurNormal         : string;

   procedure TransactionDeleteReleve;
   procedure TransactionSaveLigneReleve;

   function  GetDeviseCompte( value : string) : string;
   {$IFNDEF EAGLCLIENT}
   {$IFNDEF SANSCOMPTA}
   function  AnalytiqAuto ( vGuide : string ) : boolean;
   function  AnalytiqManuel ( vGuide : string ) : boolean;
   {$ENDIF}
   {$ENDIF}
   procedure RechParamEcriture;
   procedure SetJournalContrepartie ( Value : string );
   procedure NotifyError( vMessage, vDelphiMessage, vFunctionName : string);


  protected
   function  GetFieldPropertyReleve ( vIndex : integer ) : variant;
   procedure SetFieldPropertyReleve ( vIndex : integer ; Value : variant );
   function  GetFieldPropertyImput  ( vIndex : integer ) : variant;
   procedure SetFieldPropertyImput  ( vIndex : integer ; Value : variant );
   procedure SetFieldAnneeMoisImput ( vIndex : integer ; Value : variant );
   function  GetCodeTVA : string; // recupere le code TVA du compte d'imputation ou le codeTVA du dossier
   function  GetCodeTVAPourUnCompte ( Value : string ) : string;
   function  GetLibelleEtat : string; // libelle de l'etat d'imputation tablette CPETATIMPUTATION
   procedure CalculeSoldeReleve;
   procedure SetGeneralImput( Value : string);
   procedure SetCompteContrepartie( Value : string);
  public

   constructor Create(Aowner: TObject);
   destructor  Destroy; override;

   procedure InitializeLigneReleve;
   procedure InitializeLigneImput;
   procedure Initialize;
   procedure InitializeObject;

   function  IsValideGeneral ( vStCompte : string ) : boolean;

   function  IsValideLigneReleve : boolean;
   function  IsValideReleve      : boolean;
   function  IsValideLigneImput  : boolean;
   function  IsValideImputation  : boolean;

   function  SupprimerLigneImputVide : boolean;

   function  AddLigneReleve : TOB;
   function  AddLigneImput ( vInIndex : integer = -1 ) : TOB;
   {$IFNDEF EAGLCLIENT}
   {$IFNDEF SANSCOMPTA}
   function  AddAnalytique ( vGuide : string ) : TObjAna;
   {$ENDif}
   {$ENDIF}
   function  SaveLigneReleve : boolean;
   function  SaveReleve : boolean;
   function  DeleteLigneReleve : boolean;
   function  DeleteCurrentImputation : boolean;
   function  DeleteCurrentLigneImputation: boolean;
   function  LoadReleve ( Q : TQuery ) : boolean;
   function  LoadImputation : boolean;
   function  AppelAnalytique( vMode : TModeAna ) : boolean;
   function  GetTauxTVA ( Value : string ; vRdDebit : double ) : double;

   function  ExisteImputation( var Value : string ) : boolean;
   function  ExisteTiers( var vStAuxiliaire,vStGeneral : string ) : boolean;
   function  LibelleImputation ( Value : string ) : string;
   function  IsCollectif( Value : string ) : boolean; overload;
   function  IsCollectif : boolean; overload;
   function  IsTvaAutorise ( value : string ; vInIndex : integer = 0 ) : boolean;
   function  IsVentilable ( value : string ) : boolean; overload;
   function  IsVentilable  : boolean; overload;
   procedure RechercheAuxiliaire( G : THGrid );
   procedure AssignInfoTVA;
   procedure AssignInfoCompte;
   function  RechercheImputation : boolean;
   procedure NextImputation;
   procedure FirstImputation;
   procedure CalculSoldeImput;
   procedure RecopierLigneReleve( Value : TOB );
   procedure RecupereGuide( vTOBGuide : TOB );
   procedure CreerImputation(vTOBEcr : TOB);
   function  ProchainCompteur : integer; // retourne le prochain RLB_COMPTEUR disponible

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
   property Devise                : RDevise          read FDevise;

   property Statut                : TActionFiche     read FStatut                        write FStatut;
   property StatutImput           : TActionFiche     read FStatutImput                   write FStatutImput;
   property TOBLigneReleve        : TOB              read FTOBLigneReleve                write FTOBLigneReleve;
   property TOBReleve             : TOB              read FTOBReleve;
   property TOBLigneImputation    : TOB              read FTOBLigneImputation            write FTOBLigneImputation;
   property ZComptes              : TZCompte         read FZComptes;
   property TOBImput              : TOB              read FTOBImput;
   property TypeContexte          : TTypeContexte    read FTypeContexte                  write FTypeContexte;
   property StCurrentGuide        : string           read FStCurrentGuide;
   property StModeSaisie          : string           read FStModeSaisie;
   property StCompteurNormal      : string           read FStCompteurNormal;

  end; // class


implementation

const
 cStJournalInexistant      = 'Ce journal n''existe pas !';
 cStJournalPasMultiDevise  = 'Ce journal n''est pas multidevise !';
 cStCompteContrepartie     = 'le compte doit être égal au compte de contrepartie du compte !';
 cStCompteInterdit         = 'Ce compte est interdit pour ce journal!';
 cStPasSurCompteEcart      = 'Vous ne pouvez pas saisir sur le compte d''écart de conversion euro';
 cStPasSurCompteBilan      = 'Vous ne pouvez pas saisir sur le compte d''ouverture de bilan';
 cStManqueEtablissement    = 'Il manque le code établissement';
 cStNumeroPieceIncorrecte  = 'Le numéro de folio est incorrecte';
 cStCompteInterditPourNat  = 'Ce compte général est interdit pour cette nature de pièce';
 cStPasCompteRegul         = 'Ce compte n''est pas un compte de régularisation';
 cStPasSurCompteAux        = 'Vous ne pouvez pas saisir sur un compte collectif';



constructor TZReleveBanque.Create(Aowner: TObject);
begin

 inherited Create;

 FAowner        := AOwner;

 FZComptes      := TZCompte.Create(false);       // Objet de gestion des GENERAUX
 ZTiers         := TZTiers.Create;               // Objet de gestion des tiers
 FTOBReleve     := TOB.Create('',nil,-1);        // TOB des lignes de relevés bancaires
 FZTva          := TZTva.Create;
 FZListJournal  := TZListJournal.Create; // Objet de gestion des journaux

end;

destructor TZReleveBanque.Destroy;
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
// if assigned(FTOBTauxTVA)         then FTOBTauxTVA.Free;
 if assigned(ZTiers)              then ZTiers.Free;
 if assigned(SQL_Insert)          then SQL_Insert.Free;
 if assigned(ZComptes)            then ZComptes.Free;
 if assigned(FZTva)               then FZTva.Free;
 if assigned(FZListJournal)       then FZListJournal.Free;

 FZComptes                        := nil;
 SQL_Insert                       := nil;
 FTOBReleve                       := nil;
// FTOBTauxTVA                      := nil;
 ZTiers                           := nil;
 FListObjAnalytiq                 := nil;
 FZTva                            := nil;
 FZListJournal                    := nil;

 inherited destroy;

 except
  on E : Exception do
   begin
    NotifyError( 'Erreur lors de la destruction !',
                  E.Message ,
                 'TZReleveBanque.Destroy');
   end;
 end;

end;

function TZReleveBanque.GetFieldPropertyReleve ( vIndex : integer ) : variant;
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
 end; // case

end;

function TZReleveBanque.GetFieldPropertyImput ( vIndex : integer ) : variant;
begin

 result := '';

 if not assigned(FTOBLigneImputation) then exit;

 case vIndex of
  0  : result := FTOBLigneImputation.GetValue('CRL_COMPTEUR');
  1  : result := FTOBLigneImputation.GetValue('CRL_JOURNAL');
  2  : result := FTOBLigneImputation.GetValue('CRL_DATECOMPTABLE');
  3  : result := FTOBLigneImputation.GetValue('CRL_REFINTERNE');
  4  : result := FTOBLigneImputation.GetValue('CRL_REFEXTERNE');
  5  : result := FTOBLigneImputation.GetValue('CRL_ETAT');
  6  : result := FTOBLigneImputation.GetValue('CRL_GENERAL');
  7  : result := FTOBLigneImputation.GetValue('CRL_TAUXTVA');
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

 end; // case

end;


procedure TZReleveBanque.SetFieldPropertyReleve ( vIndex : integer ; Value : variant );
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
  18 : if Value <> '' then
        FTOBLigneReleve.PutValue('CRL_AUXILIAIRE', Value);
  19 : FTOBLigneReleve.PutValue('CRL_GENERALBQE', Value);
  20 : FTOBLigneReleve.PutValue('CRL_TVASAISIE', Value);
  21 : if IsValidDate(Value) then
        FTOBLigneReleve.PutValue('CRL_DATEVALEUR' , StrToDate(Value));
  22 : FTOBLigneReleve.PutValue('CRL_CODEAFB', Value);
  23 : FTOBLigneReleve.PutValue('CRL_REFPIECE', Value);
  24 : FTOBLigneReleve.PutValue('CRL_COTATION', Value);
 end; // case

 if FStatut <> taCreat then
  FStatut := taModif;

end;


procedure TZReleveBanque.SetFieldPropertyImput ( vIndex : integer ; Value : variant );
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
 19 : if Value <> '' then
       FTOBLigneImputation.PutValue('CRL_AUXILIAIRE', Value);
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

 end; // case

 if FStatutImput <> taCreat then
  FStatutImput := taModif;

end;


procedure TZReleveBanque.InitializeObject;
begin
 FZTva.InitializeObject;
 FZTva.ZCompte := ZComptes;
end;



procedure TZReleveBanque.InitializeLigneReleve;
var
 lDtDateFinN1 : TDateTime;
begin

 InCOMPTEUR_R             := 0;
// RdTVA_R                  := Unassigned;
 StIMPORT_R               := '-';
 StETAT_R                 := cPasImput; // ligne non importé par etabac par defaut
 InTYPE_R                 := integer(TTypeReleve); // releve bancaire par défaut
 StTVASAISIE_R            := '-';
 StDEVISE_R               := FDevise.Code;

  // initalisation de la date pour la grille de saisie
 if VH^.Suivant.Fin <> 0 then
   lDtDateFinN1 := VH^.Suivant.Fin
    else
     lDtDateFinN1 :=  VH^.EnCours.Fin;

 if Date > lDtDateFinN1 then
  DtDATECOMPTABLE_R := lDtDateFinN1
   else
    DtDATECOMPTABLE_R   := Date;

 DtDATECREATION_R         := Date;

 StGENERALBQE_R           := StCompteContrepartie; // compte de contrepartie du journal de banque

end;


procedure TZReleveBanque.Initialize;
begin

 FRdSoldeDebit  := 0;
 FRdSoldeCredit := 0;

 if FTOBReleve.Detail.Count > 0 then
  FTOBReleve.ClearDetail;

 FTOBImput := nil;

end;



function TZReleveBanque.ProchainCompteur : integer;
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
              '"1899-12-30 00:00:00.000",'           +
              '"2099-12-31 00:00:00.000", '          +
              '"-", '                                +
              '"-", '                                +
              'NULL, '                               +
              '"-", '                                +
              '1, '                                  +
              '1, '                                  +
              '"-", '                                +
              '"-" ) ');
   result := 1;
  end; // if

end;


function TZReleveBanque.AddLigneReleve : TOB;
begin

 FTOBLigneReleve                 := TOB.Create('CRELBQE',FTOBReleve,-1);
 FTOBImput                       := nil;
 InitializeLigneReleve;
 Statut                          := taCreat;
 StatutImput                     := taCreat;
 result                          := FTOBLigneReleve;

end;

procedure TZReleveBanque.TransactionSaveLigneReleve;
begin

  // on supprime toutes les lignes de releve
  ExecuteSQL('delete CRELBQE where CRL_COMPTEUR = ' + varToStr(InCOMPTEUR_R) );
  // on supprime toutes les ligesn d'analytiques
  ExecuteSQL('delete CRELBQEANALYTIQ where CRY_NUMEROPIECE = ' + varToStr(InCOMPTEUR_R) );

  // on sauve l'entete puis les lignes
  FTOBLigneReleve.SetAllModifie(true);
  FTOBLigneReleve.InsertDBByNivel(false);

end;


function TZReleveBanque.SaveLigneReleve : boolean;
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
               'TZReleveBanque.TransactionSaveLigneReleve' );

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

function TZReleveBanque.SupprimerLigneImputVide : boolean;
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

function TZReleveBanque.SaveReleve : boolean;
{$IFDEF TESTPERF}
var
 c1,c2 : integer;
{$ENDIF}
begin

 result := false;

 {$IFDEF TESTPERF}
  c1 := gettickcount;
 {$ENDIF}

 FTOBReleve.InsertDBByNivel( false );
 // result := DBInsert;

 {$IFDEF TESTTESTPERF}
  c2 := gettickcount;
  ShowMessage( inttostr(c2-c1) );
 {$ENDIF}

 result := true;

end;

procedure TZReleveBanque.TransactionDeleteReleve;
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
function TZReleveBanque.DeleteLigneReleve : boolean;
begin

  result := true;

  try

  if Blocage([cStVerrouTreso], false , cStVerrouTreso ) then
   exit;

 // suppression des enregistrement en base
 if Statut <> taCreat then
  result := Transactions(TransactionDeleteReleve,1) = oeOK  ;

 if not result then exit;

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
  Bloqueur( cStVerrouTreso , false );
 end; // try


end;

function TZReleveBanque.IsValideReleve : boolean;
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

function  TZReleveBanque.IsValideGeneral ( vStCompte : string ) : boolean;
var
 lInIndex  : integer;
begin

 result    := false;

 lInIndex  := FZComptes.GetCompte(vStCompte);

 // le compte doit etre defini est existe
 if (vStCompte = '') or (lInIndex < 0 )  then
  begin
   result := false;
   exit;
  end; // if

  // il ne doit pas etre egale au compte d'ecart de conversion euro
  if (vStCompte=VH^.EccEuroDebit) or (vStCompte=VH^.EccEuroCredit) then
    begin
     NotifyError( cStPasSurCompteEcart ,'','TZReleveBanque.IsValideGeneral');
     result := false;
     exit;
    end ;

// il ne doit pas etre interdit pour ce journal
 if EstInterdit(FZListJournal.GetValue('J_COMPTEINTERDIT'), vStCompte, 0) > 0 then
  begin
   NotifyError( cStCompteInterdit ,'','TZReleveBanque.IsValideGeneral');
   result := false;
   exit;
  end; // if

 // il ne doit pas etre un compet de bilan
 if (vStCompte=VH^.OuvreBil) then
  begin
    NotifyError( cStPasSurCompteBilan ,'','TZReleveBanque.IsValideGeneral');
    result := false;
    exit;
  end; // if

  // il ne doit pas etre un compet de bilan
 if FZComptes.IsCollectif(lInIndex) then
  begin
    NotifyError( cStPasSurCompteAux ,'','TZReleveBanque.IsValideGeneral');
    result := false;
    exit;
  end; // if

 result := true;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... : 04/07/2001
Description .. : Validation d'un ligne dans la TOB
Mots clefs ... :
*****************************************************************}
function TZReleveBanque.IsValideLigneReleve : boolean;
var
 lStGeneral : string;
begin

  result       := true;  { TODO : Gestion des montants max et min sur débit/crédit }

  lStGeneral   := StGENERAL_R;
  if lStGeneral <> '' then
   result := IsValideGeneral(lStGeneral);
  
  if not result then exit;

  // la date doit être valide
  if not IsValidDate(DtDATECOMPTABLE_R) then
   begin
    result := false;
    NotifyError(cStTexteDateNonValide + VarToStr(DtDATECOMPTABLE_R), '' , 'TZReleveBanque.IsValideLigneReleve');
    exit;
   end;

  if StJOURNAL_R = '' then
   StJOURNAL_R := StJournalContrepartie;

  result := StJOURNAL_R <> '';

  if not result then exit;

  if StETAT_R = '' then
   StETAT_R := cPasImput;

  // soit le credit ou le debit sont renseigné et supérieur à zéro
  result := (( Round(RdDEBIT_R) > 0 ) xor ( Round(RdCREDIT_R) > 0 )) ;

  if not result then
   begin
    NotifyError( cStTexteDebitEtCreditNull , '' , 'TZReleveBanque.IsValideLigneReleve');
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
Description .. : Test l'existence d'uner imputation
Mots clefs ... :
*****************************************************************}
function TZReleveBanque.ExisteImputation ( var Value : string ) : boolean;
var
 lIndex : integer;
begin
 result := false;
 lIndex := ZComptes.GetCompte(Value);
 result := ( lIndex > - 1 );
// if ( lIndex > - 1 ) then
//  result := Value; // on recupere le compte complete par des zero
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... :   /  /
Description .. : Retourne le libellé du compte general
Mots clefs ... :
*****************************************************************}
function TZReleveBanque.LibelleImputation ( Value : string ) : string;
var
 lIndex : integer;
begin
 result := '';
 lIndex := ZComptes.GetCompte(Value);
 if ( lIndex > 0 ) then
  result := ZComptes.GetValue('G_LIBELLE', lIndex) ;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/07/2001
Modifié le ... : 17/07/2001
Description .. : Chargement des lignes de releves a partir de Q ( cree dans
Suite ........ : un filtre )
Mots clefs ... :
*****************************************************************}
function TZReleveBanque.LoadReleve ( Q : TQuery ) : boolean;
begin

 Initialize;

 result := FTOBReleve.LoadDetailDB('CRELBQE','','',Q,true);
 if FTOBReleve.Detail.Count > 0 then
  begin
   FTOBLigneReleve     := FTOBReleve.Detail[0];
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
function TZReleveBanque.GetCodeTVA: string;
var
 lIndex     : integer;
 lStGeneral : string;
begin
 result := '';
 // on recherche le code tva du compte
 lStGeneral := VarToStr(StGENERAL_R);
 lIndex := ZComptes.GetCompte(lStGeneral);
 if ( lIndex > -1 ) then
  result := ZComptes.GetValue('G_TVA', lIndex) ;

 // s'il n'existe pas on prend le param par defaut
 if result = '' then
  result := GetParamSocSecur('SO_CODETVAGENEDEFAULT','');

end;

function TZReleveBanque.GetCodeTVAPourUnCompte ( Value : string ) : string;
var
 lIndex     : integer;
begin
 result := '';
 lIndex := ZComptes.GetCompte(Value);
 if ( lIndex > 0 ) then
  result := ZComptes.GetValue('G_TVA', lIndex) ;

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
procedure TZReleveBanque.AssignInfoTVA;
begin

 FZTva.Load(StGENERAL_R,RdCREDIT_R = 0);
 if FZTva.StCompteTVA = '' then
  StTVASAISIE_R := '-'
   else
    begin
     if not ( StTVASAISIE_R = 'X' ) then
      RdTAUXTVA_R   := FZTva.RdTaux;
     FStCompteTVA  := FZTva.StCompteTVA;
     StTVASAISIE_R := 'X';
    end; // if
end;

function TZReleveBanque.GetTauxTVA ( Value : string ; vRdDebit : double ) : double ;
begin

 FZTva.Load(StGENERAL_R,RdDEBIT_R > 0);
 result := FZTva.RdTaux;
 
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/07/2001
Modifié le ... :   /  /
Description .. : recherche le libelle en clair de l'etat de l'imputation
Mots clefs ... :
*****************************************************************}
function TZReleveBanque.GetLibelleEtat : string;
begin
 result := RechDom('CPETATIMPUTATION',StETAT_R,false);
end;

procedure TZReleveBanque.AssignInfoCompte;
begin
 if StLIBELLE_R = '' then
  StLIBELLE_R := LibelleImputation (StGENERAL_R);
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
function TZReleveBanque.LoadImputation : boolean;
var
 i                : integer;
 lQImput          : TQuery;
 lQAna            : TQuery;
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
   FTOBLigneImputation.AddChampSup('CRL_VENTILABLE',true);
   FTOBLigneImputation.AddChampSup('CRL_COLLECTIF',true);
 //  TOBLigneImputation.AddChampSup('CRL_ANA',true);
   StatutImput          := taConsult;
  end; // if

 // on recherche les lignes d'analytiques
 for i := 0 to ( FTOBImput.Detail.Count - 1 ) do
  begin

   // on recherche si le compte est ventilable
   FTOBLigneImputation := FTOBImput.Detail[i];
   SetGeneralImput( StGENERAL_I);

   // si le compte est ventilable et que la ligne est flaguer comme possedant de l'analytique
   if IsVentilable and ( StANA_I = 'X' ) then
    begin

     lQAna := OpenSql ( 'select * from CRELBQEANALYTIQ ' +
                        'where CRY_GENERAL = "'          + StGENERAL_I + '" '     +
                        'and CRY_NUMEROPIECE = '         + VarToStr(InCOMPTEUR_R) +
                        'and CRY_NUMLIGNE = '            + VarToStr( i + 1 ) ,
                        true );

     if not lQAna.EOF then
      begin
       // il existe des ventilations analytiques
       FTOBAnalytiq := TOB.Create('', FTOBLigneImputation , -1);
       FTOBAnalytiq.LoadDetailDB('CRELBQEANALYTIQ','','',lQAna,true);
       Ferme( lQAna );
      end; // if

    end; // if
  end; // for

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
function TZReleveBanque.RechercheImputation : boolean;
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



function TZReleveBanque.AddLigneImput ( vInIndex : integer = -1 ) : TOB;
begin

 FTOBLigneImputation              := TOB.Create('CRELBQE',FTOBImput,vInIndex);
 FTOBLigneImputation.AddChampSup('CRL_VENTILABLE',true);
 FTOBLigneImputation.AddChampSup('CRL_COLLECTIF',true);
 FTOBLigneImputation.AddChampSup('GUIDE',true);
 FTOBLigneImputation.AddChampSup('CRL_NUMGROUPEECR',true);
 InitializeLigneImput;
 StatutImput                     := taCreat;
 result                          := FTOBLigneImputation;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/07/2001
Modifié le ... :   /  /
Description .. : Initialise les valeurs par défaut d'une ligne de releve
Mots clefs ... :
*****************************************************************}
procedure TZReleveBanque.InitializeLigneImput;
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
end;


procedure TZReleveBanque.RechercheAuxiliaire( G : THGrid );
var
 lIndex       : integer;
 lStCritere   : string;
 lStColonnes  : string;
 lStOrderBy   : string;
 lStGeneral   : string;
begin

 lStGeneral := VarToStr(StGENERAL_I);
 lIndex := ZComptes.GetCompte(lStGeneral);
 lStCritere := '';

 if ( lIndex > 0 ) then
  begin

   if ZComptes.GetValue( 'G_NATUREGENE', lIndex ) = 'COF' then
    lStCritere := 'T_NATUREAUXI = "FOU"'
     else
      if ZComptes.GetValue( 'G_NATUREGENE', lIndex ) = 'COC' then
       lStCritere := 'T_NATUREAUXI = "CLI"'
        else
         if ZComptes.GetValue( 'G_NATUREGENE', lIndex )= 'COS' then
          lStCritere := 'T_NATUREAUXI = "SAL"' ;

  end; // if

    if lStCritere = '' then
     begin
      lStColonnes := 'T_LIBELLE, T_NATUREAUXI';
      lStOrderby  := 'T_AUXILIAIRE, T_NATUREAUXI';
     end
      else
       begin
        lStColonnes := 'T_LIBELLE';
        lStOrderBy  := 'T_AUXILIAIRE';
       end;

   LookupList(G,TraduireMemoire('Auxiliaire'),'TIERS','T_AUXILIAIRE',lStColonnes,lStCritere,lStOrderBy,TRUE, 2) ;

end;

function TZReleveBanque.IsCollectif( Value : string ) : boolean;
begin
 result := ZComptes.IsCollectif(Value);
end;

function TZReleveBanque.IsTvaAutorise ( value : string ; vInIndex : integer = 0 ) : boolean;
begin
 result := ZComptes.IsTvaAutorise(Value,vININdex);
end;

function TZReleveBanque.IsCollectif : boolean;
var
 lIndex : integer;
 lStGENERAL : string;
begin

 result := false;
 lStGeneral := VarToStr(StGENERAL_I);
 lIndex := ZComptes.GetCompte(lStGeneral);

 if ( lIndex > -1 ) then
  result := ZComptes.IsCollectif(lIndex);

end;

function TZReleveBanque.ExisteTiers( var vStAuxiliaire,vStGeneral : string ) : boolean;
begin
 result := ZComptes.ExisteTiers(vStAuxiliaire,vStGeneral);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/08/2001
Modifié le ... :   /  /
Description .. : On se place sur la premiere ligne d'imputation du releve
Mots clefs ... :
*****************************************************************}
procedure TZReleveBanque.FirstImputation;
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
procedure TZReleveBanque.NextImputation;
begin

 // on passe à la ligne suivante
 FCurrentLigneImput := FCurrentLigneImput + 1;
 if FCurrentLigneImput < FTOBImput.Detail.Count then
  FTOBLigneImputation     := FTOBImput.Detail[FCurrentLigneImput]
   else
    FTOBLigneImputation   := nil;

end;

procedure TZReleveBanque.CalculSoldeImput;
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
   NotifyError( cStTexteErreurSoldeImput, E.Message , 'TZReleveBanque.CalculSoldeImput' );
 end; // try

end;

function TZReleveBanque.IsValideImputation : boolean;
var
 i : integer;
begin

 if not assigned(FTOBImput) then
  begin
   NotifyError( cStTexteErreurImput , '' ,'TZReleveBanque.IsValideImputation');
   result := false;
   exit;
  end;

 result := Arrondi(FRdSoldeImput,4) = 0;

 if not result then
  begin
   NotifyError( cStTexteSoldeNonNull + StrfMontant(FRdSoldeImput, 15 , V_PGI.OkDecV, '' , true) , '' ,'TZReleveBanque.IsValideImputation');
   exit;
  end; // if    

 // validation de chaque ligne
 for i := 0 to FTOBImput.Detail.Count - 1 do
  begin
   // faire une fct NextTOB(index);
   FTOBLigneImputation := FTOBImput.Detail[i];
   result             := IsValideLigneImput;
   if not result then
    break;

  end // for

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 18/07/2001
Modifié le ... :   /  /
Description .. : Suppression des ligne d'imputations de la ligne de releves
Suite ........ : courantes
Mots clefs ... :
*****************************************************************}
function TZReleveBanque.DeleteCurrentImputation : boolean;
begin

 result              := false;
 // on se place sur la premiere ligne
 FCurrentLigneImput  := 0;

 if assigned(FTOBLigneReleve.Detail) then
  FTOBLigneReleve.Detail.Clear;

 FTOBImput           := nil;
 FStatutImput        := taConsult;
 result              := true;

end;


procedure TZReleveBanque.CalculeSoldeReleve;
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


function TZReleveBanque.IsValideLigneImput : boolean;
var
 lStGeneral : string;
begin

 result     := true;  { TODO : Gestion des montants max et min sur débit/crédit }

 lStGeneral := StGENERAL_I;

 result := isValideGeneral(lStGeneral);

 if not result then
  begin
   NotifyError( cStTexteGeneralInexistant + StGENERAL_I , '' , 'TZReleveBanque.IsValideLigneImput' );
   exit;
  end; // if

 if ( RdDEBIT_I < 0 ) and ( not VH^.MontantNegatif ) then
  begin
   result := false;
   NotifyError( cStTexteMontantNegatif + #13#10 + ' Valeur : ' + varToStr(RdDEBIT_I), '' , 'TZReleveBanque.IsValideLigneImput' );
   exit;
  end;

 if ( RdCREDIT_I < 0 ) and ( not VH^.MontantNegatif ) then
  begin
   result := false;
   NotifyError( cStTexteMontantNegatif + #13#10 + ' Valeur : ' + varToStr(RdCREDIT_I), '' , 'TZReleveBanque.IsValideLigneImput' );
   exit;
  end;

 // recupération du libellé de l'imputation s'il n'est pas définie
 if ( StGENERAL_I <> '' ) and ( StLIBELLE_I = '' ) then
  StLIBELLE_I := LibelleImputation(StGENERAL_I);

 if InCOMPTEUR_I = 0 then
  begin
   result := false;
   NotifyError( 'Erreur, la valeur du champ CRL_COMPTEUR est null', '' , 'TZReleveBanque.IsValideLigneImput' );
   exit;
  end; // if

end;


function TZReleveBanque.DeleteCurrentLigneImputation : boolean;
begin

 result := false;

 FTOBLigneImputation.free;
 FTOBLigneImputation := nil;
 FStatutImput        := taModif;
 result              := true;

end;


procedure TZReleveBanque.RechParamEcriture;
var
 Q : TQuery;
begin

 Q := nil;

 try

  Q := OpenSQL ('select J_CHOIXDATE  '        +
                //'J_CHOIXPIECEBQ, '             +
                //'J_CHOIXPIECE, '               +
                //'J_CHOIXREFERENCE '            +
                'from JOURNAL '                +
                'where J_JOURNAL="'            + StJournalContrepartie + '"' , true);

  FStChampsPieceBqe   := ''; //Q.FindField('J_CHOIXPIECEBQ').asString;
  FStChampsPiece      := ''; //Q.FindField('J_CHOIXPIECE').asString;
  FStChampsReference  := ''; //Q.FindField('J_CHOIXREFERENCE').asString;

 finally
  if assigned(Q) then Ferme(Q);
 end; // try

end;


{$IFNDEF EAGLCLIENT}
{$IFNDEF SANSCOMPTA}
function TZReleveBanque.AddAnalytique ( vGuide : string ) : TObjAna;
begin

 Result                    := TObjAna.Create ;
 Result.General            := StGENERAL_R;
 Result.Exercice           := VH^.CPExoRef.Code;
 Result.DateComptable      := DtDATECOMPTABLE_R;
 Result.NumeroPiece        := InCOMPTEUR_R;
 Result.NumLigne           := InNUMLIGNE_I;
 Result.RefInterne         := StREFINTERNE_R;
 Result.Libelle            := StLIBELLE_R;
 Result.Journal            := FStJournalContrepartie;
 Result.NaturePiece        := 'OD';
 Result.QualifPiece        := 'N';  // ecriture normale
 Result.Etabl              := VH^.EtablisDefaut;
 Result.TotalEcriture      := RdDEBIT_R + RdCREDIT_R;
 Result.TotalDevise        := Result.TotalEcriture;
 Result.TotalEuro          := Result.TotalEcriture;
 Result.Devise             := FDevise.Code;
 Result.TauxDev            := FDevise.Taux;
 Result.DateTauxDev        := FDevise.DateTaux;
 Result.Societe            := V_PGI.CodeSociete ;
 Result.Sens               := 2 - Ord(RdDEBIT_R <> 0) ;
 Result.Decimale           := FDevise.Decimale;
 Result.Utilisateur        := V_PGI.User ;
 Result.ModeOppose         := false ;
 Result.GuideA             := vGuide;
 Result.QualifQte1         := '...';
 Result.QualifQte2         := '...';
 Result.TotalQte1          := 0;
 Result.TotalQte2          := 0;
// ObjAna.RefExterne         := Ecr.GetValue('E_REFEXTERNE') ;
// ObjAna.DateRefExterne     := Ecr.GetValue('E_DATEREFEXTERNE') ;
// ObjAna.Affaire            := Ecr.GetValue('E_AFFAIRE') ;


end;


function TZReleveBanque.AnalytiqManuel( vGuide : string ) : boolean;
var
 lObjAnalytique : TObjAna;
 lP             : P_TV;
 lZGeneral      : TGGeneral;
 lRecA          : REC_ANA;
 i              : integer;
 j              : integer;
 k              : integer;
 lIndex         : integer;
 lObjMemo       : TStringList;

 procedure TransfertLocalVersGlobal ( vChamps : string );
 var
  lStChampsAna : string;
 begin
  lStChampsAna := 'Y' + Copy( vChamps,4,length(vChamps) );

  try

  // FTOBLigneAnalytiq.PutValue( vChamps , P_TV(lObjAnalytique.AA[i+1].L[j]).F[TrouveIndice(VH^.DescriAna,lStChampsAna,true)].V );

  except
   on E : Exception do
    begin

     NotifyError( cStTexteErreurAna + vChamps ,
                  E.Message ,
                  'TZReleveBanque.AnalytiqManuel' );
     // pour stopper les traitements suivants
     raise EAbort.Create('Erreur a la creation');

    end; // on
  end; // try

 end;

 procedure CreationObjAnalytique;
 var
  li            : integer;
  lj            : integer;
  lk            : integer;
  lIndex        : integer;
  lStChampsAna  : string;
 begin

  lObjAnalytique   := AddAnalytique ( vGuide );

  lIndex := 0;

  // on recherche s'il existait deja des lignes
   if FTOBLigneImputation.Detail.Count <> 0 then
    begin
     for li := 0 to ( FTOBLigneImputation.Detail.Count - 1 ) do
      begin

       // on se place sur le premier axe
       FTOBAnalytiq := FTOBLigneImputation.Detail[li];
       // on ajout un Axe a l'Obj lObjAnalytique
       lObjAnalytique.AA[li+1] := TAA.Create ;

       // on parcours les lignes d'analytiques
       for lj := 0 to ( FTOBAnalytiq.Detail.Count - 1 ) do
        begin
         // Objet stockant les lignes d'analytiques ( definit dans SAISUTIL )
         lP := P_TV.Create ;
         // on recupere la ligne courante
         FTOBLigneAnalytiq := FTOBAnalytiq.Detail[lj];

         for lk := 0 to ( FTOBLigneAnalytiq.NbChamps - 1) do
          begin
           lStChampsAna   := 'Y' + Copy( FTOBLigneAnalytiq.GetNomChamp(lk+1),4,length(FTOBLigneAnalytiq.GetNomChamp(lk+1)) );
          // lIndex         := TrouveIndice(VH^.DescriAna,lStChampsAna,True);
           lP.F[lIndex].V := FTOBLigneAnalytiq.GetValeur(lk+1);
          end;

          // affectation des champs OLDDEBIT et OLDCREDIT
          lP.F[FTOBLigneAnalytiq.NbChamps].V      := 0;
          lP.F[FTOBLigneAnalytiq.NbChamps + 1].V  := 0;

          lObjAnalytique.AA[li+1].L.Add(lP) ;

          // le memo n'est pas gerer pour l'instant
          lObjMemo       :=TStringList.Create ;
          lObjAnalytique.AA[li+1].M.Add(lObjMemo) ;


         end; // for

        end; // for

    end; // if

 end; // procedure

begin

 lZGeneral                     := TGGeneral.Create(StGENERAL_I);
 // on initialise le pointeur
 FTOBAnalytiq                  := nil;
 lObjAnalytique                := nil;

 result                        := false;

 try
  try

   if not assigned( FListObjAnalytiq ) then
    begin
     // creation de la liste stockant les Objets lObjAnalytique
     FListObjAnalytiq := TStringList.Create;
     CreationObjAnalytique;
    end
     else
      begin
       // recuperation de l'objet precedement cree
       lIndex := FListObjAnalytiq.IndexOf(IntToStr(InCOMPTEUR_I + InNUMLIGNE_I));
       if lIndex > - 1 then
        lObjAnalytique := TObjAna( FListObjAnalytiq.Objects[lIndex] )
         else
          CreationObjAnalytique;
      end; // if

   // renseignement de la structure
   lRecA.QuelEcr                 := EcrGen; // defini dans SaisUtil
   lRecA.CC                      := lZGeneral;
   lRecA.MontantEcrP             := RdDEBIT_I + RdCREDIT_I;
   lRecA.MontantEcrD             := lRecA.MontantEcrP;//Ecr.GetValue('E_DEBITDEV')+Ecr.GetValue('E_CREDITDEV');
   lRecA.MontantEcrE             := lRecA.MontantEcrP;//Ecr.GetValue('E_DEBITEURO')+Ecr.GetValue('E_CREDITEURO');
   lRecA.RefEcr                  := StREFEXTERNE_I;
   lRecA.LibEcr                  := StLIBELLE_I;
   lRecA.OBA                     := lObjAnalytique;
   lRecA.Action                  := taCreat;
   lRecA.DEV                     := FDevise;

   //if Ecr.GetValue('E_SAISIEEURO')='X' then RecA.OBA.ModeOppose:=TRUE ;
   lRecA.NumGeneAxe              := 0;
   lRecA.VerifVentil             := true;

   // appel de la fenêtre d'analytique
   SaisieAnal(lRecA);

   // on test s'il existait des lignes
   if FTOBLigneImputation.Detail.Count = 0 then
    FTOBAnalytiq := TOB.Create('',FTOBLigneImputation,-1)
     else
      begin
       // recuperation des lignes precedentes
       FTOBAnalytiq := FTOBLigneImputation.Detail[0];
       // on supprime les lignes existantes
       FTOBAnalytiq.ClearDetail;
      end; // if

   // on flag la ligne
   StANA_I := 'X';

   // ajout des enregistrements
   for i := 0 to MAXAXE - 1 do
     begin
      if ( lObjAnalytique.AA[i+1] = nil ) or ( lObjAnalytique.AA[i+1].L.Count = 0 ) then continue ;
       for j := 0 to lObjAnalytique.AA[i+1].L.Count - 1 do // on parcours les lignes d'analytique
        begin

         FTOBLigneAnalytiq := TOB.Create('CRELBQEANALYTIQ',FTOBAnalytiq,-1);

         for k := 1 to ( FTOBLigneAnalytiq.NbChamps ) do
          TransfertLocalVersGlobal ( FTOBLigneAnalytiq.GetNomChamp(k) );

   //      FTOBLigneAnalytiq.PutValue( 'CRY_NUMEROPIECE'          , P_TV(lObjAnalytique.AA[i+1].L[j]).F[TrouveIndice(VH^.DescriAna,'Y_NUMEROPIECE'     ,true)].V) ;
  //       FTOBLigneAnalytiq.PutValue( 'CRY_REFINTERNE'           , P_TV(lObjAnalytique.AA[i+1].L[j]).F[TrouveIndice(VH^.DescriAna,'Y_REFINTERNE'      ,true)].V) ;
  //       FTOBLigneAnalytiq.PutValue( 'CRY_LIBELLE'              , P_TV(lObjAnalytique.AA[i+1].L[j]).F[TrouveIndice(VH^.DescriAna,'Y_LIBELLE'         ,true)].V) ;

        end; // for
      end; // for

   // on stoke l'Objet lObjAnalytique pour pouvoir le reutiliser s'il n'etait pas present
   if FListObjAnalytiq.IndexOf(IntToStr(InCOMPTEUR_I + InNUMLIGNE_I)) = -1 then
    FListObjAnalytiq.AddObject(IntToStr(InCOMPTEUR_I + InNUMLIGNE_I),TObject(lObjAnalytique) );

   result := true;

  finally
   if assigned( lZGeneral )          then lZGeneral.Free;
  end; // try

 except
  on E : Exception do
   begin
    lObjAnalytique.Free;
    NotifyError( cStTexteErreurAnalytiq,
                 E.Message ,
                'TZReleveBanque.AnalytiqManuel');
   end; // if

 end;


end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 12/09/2001
Modifié le ... :   /  /
Description .. : Affectation des lignes d'analytiques en fonction des guides
Suite ........ : analytique ou des vnetilations par défaut
Mots clefs ... :
*****************************************************************}
function TZReleveBanque.AnalytiqAuto ( vGuide : string ) : boolean;
var
 lTOBECRITURE : TZF;
 i            : integer;
 j            : integer;
 k            : integer;
 z            : integer;
 lTOBAxe      : TOB;
 lTOBAna      : TOB;

 procedure TransfertLocalVersGlobal ( vChamps : string );
 var
  lStChampsAna : string;
 begin

  lStChampsAna := 'Y' + Copy( vChamps,4,length(vChamps) );

  try

   FTOBLigneAnalytiq.PutValue( vChamps , lTOBAna.GetValue(lStChampsAna) );

  except
   on E : Exception do
    begin

     NotifyError( cStTexteErreurAna + vChamps ,
                  E.Message ,
                  'TZReleveBanque.AnalytiqAuto' );
     // pour stopper les traitements suivants
     raise EAbort.Create('Erreur a la creation');

    end; // on
  end; // try


 end;

 procedure TransfertGlobalVersLocal ( vChamps : string );
 var
  lStChampsAna : string;
 begin
  lStChampsAna := 'Y' + Copy( vChamps,4,length(vChamps) );

  try

   lTOBAna.PutValue( lStChampsAna , FTOBLigneAnalytiq.GetValue(vChamps) );

  except
   on E : Exception do
    begin

     NotifyError( cStTexteErreurAna + vChamps ,
                  E.Message ,
                  'TZReleveBanque.AnalytiqAuto');
     // pour stopper les traitements suivants
     raise EAbort.Create('Erreur a la creation');

    end; // on
  end; // try

 end;


begin

 // on creee une TZF ( une Objet TOB sur la table ecriture ) que l' on passe
 // a la fonction VentilierTOB.
 // La structure final sera :
 // lTOBEcriture
 //    -> TOBAxe ( autant de TOB que d'axe dans la compta
 //         -> TOBAna
 //         -> TOBAna
 //         -> ...

 // TOB basé sur la table ECRITURE utilisé par la fct VentilerTOB

 // le debit et le credit n'ont pas encore été définie ( mode guide ) , on sort directement
 if ( RdDEBIT_I = 0 ) and ( RdCREDIT_I = 0 ) then
  begin
   result := true;
   exit;
  end;

 result := false;

 lTOBECRITURE     := TZF.Create('ECRITURE', nil,-1);

 // il faut creer les TOB des axes avant des les passer à la fonction
 for z := 1 to MAXAXE do
  TZF.Create('A'+IntToStr(z), lTOBECRITURE , -1) ;

 try
  // renseignement des valeurs
  lTOBECRITURE.PutValue('E_EXERCICE',        VH^.EnCours.Code);
  lTOBECRITURE.PutValue('E_JOURNAL',         StJOURNAL_I);
  lTOBECRITURE.PutValue('E_DATECOMPTABLE',   DtDATECOMPTABLE_I);
  lTOBECRITURE.PutValue('E_NUMEROPIECE',     InCOMPTEUR_I);
  lTOBECRITURE.PutValue('E_NUMLIGNE',        InNUMLIGNE_I);
  lTOBECRITURE.PutValue('E_GENERAL',         StGENERAL_I );
  lTOBECRITURE.PutValue('E_AUXILIAIRE',      StAUXILIAIRE_I );
  lTOBECRITURE.PutValue('E_REFINTERNE',      StREFINTERNE_I);
  lTOBECRITURE.PutValue('E_LIBELLE',         StLIBELLE_I);
  lTOBECRITURE.PutValue('E_NATUREPIECE',     'OD');
  lTOBECRITURE.PutValue('E_QUALIFPIECE',     'N') ;   // ecriture normale
  lTOBECRITURE.PutValue('E_TYPEMVT',         'DIV') ;
  lTOBECRITURE.PutValue('E_ETABLISSEMENT',   VH^.EtablisDefaut) ;
  lTOBECRITURE.PutValue('E_CONTROLETVA',     'RIE') ;
  lTOBECRITURE.PutValue('E_ECRANOUVEAU',     'N') ;
  lTOBECRITURE.PutValue('E_PERIODE',         GetPeriode(DtDATECOMPTABLE_I) ) ;
  lTOBECRITURE.PutValue('E_SEMAINE',         NumSemaine(DtDATECOMPTABLE_I) ) ;
  lTOBECRITURE.PutValue('E_SAISIEEURO',      '-') ;
  CSetMontants( lTOBECRITURE,
                RdDEBIT_I,
                RdCREDIT_I,
                FDevise,
                false,
                true );

  // s'il existe deja des lignes d'analytiques on les affecte a lTOBECRITURE
  if FTOBLigneImputation.Detail.Count <> 0 then
   begin
    // on parcourd les TOB representant les axes analytiques
    for i := 0 to ( FTOBLigneImputation.Detail.Count - 1 ) do
     begin
       // on se place sur les TOB representant les axes analytiques
       FTOBAnalytiq := FTOBLigneImputation.Detail[i];
       // les lignes sont precree dans la fonctioon presendente
       lTOBAxe      := lTOBECRITURE.Detail[i];

       // pour chaque axe analytiques on parcourd les lignes d'analytiques
       for j := 0 to ( FTOBAnalytiq.Detail.Count - 1 ) do
        begin
          // on cree une TOB representant les lignes d'analytique
          lTOBAna            :=TOB.Create('ANALYTIQ',lTOBAxe,-1) ;
         // on recupere les lignes d'analytiques
         FTOBLigneAnalytiq   := FTOBAnalytiq.Detail[j];

         // on affecte les champs lTOBAna avec ceux de lTOBAna
         for k := 1 to ( FTOBLigneAnalytiq.NbChamps ) do
          TransfertGlobalVersLocal ( FTOBLigneAnalytiq.GetNomChamp(k) );

        end; // for

     end; // for
   end; // if

  // on appelle la fonction de ventilation de SAISUTIL
  VentilerTOB( lTOBECRITURE, '' , 0, FDevise.Decimale , FALSE );

  // on parcours les axes analytiques
  for i := 0 to ( lTOBECRITURE.Detail.Count - 1 ) do
   begin
     // on se place sur i Axe
     lTOBAxe      := lTOBECRITURE.Detail[i];

     // s'il n'y a pas de ligne on sort directement
     if lTOBAxe.Detail.Count = 0 then
      break;

     if FTOBLigneImputation.Detail.Count <> 0 then
      begin
       FTOBAnalytiq := FTOBLigneImputation.Detail[i]; // on recupere les lignes existantes
       FTOBAnalytiq.ClearDetail; // on supprime les lignes existantes pour les recreer
      end
       else
        FTOBAnalytiq := TOB.Create('',FTOBLigneImputation,-1);

     // on flag la ligne
     StANA_I := 'X';

      // on parcourd les ligne d'analytiques
     for j := 0 to ( lTOBAxe.Detail.Count - 1 ) do
      begin

        lTOBAna            := lTOBAxe.Detail[j];

        FTOBLigneAnalytiq  := TOB.Create('CRELBQEANALYTIQ',FTOBAnalytiq,-1);

        // on affecte les enr.
        for k := 1 to ( FTOBLigneAnalytiq.NbChamps ) do
         TransfertLocalVersGlobal ( FTOBLigneAnalytiq.GetNomChamp(k) );

      end; // for

     // on pointe sur rien a la fin de la boucle
     FTOBLigneImputation := nil;

   end; // if

  result := true;

 finally
  if assigned( lTOBECRITURE ) then lTOBECRITURE.Free;
 end; // try

end;
{$ENDIF}
{$ENDIF}

function TZReleveBanque.AppelAnalytique ( vMode : TModeAna ) : boolean;
begin
 {$IFNDEF EAGLCLIENT}
 {$IFNDEF SANSCOMPTA}
 result := false;
 exit;

 case vMode of
  TModeVisuAna  : result := AnalytiqManuel('');
  TModeAutoAna  : result := AnalytiqAuto('');
  TModeGuideAna : if VH^.ZSaisieAnal then
                   result := AnalytiqManuel(FCodeGuide)
                    else
                     result := AnalytiqAuto(FCodeGuide);
  TModeManuAna  : if VH^.ZSaisieAnal then
                   result := AnalytiqManuel('')
                    else
                     result := AnalytiqAuto('');
 end; // case
 {$ENDIF}
 {$ENDIF}
end;




function TZReleveBanque.IsVentilable( value : string ) : boolean;
var
 lIndex : integer;
begin

 result := false;
 exit;
 lIndex := ZComptes.GetCompte(Value);
 if ( lIndex > - 1 ) then
  result := ZComptes.IsVentilable(0, lIndex);

end;

function TZReleveBanque.IsVentilable : boolean;
begin
 result := StVENTILABLE_I = 'X';
end;


function TZReleveBanque.GetDeviseCompte( value : string ) : string;
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

procedure TZReleveBanque.SetGeneralImput( Value : string );
var
 lIndex      : integer;
 lStGeneral  : string;
begin

 FTOBLigneImputation.PutValue('CRL_GENERAL', Value);

 lStGeneral := VarToStr(StGENERAL_I);

 lIndex     := ZComptes.GetCompte(lStGeneral);
 // lIndex := ZComptes.GetCompte(Value);
 if ( lIndex > - 1 ) and ( lStGeneral = Value )  then
  begin

   // est ce que le compte est ventilable
   if ZComptes.IsVentilable(0, lIndex) then
    StVENTILABLE_I := 'X'
     else
      StVENTILABLE_I := '-';

   // est ce que le compte est collectif
   if ZComptes.IsCollectif(lIndex) then
    StCOLLECTIF_I := 'X'
     else
      StCOLLECTIF_I := '-';

   if ZComptes.IsTvaAutorise('',lIndex) then
    StTVA_I := ZComptes.GetValue('G_TVA', lIndex) ;

  end; //if

end;

procedure TZReleveBanque.SetCompteContrepartie( Value : string );
begin
 FStCompteContrepartie       := Value;
 FDevise.Code                := GetDeviseCompte(Value);
 GETINFOSDEVISE(FDevise);
 FDevise.Taux                := GetTaux(FDevise.Code , FDevise.DateTaux, V_PGI.DateEntree) ;
end;



procedure TZReleveBanque.RecopierLigneReleve ( Value : TOB );
begin
 FTOBLigneReleve := TOB.Create('CRELBQE',FTOBReleve,-1);
 FTOBLigneReleve.Dupliquer(Value,true,true);
end;

procedure TZReleveBanque.CreerImputation ( vTOBEcr : TOB );
var
 lTOBLigneEcr        : TOB;
 lInNumLigne         : integer;
begin

 //FTOBGuide := vTOBGuide;

 if vTOBEcr.Detail.Count < 0 then exit;

 FTOBImput := TOB.Create('',FTOBLigneReleve,-1);

 lInNumLigne    := 0;

 lTOBLigneEcr   := vTOBEcr.Detail[lInNumLigne];
 FCodeGuide     := lTOBLigneEcr.GetValue('GUIDE'); // on recupere le code du guide courant

 while assigned(lTOBLigneEcr) do
  begin

   AddLigneImput;

   InNUMLIGNE_I     := lInNumLigne + 1;
   StGENERAL_I      := lTOBLigneEcr.GetValue('CRL_GENERAL');
   StLIBELLE_I      := lTOBLigneEcr.GetValue('CRL_LIBELLE');
   CSetMontants( FTOBLigneImputation ,
                 lTOBLigneEcr.GetValue('CRL_DEBIT') ,
                 lTOBLigneEcr.GetValue('CRL_CREDIT') ,
                 FDevise,
                 false,
                 true );

   StREFINTERNE_I   := lTOBLigneEcr.GetValue('CRL_REFINTERNE');
   StGUIDE_I        := lTOBLigneEcr.GetValue('GUIDE');
   FStCurrentGuide  := StGUIDE_I;

   // recherche de l'analytique pour la nouvelle ligne crée
   {$IFNDEF EAGLCLIENT}
   //if IsVentilable then
  //  AnalytiqAuto(FCodeGuide);
   {$ENDIF}

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

procedure TZReleveBanque.RecupereGuide ( vTOBGuide : TOB );
var
 lTOBLigneGuide      : TOB;
 lInNumLigne         : integer;
begin

// FTOBGuide := vTOBGuide;

 if ( vTOBGuide.Detail.Count < 0 ) or not assigned(FTOBImput) then exit;

 lInNumLigne    := 0;

 lTOBLigneGuide := vTOBGuide.Detail[lInNumLigne];
 FCodeGuide     := lTOBLigneGuide.GetValue('EG_GUIDE'); // on recupere le code du guide courant

 while assigned(lTOBLigneGuide) do                   //visureleve
  begin

   FTOBLigneImputation := FTOBImput.Detail[lInNumLigne];

   InNUMLIGNE_I     := lInNumLigne + 1;
   StGENERAL_I      := lTOBLigneGuide.GetValue('EG_GENERAL');
   StLIBELLE_I      := lTOBLigneGuide.GetValue('EG_LIBELLE');
   CSetMontants( FTOBLigneImputation ,
                 Valeur(lTOBLigneGuide.GetValue('EG_DEBITDEV')) ,
                 Valeur(lTOBLigneGuide.GetValue('EG_CREDITDEV')) ,
                 FDevise,
                 false,
                 true );
   StREFINTERNE_I   := lTOBLigneGuide.GetValue('EG_REFINTERNE');
   StGUIDE_I        := lTOBLigneGuide.GetValue('EG_GUIDE');
   FStCurrentGuide  := StGUIDE_I;

   // recherche de l'analytique pour la nouvelle ligne crée
   {$IFNDEF EAGLCLIENT}
   {$IFNDEF SANSCOMPTA}
   if IsVentilable then
    AnalytiqAuto(FCodeGuide);
   {$ENDIF}
   {$ENDIF}
   Inc(lInNumLigne);

   if lInNumLigne < vTOBGuide.Detail.Count then
    lTOBLigneGuide     := vTOBGuide.Detail[lInNumLigne]
     else
      lTOBLigneGuide   := nil;

  end; // while

end;


procedure TZReleveBanque.SetJournalContrepartie( Value : string );
begin

 FStJournalContrepartie := Value;
 RechParamEcriture;

 if FZListJournal.Load(Value) then
  begin
   FStModeSaisie        := FZListJournal.GetValue('J_MODESAISIE');
   FStCompteurNormal    := FZListJournal.GetValue('J_COMPTEURNORMAL');
  end; // if
  
end;

procedure TZReleveBanque.NotifyError( vMessage, vDelphiMessage , vFunctionName : string);
var
 lStMessage : string;
begin

 if vDelphiMessage = '' then
  lStMessage   := vMessage
   else
    lStMessage := vMessage + #13#10 + vDelphiMessage + #13#10 + vFunctionName;

 if ( V_PGI.LastSQLError <> '') or ( FTypeContexte <> TModeAuto ) then
  MessageAlerte( lStMessage );

end;

procedure TZReleveBanque.SetFieldAnneeMoisImput( vIndex : integer ; Value : variant);
begin
 FTOBLigneImputation.PutValue('CRL_PERIODE', intToStr( ent1.GetPeriode(Value) ) );
end;



end.
