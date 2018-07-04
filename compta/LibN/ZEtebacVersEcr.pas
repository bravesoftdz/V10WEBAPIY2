unit ZEtebacVersEcr;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
 HCtrls, HEnt1, HMsgBox,
 {$IFNDEF EAGLCLIENT}
 db,
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
 {$ENDIF}
 // VCL
 Dialogs,
 Windows,
 // Lib

{$IFDEF VER150}
 Variants,
{$ENDIF}

 ULibEcriture,
 UCstEtebac,
 Ent1,
 LettUtil ,
 UtilSais, // pour MajSoldesEcritureTOB
 SAISUTIL, // pour RDevise
 TZ,       // pour le TZF
 ed_tools, // pour le TQRProgressForm
 ZCompte, // pour le TZComptes
 UtilSoc, // pour le marquer publifi
 UtobDebug,
 UObjGen,
 UTOB;

const
 MaxChamps = 200;

type

 TZTresoVersEcr = Class(TObjetCompta)
 private
  FObjCib                         : TObjCIBModePaie;{Pour la conversion des cib en mode de paiement}
  FTOBEcriture                    : TOB;
  FTOBReleve                      : TOB;
  FTOBLettrage                    : TOB ;
  FInNumFolio                     : integer;
  FInNumLigne                     : integer;
  FInNumGroupeEcr                 : integer;
  FInCompteur                     : integer;
  FInPeriode                      : integer;
  FStJournalContrepartie          : string;
  FStModeSaisie                   : string;
  FStCompteurNormal               : string;
  FDtDateComptable                : TDateTime;
  PListeAIntegrer                 : TStringList;
  FStCompteContrepartie           : string ;
  FTypeContexte                   : TTypeContexte;
  FMessCompta                     : TMessageCompta ;
  FBoASupp                        : boolean ;
  FStEtablissement                : string ;
  procedure InitCompteur;
  procedure NumeroteBordereau( vTOBLigneReleve : TOB ; var vTOBLignePiece : TOB ; var vInNumeroPiece : integer ) ;
  function  TransfertModeBordereau( var vInNumeroPiece : integer ) : boolean;
  function  TransfertModeLibre( var vInNumeroPiece : integer ) : boolean;
  function  TransfertModePiece : boolean;
  function  AffecteTOBLigneEcr (vTOBLigneEcriture, vTOBLigneReleve : TOB ) : boolean;
  procedure TransactionSaveEcriture;
//  procedure TransactionSaveMAJCompte;
//  procedure NotifyError( vMessage  : string ; vBoAJoutLog : boolean );
  procedure OnError(sender: TObject; Error: TRecError);
  procedure SoldeCompte ( vTOB : TOB ) ;

 public

  constructor Create( vInfoEcr : TInfoEcriture ); override ;
  destructor  Destroy; override;

  function    LoadReleve ( vListeAIntegrer : TStringList ) : boolean;
  function    Transfert( vTOB1 : TOB ) : boolean;
  {JP 17/01/07 : Ajout du Boolean AvecGeneration, car dans UTomEEXBQ, je lance GenereEcriture avant le
                 Save afin de pouvoir modifier les écritures avant écriture dans la Base}
  function    Save          (var vInNumeroPiece : integer; AvecGeneration : Boolean = True) : TIoErr;
  {JP 17/01/07 : j'ai besoin d'une étape intermédiaire, pour pouvoir modifier les écritures
                 avant leur intégration à la base}
  function    GenereEcriture(var vInNumeroPiece : integer ) : Boolean;

  property    StJournalContrepartie : string            read FStJournalContrepartie  write FStJournalContrepartie;
  property    StCompteContrepartie  : string            read FStCompteContrepartie   write FStCompteContrepartie;
  property    StModeSaisie          : string            read FStModeSaisie           write FStModeSaisie;
  property    StCompteurNormal      : string            read FStCompteurNormal       write FStCompteurNormal;
  property    PTOBLettrage          : TOB               read FTOBLettrage            write FTOBLettrage ;
  property    TypeContexte          : TTypeContexte     read FTypeContexte           write FTypeContexte;
  property    StEtablissement       : string            read FStEtablissement        write FStEtablissement;
  {JP 17/01/07 : j'ai besoin d'accèder à la Tob dans le pointage pour mettre à jour les champs de pointages}
  property    TOBEcriture           : TOB               read FTOBEcriture            write FTOBEcriture;

 end; // class



implementation

uses
 ZEtebac,
 Constantes,
 uLibAnalytique ;

const
 cInMaxEnrBor = 50;


{ TZTResoVersEcr }

constructor TZTResoVersEcr.Create( vInfoEcr : TInfoEcriture );
begin

 inherited ;

 FTOBEcriture       := TOB.Create('',nil,-1);
 FTOBReleve         := TOB.Create('',nil,-1);
 // Objet d'affichage des messages
 FMessCompta        := TMessageCompta.Create('') ;
 Info.OnError       := OnError ;
 {JP 19/01/07 : Récupération des modes de paiement à partir des CIBs}
 FObjCib            := TObjCIBModePaie.Create;
end;

destructor TZTResoVersEcr.Destroy;
begin

 try

  if assigned( FTOBReleve )        then FTOBReleve.Free;
  if assigned( FTOBEcriture )      then FTOBEcriture.Free;
  if assigned(FMessCompta)         then FMessCompta.Free;
  if assigned(FObjCib)             then FObjCib.Free ; 

 finally
  FTOBEcriture                     := nil;
  FMessCompta                      := nil;
  FTOBReleve                       := nil;
  FObjCib                          := nil ;
 end; // try

 inherited;

end;

procedure TZTResoVersEcr.InitCompteur;
begin
 FInNumLigne        := 1;
 FInNumFolio        := 0;
 FInNumGroupeEcr    := 1;
 FInCompteur        := 0;
 FInPeriode         := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 05/01/2007
Modifié le ... :   /  /    
Description .. : - LG - FB 19463 - 05/01/2007 - ajout du ferm(lq) plantais ur 
Suite ........ : les gros volumes
Mots clefs ... : 
*****************************************************************}
function TZTResoVersEcr.LoadReleve( vListeAIntegrer : TStringList ) : boolean;
var
 i                : integer;
 lQ               : TQuery ;
begin

 PListeAIntegrer  := vListeAIntegrer;
 FTOBReleve.ClearDetail ;

 for i := 0  to ( PListeAIntegrer.Count - 1 ) do
  begin

   lQ := OpenSql ( ' select *'                     +
                   ' from CRELBQE '                +
                   ' where CRL_COMPTEUR ='         + VarToStr(PListeAIntegrer.Strings[i]) +
                   ' and CRL_NUMLIGNE <> 0 and ( CRL_ETAT=2 or CRL_ETAT=3 )' , true);

   FTOBReleve.LoadDetailDB('CRELBQE','','',lQ,true) ;

   Ferme(lQ) ;
   
  end ; // for

 if ( FTOBReleve.Detail <> nil )  and ( FTOBReleve.Detail.Count > 0 ) then 
  CZEteAddChampSupp(FTOBReleve.Detail[0],true) ;

 // on trie par date et mois
 FTOBReleve.Detail.Sort('CRL_DATECOMPTABLE;CRL_COMPTEUR;CRL_PERIODE;CRL_NUMLIGNE');

 result := true;

end;


function TZTResoVersEcr.Transfert( vTOB1 : TOB ) : boolean;
var
 i           : integer ;
 j           : integer ;
 lTOBPasBien : TOB ;
begin

 if PListeAIntegrer = nil then
  begin
   PListeAIntegrer := TStringList.Create ;
   FBoASupp        := true ;
  end
   else
    PListeAIntegrer.Clear ;

 lTOBPasBien := TOB.Create( '', nil , -1 ) ;

 try

 for i := 0 to vTOB1.Detail.Count - 1 do
  for j := 0 to vTOB1.Detail[i].Detail.Count - 1 do
   while vTOB1.Detail[i].Detail[j].Detail.Count > 0 do
    begin
     if vTOB1.Detail[i].Detail[j].Detail[0].GetValue('CRL_ETAT') = 1 then
      vTOB1.Detail[i].Detail[j].Detail[0].ChangeParent(lTOBPasBien,-1 )
       else
        begin
         if PListeAIntegrer.IndexOf(intToStr(vTOB1.Detail[i].Detail[j].Detail[0].GetValue('CRL_COMPTEUR')) ) = - 1 then
          PListeAIntegrer.Add(intToStr(vTOB1.Detail[i].Detail[j].Detail[0].GetValue('CRL_COMPTEUR')) ) ;
         vTOB1.Detail[i].Detail[j].Detail[0].ChangeParent(FTOBReleve,-1 ) ;
        end ;
    end ;

 // on trie par date et mois
 if FTOBReleve.Detail.Count > 0 then
  begin
   FTOBReleve.Detail.Sort('CRL_DATECOMPTABLE;CRL_COMPTEUR;CRL_PERIODE;CRL_NUMLIGNE');
   result := true;
  end
   else
    result := false ;
    
 finally
  lTOBPasBien.Free ;
 end ;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/10/2002
Modifié le ... : 25/04/2007
Description .. : -lg - 21/10/2002 - les info de lettrages des tiers n'etaient
Suite ........ : initialises
Suite ........ : - FB 13117 - 11/06/2004 - recopie de l'etablissement
Suite ........ : - FB 15522 - afectation du regime tva
Suite ........ : - FB 16725 - 20/09/2005 - LG - affectation des info de 
Suite ........ : pointage
Suite ........ : - LG - 25/04/2007 - on passe le code guide pour gerer les 
Suite ........ : ventilation par section
Mots clefs ... : 
*****************************************************************}
function TZTResoVersEcr.AffecteTOBLigneEcr ( vTOBLigneEcriture, vTOBLigneReleve : TOB ) : boolean;
var
 lInIndex     : integer;
 lStCompte    : string;
 lInResult    : integer ;
 lBoLettrable : boolean ;
 lModePaie    : string; {JP 19/01/07}
begin
  result := false;
  CPutDefautEcr(vTOBLigneEcriture) ;
  vTOBLigneEcriture.PutValue('E_JOURNAL'          , vTOBLigneReleve.getValue('CRL_JOURNAL') );
  vTOBLigneEcriture.PutValue('E_EXERCICE'         , QuelExoDt ( vTOBLigneReleve.getValue('CRL_DATECOMPTABLE') ) );
  vTOBLigneEcriture.PutValue('E_DATECOMPTABLE'    , vTOBLigneReleve.getValue('CRL_DATECOMPTABLE') );
  vTOBLigneEcriture.PutValue('E_NUMEROPIECE'      , FInNumFolio );
  vTOBLigneEcriture.PutValue('E_NUMLIGNE'         , FInNumLigne );
  vTOBLigneEcriture.PutValue('E_GENERAL'          , vTOBLigneReleve.getValue('CRL_GENERAL') );
  if vTOBLigneReleve.getValue('CRL_ETABLISSEMENT') <> '' then
   vTOBLigneEcriture.PutValue('E_ETABLISSEMENT'    , vTOBLigneReleve.getValue('CRL_ETABLISSEMENT') )
    else
     if FStEtablissement <> '' then
      vTOBLigneEcriture.PutValue('E_ETABLISSEMENT'    , FStEtablissement )
       else
        vTOBLigneEcriture.PutValue('E_ETABLISSEMENT'    , VH^.EtablisDefaut ) ;
  vTOBLigneEcriture.PutValue('E_CONTREPARTIEGEN'  , vTOBLigneReleve.getValue('CRL_CONTREPARTIEGEN') );
  vTOBLigneEcriture.PutValue('E_AUXILIAIRE'       , vTOBLigneReleve.getValue('CRL_AUXILIAIRE') );
  Info.Devise.AffecteTaux(vTOBLigneReleve.getValue('CRL_DATECOMPTABLE')) ;
  CSetMontants ( vTOBLigneEcriture ,
                 vTOBLigneReleve.getValue('CRL_DEBIT') ,
                 vTOBLigneReleve.getValue('CRL_CREDIT') ,
                 Info.Devise.Dev ,
                 true );
  vTOBLigneEcriture.PutValue('E_REFINTERNE'       , vTOBLigneReleve.getValue('CRL_REFINTERNE') );
  {JP 19/01/07 : Le libellé enrichi est stocké dans la référence externe}
  vTOBLigneEcriture.PutValue('E_REFINTERNE'       , vTOBLigneReleve.getValue('CRL_REFINTERNE') );
  vTOBLigneEcriture.PutValue('E_LIBELLE'          , vTOBLigneReleve.getValue('CRL_LIBELLE') );
  vTOBLigneEcriture.PutValue('E_NATUREPIECE'      , vTOBLigneReleve.getValue('CRL_NATUREPIECE') );
  vTOBLigneEcriture.PutValue('E_TVA'              , vTOBLigneReleve.getValue('CRL_TVA') );
  vTOBLigneEcriture.PutValue('E_ANA'              , vTOBLigneReleve.getValue('CRL_ANA') );
  vTOBLigneEcriture.PutValue('E_QUALIFPIECE'      , 'N');
  vTOBLigneEcriture.PutValue('E_VALIDE'           , '-' );
  vTOBLigneEcriture.PutValue('E_NUMECHE'          , 0 );
  vTOBLigneEcriture.PutValue('E_ECRANOUVEAU'      , 'N' );
  vTOBLigneEcriture.PutValue('E_QUALIFORIGINE'    , 'TRE' );
  vTOBLigneEcriture.PutValue('E_ECRANOUVEAU'      , 'N' );
  vTOBLigneEcriture.PutValue('E_IO'               , 'X' );
  vTOBLigneEcriture.PutValue('E_RIB'              , CGetRIB(vTOBLigneEcriture));
  {JP 17/01/07 : On mémorise éventuellement le code de pointage, la date et la référence de pointage seront
                 mis à jour plus tard}
  vTOBLigneEcriture.PutValue('E_NATURETRESO'      , vTOBLigneReleve.getValue('CRL_CODEPOINTAGE') );

  CSetPeriode( vTOBLigneEcriture );

  // controle du journal
  Info.Journal.Load([vTOBLigneReleve.getValue('CRL_JOURNAL')]) ;
  lInResult := CIsValidJournal(vTOBLigneEcriture,Info) ;
  if lInResult <> RC_PASERREUR then
   begin
    NotifyError(lInResult,'' ) ;
    exit ;
   end ;

  // controle du compte
  lStCompte := vTOBLigneReleve.getValue('CRL_GENERAL');
  lInIndex  := Info.Compte.GetCompte(lStCompte);

  if lInIndex > -1 then
   begin

    lInResult := CIsValidCompte(vTOBLigneEcriture,Info) ;
    if lInResult <> RC_PASERREUR then
     begin
      NotifyError(lInResult,'' ) ;
      exit ;
     end ;

    if Info.Compte.IsCollectif(lInIndex) then
     begin
      // chargement du tier
      if not Info.LoadAux(vTOBLigneReleve.getValue('CRL_AUXILIAIRE'))  then exit ;
      lBoLettrable := (Info.Aux_GetValue('T_LETTRABLE') = 'X') ;
     end // if
      else
       lBoLettrable := false ;

    if (Info.Compte.IsLettrable(lInIndex)) or lBoLettrable then
     CRemplirInfoLettrage(vTOBLigneEcriture)
      else
       CSupprimerInfoLettrage(vTOBLigneEcriture);

    if Info.Compte.IsPointable(lInIndex) then CRemplirInfoPointage(vTOBLigneEcriture);

    CGetRegimeTVA (vTOBLigneEcriture,Info) ;
    CGetEch (vTOBLigneEcriture,Info) ;

    if vTOBLigneReleve.getValue('CRL_ALETTRER') = 'X' then
     begin
      vTOBLigneEcriture.PutValue('E_ETATLETTRAGE'   , vTOBLigneReleve.getValue('CRL_ETATLETTRAGE') );
      vTOBLigneEcriture.PutValue('E_LETTRAGE'       , vTOBLigneReleve.getValue('CRL_LETTRAGE') );
      vTOBLigneEcriture.PutValue('E_DATEPAQUETMAX'  , vTOBLigneReleve.getValue('CRL_DATEPAQUETMAX') );
      vTOBLigneEcriture.PutValue('E_DATEPAQUETMIN'  , vTOBLigneReleve.getValue('CRL_DATEPAQUETMIN') );
      vTOBLigneEcriture.PutValue('E_COUVERTURE'     , vTOBLigneReleve.getValue('CRL_COUVERTURE') );
      vTOBLigneEcriture.PutValue('E_COUVERTUREDEV'  , vTOBLigneReleve.getValue('CRL_COUVERTUREDEV') );
      vTOBLigneEcriture.PutValue('E_TRESOSYNCHRO'   , ets_Lettre) ;
     end ;

    if Info.Compte.IsVentilable(-1,lInIndex) then
     begin
      vTOBLigneEcriture.PutValue('E_ANA' , 'X') ;
      AlloueAxe( vTOBLigneEcriture ) ; // SBO 25/01/2006
      CVentilerTOB(vTOBLigneEcriture,Info , false , vTOBLigneReleve.getValue('CRL_GUIDE') ) ;
     end; // if

    {JP 19/01/07 : Gestion éventuelle des CIB si l'on est sur la ligne bancaire}
    if vTOBLigneReleve.getValue('CRL_CODEPOINTAGE') <> '' then begin
      if ((vTOBLigneReleve.GetDouble('CRL_DEBIT') - vTOBLigneReleve.GetDouble('CRL_CREDIT')) > 0) or
         ((vTOBLigneReleve.GetDouble('CRL_DEBITDEV') - vTOBLigneReleve.GetDouble('CRL_CREDITDEV')) > 0) then
        lModePaie := 'ENC'
      else
        lModePaie := 'DEC';

      lModePaie := FObjCib.GetModePaie(vTOBLigneReleve.GetString('CRL_GENERAL'),
                                       vTOBLigneReleve.GetString('CRL_CODEAFB'),
                                       lModePaie);
      if lModePaie <> '' then
        vTOBLigneEcriture.PutValue('E_MODEPAIE', lModePaie);
      {CRemplirInfoPointage met 'E_TRESOSYNCHRO' à ets_Nouveau sauf si l'on est en Tréso}
      vTOBLigneEcriture.PutValue('E_TRESOSYNCHRO', ets_Nouveau);
     end;
   end; // if

 result := true;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/07/2005
Modifié le ... :   /  /    
Description .. : FB 15538 - LG  - 05/07/2005 - generation de l'ecriture 
Suite ........ : apres la saisie de l'imput
Mots clefs ... : 
*****************************************************************}
procedure TZTResoVersEcr.NumeroteBordereau ( vTOBLigneReleve : TOB ; var vTOBLignePiece : TOB ; var vInNumeroPiece : integer ) ;
var
 lBoCouperFolio : boolean;
begin

 lBoCouperFolio  := false; //( FInNumLigne > cInMaxEnrBor );
 //result          := vTOBLignePiece;

 if ( vInNumeroPiece <> -1 ) then
  begin
   if ( FInPeriode <> vTOBLigneReleve.GetValue('CRL_PERIODE') ) then
    begin
     FInNumFolio     := vInNumeroPiece ;
     FInPeriode      := vTOBLigneReleve.GetValue('CRL_PERIODE');
     FInNumGroupeEcr := 1;
     FInNumLigne     := 1;
     vTOBLignePiece  := TOB.Create( '' , FTOBEcriture , -1 );
     exit ;
    end
     else
      FInCompteur := vTOBLigneReleve.GetValue('CRL_COMPTEUR');
  end ;  

 if ( FInPeriode <> vTOBLigneReleve.GetValue('CRL_PERIODE') ) then
  begin
   // on recherche en base un numero de piece pour la periode concerné
   FInNumFolio := GetNewNumJal( vTOBLigneReleve.GetValue('CRL_JOURNAL') ,
                                   true,
                                   vTOBLigneReleve.GetValue('CRL_DATECOMPTABLE')
                                    );
   FInPeriode      := vTOBLigneReleve.GetValue('CRL_PERIODE');
   FInNumGroupeEcr := 1;
   FInNumLigne     := 1;
   vTOBLignePiece  := TOB.Create( '' , FTOBEcriture , -1 );
  end // if
   else
     if ( FInCompteur <> vTOBLigneReleve.GetValue('CRL_COMPTEUR') ) and not lBoCouperFolio then
      Inc(FInNumGroupeEcr)
       else
        if ( FInCompteur <> vTOBLigneReleve.GetValue('CRL_COMPTEUR') ) and lBoCouperFolio then
         begin
           // la periode reste la même -> on incremente le numero de folio
           Inc(FInNumFolio);
           FInNumGroupeEcr := 1;
           FInNumLigne     := 1;
         end;

 FInCompteur := vTOBLigneReleve.GetValue('CRL_COMPTEUR');
 
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/03/2005
Modifié le ... : 05/07/2005
Description .. : - LG - 22/03/2005 - FB 15522 - affectation de e_regimetva
Suite ........ : FB 15538 - LG  - 05/07/2005 - generation de l'ecriture 
Suite ........ : apres la saisie de l'imput
Mots clefs ... : 
*****************************************************************}
function TZTResoVersEcr.TransfertModeBordereau( var vInNumeroPiece : integer ) : boolean;
var
 i                 : integer;
 lTOBLigneReleve   : TOB;
 lTOBLignePiece    : TOB;
 lTOBLigneEcriture : TZF;
 lRecError         : TRecError ;
begin

 result          := false;
 lTOBLignePiece:=Nil;
 InitCompteur;

 for i := 0 to ( FTOBReleve.Detail.Count - 1 ) do
  begin

   lTOBLigneReleve   := FTOBReleve.Detail[i];
   NumeroteBordereau(lTOBLigneReleve,lTOBLignePiece,vInNumeroPiece);

   lTOBLigneEcriture := TZF.Create( 'ECRITURE' , lTOBLignePiece , -1);

   if not AffecteTOBLigneEcr ( lTOBLigneEcriture , lTOBLigneReleve ) then exit ;

   lTOBLigneEcriture.PutValue('E_MODESAISIE'       , 'BOR');
   lTOBLigneEcriture.PutValue('E_NUMGROUPEECR'     , FInNumGroupeEcr );

   inc(FInNumLigne);

   // gestion de la fenetre d'attente
   MoveCurProgressForm('Lecture de l''écriture : ' + IntToStr( lTOBLigneReleve.GetValue('CRL_COMPTEUR') ) ) ;

   end; // for

  vInNumeroPiece := FInNumFolio ;  

  for i := 0 to FTOBEcriture.Detail.Count - 1 do
   begin
    CAffectCompteContrePartie(FTOBEcriture.Detail[i],Info);
    CAffectRegimeTva(FTOBEcriture.Detail[i]) ;
    if vInNumeroPiece <> - 1 then
     CNumeroLigneBor(FTOBEcriture.Detail[i])
      else
       CNumeroPiece(FTOBEcriture.Detail[i]);
    lRecError := CIsValidSaisiePiece(FTOBEcriture.Detail[i],Info) ;
    if lRecError.RC_Error <> RC_PASERREUR then
     begin
      NotifyError(lRecError.RC_Error,lRecError.RC_Message , '' ) ;
      exit ;
     end ;
   end; // if

 result := true;

end;

procedure TZTResoVersEcr.SoldeCompte ( vTOB : TOB ) ;
var
 i        : integer ;
 lRdSolde : double ;
begin
 i        := 0 ;
 lRdSolde := 0 ;
 while i < ( vTOB.Detail.Count - 1 ) do
  begin
   if FStCompteContrepartie = vTOB.Detail[i].GetValue('E_GENERAL') then
    begin
     lRdSolde := Arrondi( lRdSolde + vTOB.Detail[i].GetValue('E_DEBIT') - vTOB.Detail[i].GetValue('E_CREDIT') , 2 ) ;
     vTOB.Detail[i].Free ;
    end
     else
      Inc(i) ;
  end ;

 lRdSolde := Arrondi( lRdSolde + vTOB.Detail[i].GetValue('E_DEBIT') - vTOB.Detail[i].GetValue('E_CREDIT') , 2 ) ;

 if lRdSolde > 0 then
  CSetMontants(vTOB.Detail[i],lRdSolde,0,Info.Devise.Dev,true)
   else
    CSetMontants(vTOB.Detail[i],0,lRdSolde*(-1),Info.Devise.Dev,true) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/07/2005
Modifié le ... :   /  /    
Description .. : FB 15538 - LG  - 05/07/2005 - generation de l'ecriture 
Suite ........ : apres la saisie de l'imput
Mots clefs ... : 
*****************************************************************}
function TZTResoVersEcr.TransfertModeLibre( var vInNumeroPiece : integer ) : boolean;
var
 i                 : integer;
 lTOBLigneReleve   : TOB;
 lTOBLignePiece    : TOB;
 lTOBLigneEcriture : TZF;
 lRecError         : TRecError ;
begin

 result          := false;
 lTOBLignePiece  := nil;
 InitCompteur;

 // controle du journal
 Info.Journal.Load([StJournalContrepartie]) ;

  for i := 0 to ( FTOBReleve.Detail.Count - 1 ) do
   begin

    lTOBLigneReleve := FTOBReleve.Detail[i];
    NumeroteBordereau(lTOBLigneReleve,lTOBLignePiece,vInNumeroPiece);

    lTOBLigneEcriture := TZF.Create( 'ECRITURE' , lTOBLignePiece , -1);

    if not AffecteTOBLigneEcr ( lTOBLigneEcriture , lTOBLigneReleve ) then exit ;

    lTOBLigneEcriture.PutValue('E_MODESAISIE'       , 'LIB');
    lTOBLigneEcriture.PutValue('E_NUMGROUPEECR'     , 1);

    inc(FInNumLigne);

    // gestion de la fenetre d'attente                                                à
   MoveCurProgressForm( 'Lecture de l''écriture : ' + IntToStr( lTOBLigneReleve.GetValue('CRL_COMPTEUR') ) ) ;

   end; // for

  vInNumeroPiece := FInNumFolio ;

  if Info.Journal.GetValue('J_TRESOLIBRE') = 'X' then
   for i := 0 to FTOBEcriture.Detail.Count - 1 do
    SoldeCompte(FTOBEcriture.Detail[i]) ;

  for i := 0 to FTOBEcriture.Detail.Count - 1 do
   begin
    CAffectCompteContrePartie(FTOBEcriture.Detail[i],Info);
    if vInNumeroPiece <> - 1 then
     CNumeroLigneBor(FTOBEcriture.Detail[i])
      else
       CNumeroPiece(FTOBEcriture.Detail[i]);
    lRecError := CIsValidSaisiePiece(FTOBEcriture.Detail[i],Info) ;
    if lRecError.RC_Error <> RC_PASERREUR then
     begin
      NotifyError(lRecError.RC_Error,lRecError.RC_Message , '' ) ;
      exit ;
     end ;
   end; // if

  result := true;

end;


function TZTResoVersEcr.TransfertModePiece : boolean;
var
 i                 : integer;
 lTOBLigneReleve   : TOB;
 lTOBLigneEcriture : TZF;
 lTOBLignePiece    : TOB;
 lStAnneeMois      : string;
 lRecError         : TRecError ;
begin

 result          := false;

 lStAnneeMois    := '';
 lTOBLignePiece := nil;
 InitCompteur;

 for i := 0 to ( FTOBReleve.Detail.Count - 1 ) do
  begin

   lTOBLigneReleve := FTOBReleve.Detail[i];

   if  FInCompteur <>  lTOBLigneReleve.GetValue('CRL_COMPTEUR') then
    begin

     FInCompteur       := lTOBLigneReleve.GetValue('CRL_COMPTEUR');
     FInNumLigne       := 1;
     FDtDateComptable  := lTOBLigneReleve.GetValue('CRL_DATECOMPTABLE');
     FInNumFolio       := GetNewNumJal( lTOBLigneReleve.GetValue('CRL_JOURNAL') ,
                                         true,
                                         FDtDateComptable
                                          );
//     FListeLog.Add('Création de la pièce numéro ' + intToStr(FInNumFolio) + ' pour la période '  +
//                    FormatDateTime('mmmmm yyyy', lTOBLigneReleve.GetValue('CRL_DATECOMPTABLE') ) );
     lTOBLignePiece   := TOB.Create( '' , FTOBEcriture , -1 );

    end; // if

   lTOBLigneEcriture := TZF.Create( 'ECRITURE' , lTOBLignePiece , -1);

   if not AffecteTOBLigneEcr ( lTOBLigneEcriture , lTOBLigneReleve ) then exit ;

   lTOBLigneEcriture.PutValue('E_MODESAISIE'       , '-');
   lTOBLigneEcriture.PutValue('E_NUMGROUPEECR'     , 1);

   // gestion de la fenetre d'attente
   MoveCurProgressForm('Lecture de l''écriture : ' + IntToStr( lTOBLigneReleve.GetValue('CRL_COMPTEUR') ) ) ;
   inc(FInNumLigne);

  end; // for
  // a refaire
  for i := 0 to FTOBEcriture.Detail.Count - 1 do
   begin
    CAffectCompteContrePartie(FTOBEcriture.Detail[i],Info);
    lRecError := CIsValidSaisiePiece(FTOBEcriture.Detail[i],Info) ;
    if lRecError.RC_Error <> RC_PASERREUR then
     begin
      NotifyError(lRecError.RC_Error,lRecError.RC_Message , '' ) ;
      exit ;
     end ;
    CNumeroPiece(FTOBEcriture.Detail[i]);
   end; // if

  result          := true;

end;

procedure TZTresoVersEcr.TransactionSaveEcriture;
var
 i   : integer;
begin

  (* JP 17/01/07 : Déplacer vers .Save comme pour les saisies libre et bordereau
  // numerotation et transfert vers la table ecriture pour le mode piece
  if ( FStModeSaisie[1] = '-' ) then
   begin
   if not TransfertModePiece  then
   begin
    V_PGI.IoError := oeSaisie;
    NotifyError( -1 , 'Erreur sur le transfert des enregistrements' );
    exit;
    end;
   end; // if
   *)
 // gestion de la fenetre d'attente
 MoveCurProgressForm(cStTexteEnrEnCours) ;

 try

 //tobdebug(FTOBEcriture) ;

 for i := 0 to FTOBEcriture.Detail.Count - 1 do
  if not CEnregistreSaisie(FTOBEcriture.Detail[i],false,false,true,Info) then
   begin
    V_PGI.IOError := oeSaisie ;
    exit ;
   end ;


 if ( PTOBLettrage <> nil ) and ( PTOBLettrage.Detail <> nil ) and ( PTOBLettrage.Detail.Count > 0 ) then
  for i := 0 to PTOBLettrage.Detail.Count - 1 do
   if not GoReqMajLet( TOBM(FTOBLettrage.Detail[i]),
                       FTOBLettrage.Detail[i].GetValue('E_GENERAL') ,
                       FTOBLettrage.Detail[i].GetValue('E_AUXILIAIRE'),
                       now,
                       true ) then
    begin
     V_PGI.IOError := oeLettrage ;
     Exit ;
    end ;


 // gestion de la fenetre d'attente
 MoveCurProgressForm('Suppression des lignes de la table CRELBQE') ;

 for i := 0 to ( PListeAIntegrer.Count - 1 ) do
  ExecuteSQL('delete from CRELBQE where CRL_COMPTEUR= ' + VarToStr(PListeAIntegrer.Strings[i]) );

 except
  on E : Exception do
   PGIError('Erreur lors de l''enregistrement en base !' + #10#13 + E.Message ) ;
 end ;


end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 28/10/2002
Modifié le ... :   /  /
Description .. : -28/10/2002 - la fct MajSoldesEcritures n'etait pas appelle
Suite ........ : avec les bon parametres ( les montants etait signes )
Mots clefs ... : 
*****************************************************************}
{procedure TZTresoVersEcr.TransactionSaveMAJCompte;
var
 i              : integer;
begin
   // gestion de la fenetre d'attente
 MoveCurProgressForm('Mise à jour des comptes') ;

 for i := 0 to FTOBEcriture.Detail.Count - 1 do
  MajSoldesEcritureTOB( FTOBEcriture.Detail[i], true);

end;}

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 25/04/2007
Modifié le ... :   /  /    
Description .. : - LG - 25/04/2007 - supression d'un warning
Mots clefs ... : 
*****************************************************************}
function TZTresoVersEcr.GenereEcriture(var vInNumeroPiece : integer ) : Boolean;
{---------------------------------------------------------------------------------------}
begin
 result := false ;
    // numerotation et transfert vers la table ecriture
  case FStModeSaisie[1] of
   {JP 17/01/07 : mise ne place du même fonctionnement pour toutes les saisie
   '-' : lBoResult := true;}
   '-' : Result := TransfertModePiece;
   'L' : Result := TransfertModeLibre( vInNumeroPiece );
   'B' : Result := TransfertModeBordereau( vInNumeroPiece );
  end; // case
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : - LG - 18/09/2007 - FB 21343 - suppression du blocage 
Suite ........ : monoposte
Mots clefs ... : 
*****************************************************************}
function TZTresoVersEcr.Save ( var vInNumeroPiece : integer; AvecGeneration : Boolean = True) : TIoErr;
var
 lBoResult : boolean;
begin

 result        := oeOk;

 try

  {JP 17/01/07 : la génération des pièces est déplacée dans GenereEcriture.
                 Dans UTomEEXBQ, la génération est déjà faite}
  if AvecGeneration then
    lBoResult := GenereEcriture(vInNumeroPiece)
  else
    lBoResult := True;

  if not lBoResult then
   begin
    result := oeSaisie;
    NotifyError( -1 , 'Erreur sur le transfert des enregistrements');
    exit;
   end; // if

  if Transactions( TransactionSaveEcriture , 1 ) <> oeOK then
   begin
    result := oeStock;
    NotifyError( -1 , 'Erreur sur l''enregistrement en base' );
    exit;
   end; // if

 MarquerPublifi(TRUE) ;

 finally
  if FBoASupp then PListeAIntegrer.Free ;
  FBoASupp := false ;
 end; // try

end;

procedure TZTresoVersEcr.OnError (sender : TObject; Error : TRecError ) ;
begin
 if trim(Error.RC_Message)<>'' then PGIInfo(Error.RC_Message, 'Intégration des écritures' )
  else
   //if not (Error.RC_Error in [RC_COMPTEINEXISTANT,RC_NATAUX, RC_AUXINEXISTANT, RC_AUXOBLIG , RC_PASCOLLECTIF ] ) then
    FMessCompta.Execute(Error.RC_Error);
end;




end.
