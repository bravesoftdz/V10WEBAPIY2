unit ZTresoVersEcr;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
 HCtrls, HEnt1, HMsgBox,
 {$IFNDEF EAGLCLIENT}
 db,dbTables,
 {$ENDIF}
 // VCL
 Dialogs,
 Windows,
 // Lib
 ULibEcriture,
 UCstEcriture,
 Ent1,
 ZReleveBanque,
 UtilSais, // pour MajSoldesEcritureTOB
 SAISUTIL, // pour RDevise
 TZ,       // pour le TZF
 ed_tools, // pour le TQRProgressForm
 ZCompte, // pour le TZComptes
 UTOB;

const
 MaxChamps = 200;

type

 TZTresoVersEcr = Class
 private
  FTOBEcriture                    : TOB;
  FTOBReleve                      : TOB;
  FTOBAnaReleve                   : TOB;
  FZReleveBqe                     : TZReleveBanque;
  FInNumFolio                     : integer;
  FInNumLigne                     : integer;
  FInNumGroupeEcr                 : integer;
  FInCompteur                     : integer;
  FInPeriode                      : integer;
  FStJournalContrepartie          : string;
  FStModeSaisie                   : string;
  FStCompteurNormal               : string;
  FDtDateComptable                : TDateTime;
  FTT                             : TQRProgressForm;
  FListeAIntegrer                 : TStringList;
  FListeLog                       : TStringList;
  FTypeContexte                   : TTypeContexte;
  FZComptes                       : TZCompte;  // objet de gestion des comptes

  procedure InitCompteur;
  function  NumeroteBordereau( vTOBLigneReleve , vTOBLignePiece : TOB ) : TOB;
  function  TransfertModeBordereau : boolean;
  function  TransfertModeLibre : boolean;
  function  TransfertModePiece : boolean;
  procedure AffecteTOBLigneEcr (vTOBLigneEcriture, vTOBLigneReleve : TOB );
  procedure TransactionSaveEcriture;
  procedure TransactionSaveMAJCompte;
  procedure NotifyError( vMessage  : string ; vBoAJoutLog : boolean );

 public

  constructor Create;
  destructor  Destroy; override;

  function    LoadReleve ( vListeAIntegrer : TStringList ) : boolean;

  function    Save : TIoErr;

  property    StJournalContrepartie : string            read FStJournalContrepartie  write FStJournalContrepartie;
  property    StModeSaisie          : string            read FStModeSaisie           write FStModeSaisie;
  property    StCompteurNormal      : string            read FStCompteurNormal       write FStCompteurNormal;
  property    TT                    : TQRProgressForm   read FTT                     write FTT;
  property    TypeContexte          : TTypeContexte     read FTypeContexte           write FTypeContexte;

 end; // class



implementation

const
 cInMaxEnrBor = 50;


{ TZTResoVersEcr }

constructor TZTResoVersEcr.Create;
begin

 FTOBEcriture       := TOB.Create('',nil,-1);
 FTOBReleve         := TOB.Create('',nil,-1);
 FTOBAnaReleve      := TOB.Create('',nil,-1);
 FListeLog          := TStringList.Create;
 FZReleveBqe        := TZReleveBanque.Create(nil);
 FZComptes          := TZCompte.Create;

end;

destructor TZTResoVersEcr.Destroy;
begin

 try

  if assigned( FTOBReleve )        then FTOBReleve.Free;
  if assigned( FTOBEcriture )      then FTOBEcriture.Free;
  if assigned( FListeLog )         then FListeLog.Free;
  if assigned( FZReleveBqe )       then FZReleveBqe.Free;
  if assigned( FTOBAnaReleve )     then FTOBAnaReleve.Free;
  if assigned( FZComptes )         then FZComptes.Free;

 finally
  FTOBEcriture                     := nil;
  FTOBReleve                       := nil;
  FListeLog                        := nil;
  FZReleveBqe                      := nil;
  FTOBAnaReleve                    := nil;
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

function TZTResoVersEcr.LoadReleve( vListeAIntegrer : TStringList ) : boolean;
var
 i                : integer;
 lTOBLigneImput   : TOB;
begin

 result           := false;
 FListeAIntegrer  := vListeAIntegrer;
 FZReleveBqe.AddLigneReleve;
 FZReleveBqe.StETAT_R := '3';

 for i := 0  to ( FListeAIntegrer.Count - 1 ) do
  begin

   if assigned(FZReleveBqe.TOBLigneReleve) then
    FZReleveBqe.TOBLigneReleve.ClearDetail;
   FZReleveBqe.InCOMPTEUR_R := VarToStr(FListeAIntegrer.Strings[i]);
   FZReleveBqe.LoadImputation;

   if not assigned(FZReleveBqe.TOBImput ) then continue; // pas de ligne trouve pour ce releve

   while FZReleveBqe.TOBImput.Detail.Count > 0 do
    begin

     lTOBLigneImput := FZReleveBqe.TOBImput.Detail[0];

     while ( lTOBLigneImput.Detail.Count > 0 ) do
      lTOBLigneImput.Detail[0].ChangeParent(FTOBAnaReleve , -1);

     lTOBLigneImput.ChangeParent( FTOBReleve , - 1);

    end; // while

  end; // for

 // on trie par date et mois
 FTOBReleve.Detail.Sort('CRL_DATECOMPTABLE;CRL_COMPTEUR;CRL_PERIODE;CRL_NUMLIGNE');

 result := true;

end;

procedure TZTResoVersEcr.AffecteTOBLigneEcr ( vTOBLigneEcriture, vTOBLigneReleve : TOB );
var
 lInIndex  : integer;
 lStCompte : string;
 i         : integer;
begin
  vTOBLigneEcriture.PutValue('E_JOURNAL'          , vTOBLigneReleve.getValue('CRL_JOURNAL') );
  vTOBLigneEcriture.PutValue('E_EXERCICE'         , QuelExoDt ( vTOBLigneReleve.getValue('CRL_DATECOMPTABLE') ) );
  vTOBLigneEcriture.PutValue('E_DATECOMPTABLE'    , vTOBLigneReleve.getValue('CRL_DATECOMPTABLE') );
  vTOBLigneEcriture.PutValue('E_NUMEROPIECE'      , FInNumFolio );
  vTOBLigneEcriture.PutValue('E_NUMLIGNE'         , FInNumLigne );
  vTOBLigneEcriture.PutValue('E_GENERAL'          , vTOBLigneReleve.getValue('CRL_GENERAL') );
  vTOBLigneEcriture.PutValue('E_CONTREPARTIEGEN'  , vTOBLigneReleve.getValue('CRL_CONTREPARTIEGEN') );
  vTOBLigneEcriture.PutValue('E_AUXILIAIRE'       , vTOBLigneReleve.getValue('CRL_AUXILIAIRE') );
  vTOBLigneEcriture.PutValue('E_DEBIT'            , vTOBLigneReleve.getValue('CRL_DEBIT') ) ;
  vTOBLigneEcriture.PutValue('E_CREDIT'           , vTOBLigneReleve.getValue('CRL_CREDIT') ) ;
  vTOBLigneEcriture.PutValue('E_DEBITDEV'         , vTOBLigneReleve.getValue('CRL_DEBITDEV') );
  vTOBLigneEcriture.PutValue('E_CREDITDEV'        , vTOBLigneReleve.getValue('CRL_CREDITDEV') );
  vTOBLigneEcriture.PutValue('E_DEBITEURO'        , vTOBLigneReleve.getValue('CRL_DEBITEURO') ) ;
  vTOBLigneEcriture.PutValue('E_CREDITEURO'       , vTOBLigneReleve.getValue('CRL_CREDITEURO') ) ;
  vTOBLigneEcriture.PutValue('E_SAISIEEURO'       , vTOBLigneReleve.getValue('CRL_SAISIEEURO') ) ;
  vTOBLigneEcriture.PutValue('E_DEVISE'           , vTOBLigneReleve.getValue('CRL_DEVISE') ) ;
  vTOBLigneEcriture.PutValue('E_TAUXDEV'          , vTOBLigneReleve.getValue('CRL_TAUXDEV') ) ;
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
  CSetPeriode( vTOBLigneEcriture );
  lStCompte := vTOBLigneReleve.getValue('CRL_GENERAL');
  lInIndex  := FZComptes.GetCompte(lStCompte);

  if lInIndex > -1 then
   begin
    if FZComptes.IsLettrable(lInIndex) then
     CRemplirInfoLettrage(vTOBLigneEcriture)
      else
       CSupprimerInfoLettrage(vTOBLigneEcriture);

     if FZComptes.IsVentilable(-1,lInIndex) then
      begin
       for i := 1 to MAXAXE do TOB.Create('A'+IntToStr(i), vTOBLigneEcriture, -1) ;
       VentilerTOB(vTOBLigneEcriture,'',0,V_PGI.OkDecV,false) ;
       vTOBLigneEcriture.PutValue('E_ANA' , 'X') ;
      end; // if

    if (FZComptes.IsPointable(lInIndex)) or (FZComptes.IsBqe(lInIndex))  then
     CRemplirInfoPointage(vTOBLigneEcriture);

   end; // if

end;


function TZTResoVersEcr.NumeroteBordereau ( vTOBLigneReleve , vTOBLignePiece : TOB ) : TOB;
var
 lBoCouperFolio : boolean;
begin

 lBoCouperFolio  := false; //( FInNumLigne > cInMaxEnrBor );
 result          := vTOBLignePiece;

 if ( FInPeriode <> vTOBLigneReleve.GetValue('CRL_PERIODE') ) then
  begin
   // on recherche en base un numero de piece pour la periode concerné
   FInNumFolio := GetNewNumJal( vTOBLigneReleve.GetValue('CRL_JOURNAL') ,
                                   true,
                                   vTOBLigneReleve.GetValue('CRL_DATECOMPTABLE'),
                                   true
                                    );
   FInPeriode      := vTOBLigneReleve.GetValue('CRL_PERIODE');
   FInNumGroupeEcr := 1;
   FInNumLigne     := 1;

   FListeLog.Add('Création du bordereau numéro ' + intToStr(FInNumFolio) + ' pour la période '  +
                  FormatDateTime('dd/mm/yyyy', vTOBLigneReleve.GetValue('CRL_DATECOMPTABLE') ) );
   // creation bordereau
   result   := TOB.Create( '' , FTOBEcriture , -1 );
   
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

           FListeLog.Add('Création du bordereau numéro ' + intToStr(FInNumFolio) + ' pour la période '  +
                          FormatDateTime('dd/mm/yyyy', vTOBLigneReleve.GetValue('CRL_DATECOMPTABLE') ) );
           // creation bordereau

         end;

end;


function TZTResoVersEcr.TransfertModeBordereau : boolean;
var
 i                 : integer;
 lTOBLigneReleve   : TOB;
 lTOBLignePiece    : TOB;
 lTOBLigneEcriture : TZF;
begin

 result          := false;

 InitCompteur;

 for i := 0 to ( FTOBReleve.Detail.Count - 1 ) do
  begin

   lTOBLigneReleve   := FTOBReleve.Detail[i];

   lTOBLignePiece    := NumeroteBordereau(lTOBLigneReleve,lTOBLignePiece);

   FInCompteur       := lTOBLigneReleve.GetValue('CRL_COMPTEUR');

   lTOBLigneEcriture := TZF.Create( 'ECRITURE' , lTOBLignePiece , -1);

   AffecteTOBLigneEcr ( lTOBLigneEcriture , lTOBLigneReleve );

   lTOBLigneEcriture.PutValue('E_MODESAISIE'       , 'BOR');
   lTOBLigneEcriture.PutValue('E_NUMGROUPEECR'     , FInNumGroupeEcr );

   inc(FInNumLigne);

   // gestion de la fenetre d'attente
    TT.Value     := i ;
    TT.SubText   := 'Lecture de l''écriture : ' + IntToStr( lTOBLigneReleve.GetValue('CRL_COMPTEUR') ) ;
    Application.ProcessMessages;

   end; // for

  for i := 0 to FTOBEcriture.Detail.Count - 1 do
   begin
    CAffectCompteContrePartie(FTOBEcriture.Detail[i]);
    result := CequilibrePiece( FTOBEcriture.Detail[i] );
    if not result then exit;
    CNumeroPiece(FTOBEcriture.Detail[i]);
   end; // if

 result := true;

end;

function TZTResoVersEcr.TransfertModeLibre : boolean;
var
 i                 : integer;
 lTOBLigneReleve   : TOB;
 lTOBLignePiece    : TOB;
 lTOBLigneEcriture : TZF;
begin

 result          := false;

 InitCompteur;

  for i := 0 to ( FTOBReleve.Detail.Count - 1 ) do
   begin

    lTOBLigneReleve := FTOBReleve.Detail[i];

    lTOBLignePiece  := NumeroteBordereau(lTOBLigneReleve,lTOBLignePiece);

    FInCompteur     := lTOBLigneReleve.GetValue('CRL_COMPTEUR');

    lTOBLigneEcriture := TZF.Create( 'ECRITURE' , lTOBLignePiece , -1);

    AffecteTOBLigneEcr ( lTOBLigneEcriture , lTOBLigneReleve );

    lTOBLigneEcriture.PutValue('E_MODESAISIE'       , 'LIB');
    lTOBLigneEcriture.PutValue('E_NUMGROUPEECR'     , 1);

    inc(FInNumLigne);

    // gestion de la fenetre d'attente
    if FTT.Canceled then  exit;
    TT.Value     := i ;
    TT.SubText   := 'Lecture de l''écriture : ' + IntToStr( lTOBLigneReleve.GetValue('CRL_COMPTEUR') ) ;
    Application.ProcessMessages;

   end; // for

  for i := 0 to FTOBEcriture.Detail.Count - 1 do
   begin
    CAffectCompteContrePartie(FTOBEcriture.Detail[i]);
    result := CequilibrePiece( FTOBEcriture.Detail[i] );
    if not result then exit;
    CNumeroPiece(FTOBEcriture.Detail[i]);
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
begin

 result          := false;

 lStAnneeMois    := '';

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
                                         FDtDateComptable,
                                         true
                                          );
     FListeLog.Add('Création de la pièce numéro ' + intToStr(FInNumFolio) + ' pour la période '  +
                    FormatDateTime('mmmmm yyyy', lTOBLigneReleve.GetValue('CRL_DATECOMPTABLE') ) );
     lTOBLignePiece   := TOB.Create( '' , FTOBEcriture , -1 );

    end; // if

   lTOBLigneEcriture := TZF.Create( 'ECRITURE' , lTOBLignePiece , -1);

   AffecteTOBLigneEcr ( lTOBLigneEcriture , lTOBLigneReleve );

   lTOBLigneEcriture.PutValue('E_MODESAISIE'       , '-');
   lTOBLigneEcriture.PutValue('E_NUMGROUPEECR'     , 1);

   // gestion de la fenetre d'attente
   if FTT.Canceled then exit;

   FTT.Value     := i ;
   FTT.SubText   := 'Lecture de l''écriture : ' + IntToStr( lTOBLigneReleve.GetValue('CRL_COMPTEUR') ) ;
   Application.ProcessMessages;

   inc(FInNumLigne);

  end; // for
  // a refaire
  for i := 0 to FTOBEcriture.Detail.Count - 1 do
   begin
    CAffectCompteContrePartie(FTOBEcriture.Detail[i]);
    result := CequilibrePiece( FTOBEcriture.Detail[i] );
    if not result then exit;
    CNumeroPiece(FTOBEcriture.Detail[i]);
   end; // if

  result          := true;

end;

procedure TZTresoVersEcr.TransactionSaveEcriture;
var
 i   : integer;
begin

  // numerotation et transfert vers la table ecriture pour le mode piece
  if ( FStModeSaisie[1] = '-' ) then
   begin
   if not TransfertModePiece  then
   begin
    V_PGI.IoError := oeSaisie;
    NotifyError( 'Erreur sur le transfert des enregistrements' , true );
    exit;
    end;
   end; // if

 // gestion de la fenetre d'attente
 TT.Value     := 5 ;
 TT.SubText   := cStTexteEnrEnCours;
 Application.ProcessMessages;

 FTOBEcriture.InsertDBbyNivel(false);

 // gestion de la fenetre d'attente
 TT.Value     := 75 ;
 TT.SubText   := 'Suppression des lignes de la table CRELBQE';
 Application.ProcessMessages;

 for i := 0 to ( FListeAIntegrer.Count - 1 ) do
  ExecuteSQL('delete from CRELBQE where CRL_COMPTEUR= ' + VarToStr(FListeAIntegrer.Strings[i]) );

end;

procedure TZTresoVersEcr.TransactionSaveMAJCompte;
var
 i              : integer;
begin
   // gestion de la fenetre d'attente
 TT.Value     := 50 ;
 TT.SubText   := 'Mise à jour des comptes';
 Application.ProcessMessages;

 for i := 0 to FTOBEcriture.Detail.Count - 1 do
  MajSoldesEcritureTOB( FTOBEcriture.Detail[i], false);

end;



function TZTresoVersEcr.Save : TIoErr;
var
 lBoResult : boolean;
begin

 result := oeOk;

 FListeLog.Add('Compte rendu d''intégration du ' + FormatDateTime('dddd dd mmmm yyyy "à" hh:nn',Now) );
 FListeLog.Add('');
 FListeLog.Add('Journal : ' + FStJournalContrepartie );

 lBoResult := false;

 try

 // blocage de toutes les fct sur l'appli
 if not BlocageMonoPoste ( true ) then
  begin
   result := oePointage;
   NotifyError( 'Erreur sur le blocage monoposte' , true );
   exit;
  end; // if


    // numerotation et transfert vers la table ecriture
  case FStModeSaisie[1] of
   '-' : lBoResult := true;
   'L' : lBoResult := TransfertModeLibre;
   'B' : lBoResult := TransfertModeBordereau;
  end; // case

  if not lBoResult then
   begin
    result := oeSaisie;
    NotifyError( 'Erreur sur le transfert des enregistrements' , true );
    exit;
   end; // if


  if Transactions( TransactionSaveEcriture , 1 ) <> oeOK then
   begin
    result := oeStock;
    NotifyError( 'Erreur sur l''enregistrement en base' , true );
    exit;
   end; // if


  if Transactions( TransactionSaveMAJCompte , 1 ) <> oeOK then
   begin
    result := oeReseau;
    NotifyError( 'Erreur sur la mise à jour des comptes' , true );
    exit;
   end; // if

 finally
  // pas de verif des erreurs
  {$I-}
  // FListeLog.SaveToFile('c:\rapport.log');
  {$I+}
  DeblocageMonoPoste( true );
 end; // try

 PGIInfo( cStTexteTransfertReussie , cStTexteTitreFenetre );

end;

procedure TZTresoVersEcr.NotifyError( vMessage : string ; vBoAJoutLog : boolean );
var
 lStMessage : string;
begin

 lStMessage := vMessage + #13#10 + V_PGI.LastSQLError;

 if ( V_PGI.LastSQLError <> '') or ( FTypeContexte <> TModeAuto ) then
  MessageAlerte( lStMessage );

 if vBoAJoutLog then
  FListeLog.Add( lStMessage );

end;




end.
