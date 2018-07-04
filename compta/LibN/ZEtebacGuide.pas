{***********UNITE*************************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : Classe gérant la recherche en guide pour la saisie de
Suite ........ : trésorie.
Mots clefs ... :
*****************************************************************}
unit ZEtebacGuide;

interface

uses

 Classes,
 {$IFNDEF EAGLCLIENT}
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
 {$ENDIF}

 {$IFDEF VER150}
   Variants,
 {$ENDIF}

 SysUtils,
 Dialogs,
 Windows,
 // composant AGL
 HTB97,
 HCtrls,
 // Lib
 Ent1,
 Hent1,
 HMsgBox,   // pour le PGIInfo
 formule,   // pour GFormule
 ed_tools, // pour le VideListe
 // Objet de gestion
 UCstEtebac,
 ZCompte, // pour le TZCompte
 // AGL
 LookUp,
 SaisUtil,
 ULibEcriture,
 UTOFLOOKUPTOB,
 UTOB ;


type
 {JP 19/01/07 : ajout des trois champs libEnrichis}
 TTypeRecherche  = ( TColGeneral,TColDebit,TColCredit,TColLibelle,TColTVA, TColDate,
                     TColRefInterne, TColRefExterne, TColCodeAFB , TColNaturePiece,
                     TColLibEnrichi1, TColLibEnrichi2, TColLibEnrichi3);
 TTypeRecherches = set of TTypeRecherche;

 TProcAction     = function : string of object;

 TRecRech       = record
  Value         : string;
  ProcAction    : TProcAction;
 end;

 PRecLib       = ^TRecLib ;
 TRecLib       = record
  Value         : string;
  Prio          : integer;
  PTOB          : TOB ;
 end;

 TZEtebacGuide = Class(TObjetCompta)
  private
   FTypeContexte             : TTypeContexte;
   FTOB                      : TOB;             // TOB contenant l'ensemble des guides
   FCurrentTOB               : TOB;             // pointeur sur la ligne de guide courante
   FTOBEcr                   : TOB;             // TOB genere par le moteur des guides
   FStEG_GUIDE               : string;          // code du guide courant
   FGrilleReleve             : THGrid;          // pointeur sur la grille de saisie
   FListeLibelle             : TList;           // Liste des libelles de recherches pour les guides
   FListeMontant             : TStringList;     // Liste des montant de recherche pour les guides
   FListeGeneral             : TStringList;     // Liste des comptes de recherche pour les guides
   {JP 16/01/07 : Pour limiter le décryptage des guides à certain champs : dans mon cas, le libellé :
                  cet ensemble est fixe et définit dans le create. Quand on réinitialise FTypeRecherche,
                  on le fait maintenant avec FTypeRechercheReference}
   FTypeRechercheReference   : TTypeRecherches; // Liste des recherche a effectuer fixée dans le Create
   FTypeRecherche            : TTypeRecherches; // Liste des recherche a effectuer

   FStJournalContrepartie    : string;          // code du journal de la contrepartie
   FStCompteTVA              : string;          // Compte de TVA
   FStEtablissement          : string;          // Code de l'etablissement pour les recherche des guides
   FStCompteContrepartie     : string;          // Compte de contrpartie : compte de banque du journal de tresorie
   FInNumLigne               : integer;
   FTabRech                  : array[TTypeRecherche] of TRecRech;
   FStDevise                 : string;          // code de la devise pour la recherche des guides
   FStLeNomTable             : string;          // Nom de la table -> utilisé pour cree la TOB de retour
   FStPrefixe                : string;
   FZCompte                  : TZCompte;  // objet de gestion des comptes
   FBoPasTVASaisie           : boolean;
   FStTauxTvaInitial         : string ;
   FStSolde                  : string ;
   FSilentMode               : boolean ;
   FStSQL                    : string ;

   procedure SetLeNomTable ( Value : string );
   function  GetFieldProperty ( vIndex : integer ) : variant;
   function  GetFormule( StFormule : hstring ) : variant;
   procedure FindNext;

   function  RechercheGuideParLibelle : string;
   function  RechercheGuideParDebit : string;
   function  RechercheGuideParCredit : string;
   function  RechercheGuideParImputation : string;
   function  RechercheGuideTotal : string;
   function  VerifSensDuGuide( Value : TOB ) : boolean;
   function  CalculeGuide : boolean;
   function  GenerationGuide : boolean;
   function  MoteurCalculGuide (vTOBLigneEcr : TOB ; vInNumLigne : integer) : boolean;
   function  GetRowFormule(var StFormule: hstring): Integer;
   procedure DeleteGuide(vStGUIDE: string);
   procedure RecupereValeurTOB( vTOB : TOB ; vBoInitialisation : boolean = false) ;
   function  RechercheGuideParAFB ( vStLibelle : string ) : boolean ;
   {JP 18/01/07 : Gestion des libellés enrichis}
   function  HasLibelleEnrichi(aLibelle : string) : Boolean;
   function  TestCarJocker ( Const vStTexte : string ; vStChaine : string = '' ) : boolean ;
//   function  GetLibelleCourant : string ;

   property StGUIDE                      : variant index 1  read GetFieldProperty;
   property StGENERAL                    : variant index 3  read GetFieldProperty;
   property StREFINTERNE                 : variant index 4  read GetFieldProperty;
   property StLIBELLE                    : variant index 6  read GetFieldProperty;
   property StDEBIT                      : variant index 7  read GetFieldProperty;
   property StCREDIT                     : variant index 8  read GetFieldProperty;
   property StARRET                      : variant index 9  read GetFieldProperty;
 //  property StLIBELLEGUIDE               : variant index 10  read GetFieldProperty;
   property StAUXILIAIRE                 : variant index 11 read GetFieldProperty;
   property StETABLISSGUID               : variant index 12 read GetFieldProperty;

  protected
   procedure NotifyError(vMessage, vDelphiMessage, vFunctionName: string); override ;

  public

   constructor Create( vInfoEcr : TInfoEcriture ) ; overload; override ;
   constructor CreateTTR( vInfoEcr : TInfoEcriture; aTypRech : TTypeRecherches); overload;
   destructor Destroy; override;

   function  RecalculGuide( vTOBEcr : TOB ; vInNumLigne : integer) : TOB;
   function  RechercheGuideEnBase( vTOB,vTOBResult : TOB ; G : THGrid = nil ) : boolean;
   function  RechercheDernierGuide(vTOB,vTOBResult : TOB ; G : THGrid = nil ) : boolean;
   function  IsColStop(ARow,ACol : integer) : boolean;
   function  FirstGuide( Value : string) : integer;
   function  PossedeArret : boolean;
   procedure ClearGuide;

   function  Load : boolean;
   function  EstVide : boolean;


   property StJournalContrepartie        : string           read FStJournalContrepartie     write FStJournalContrepartie;
   property StEtablissement              : string           read FStEtablissement           write FStEtablissement;
   property StCompteContrepartie         : string           read FStCompteContrepartie      write FStCompteContrepartie;
   property StCompteTVA                  : string           read FStCompteTVA               write FStCompteTVA;
   property ZCompte                      : TZCompte         read FZCompte                   write FZCompte;
   property StDevise                     : string           read FStDevise                  write FStDevise;
//   property StLibelleCourant             : string           read GetLibelleCourant ;
   property StLeNomTable                 : string           read FStLeNomTable              write SetLeNomTable;
   property SilentMode                   : boolean          read FSilentMode                write FSilentMode default false ;
 end;


function CCreerCommeGuide( vTOB : TOB ; vStPref : string = 'E_' ) : string ; 

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  GuidUtil;

const
 cStIdCodeAFB     = '$';
 cStIdCodeAFBSupp = '&' ;
 cStJocker        = '*' ;
 cStLibEnrichi    = '#';

constructor TZEtebacGuide.Create( vInfoEcr : TInfoEcriture ) ;
begin
 inherited ;
 FListeLibelle                        := TList.Create;
 FListeMontant                        := TStringList.Create;
 FListeGeneral                        := TStringList.Create;
 FTypeRechercheReference              := [TColGeneral,TColDebit,TColCredit,TColLibelle];
 FTypeRecherche                       := FTypeRechercheReference;

 FTabRech[TColDebit].ProcAction       := RechercheGuideParDebit;
 FTabRech[TColCredit].ProcAction      := RechercheGuideParCredit;
 FTabRech[TColLibelle].ProcAction     := RechercheGuideParLibelle;
 FTabRech[TColGeneral].ProcAction     := RechercheGuideParImputation;

 FTypeContexte                        := TModeGuide;
 FStSQL                               := '' ;

end;

constructor TZEtebacGuide.CreateTTR(vInfoEcr : TInfoEcriture; aTypRech : TTypeRecherches);
begin

  inherited Create(vInfoEcr);
  FListeLibelle                        := TList.Create;
  FListeMontant                        := TStringList.Create;
  FListeGeneral                        := TStringList.Create;
  FTypeRechercheReference              := aTypRech;
  FTypeRecherche                       := FTypeRechercheReference;

  if TColDebit in FTypeRechercheReference then
    FTabRech[TColDebit].ProcAction   := RechercheGuideParDebit;
  if TColCredit in FTypeRechercheReference then
    FTabRech[TColCredit].ProcAction  := RechercheGuideParCredit;
  if TColLibelle in FTypeRechercheReference then
    FTabRech[TColLibelle].ProcAction := RechercheGuideParLibelle;
  if TColGeneral in FTypeRechercheReference then
    FTabRech[TColGeneral].ProcAction := RechercheGuideParImputation;

  FTypeContexte                        := TModeGuide;
end;

destructor TZEtebacGuide.Destroy;
var
 i : integer ;
begin

  for i := 0 to FListeLibelle.Count - 1 do
   if FListeLibelle.items[i] <> nil then dispose(FListeLibelle.items[i]) ;

  if assigned(FTOB)           then FTOB.Free;
  if assigned(FListeLibelle)  then FListeLibelle.Free;
  if assigned(FListeMontant)  then FListeMontant.Free;
  if assigned(FListeGeneral)  then FListeGeneral.Free;

  FTOB                        := nil;
  FListeMontant               := nil;
  FListeLibelle               := nil;
  FListeGeneral               := nil;

  inherited destroy;

end;

function TZEtebacGuide.GetFieldProperty ( vIndex : integer ) : variant;
begin

 if not assigned(FCurrentTOB) then
  begin
   result := '';
   exit;
  end;

 case vIndex of
  0 : result := FCurrentTOB.GetValue('EG_TYPE');
  1 : result := FCurrentTOB.GetValue('EG_GUIDE');
  2 : result := FCurrentTOB.GetValue('EG_NUMLIGNE');
  3 : result := FCurrentTOB.GetValue('EG_GENERAL');
  4 : result := FCurrentTOB.GetValue('EG_REFINTERNE');
  5 : result := FCurrentTOB.GetValue('EG_REFEXTERNE');
  6 : result := FCurrentTOB.GetValue('EG_LIBELLE');
  7 : result := FCurrentTOB.GetValue('EG_DEBITDEV');
  8 : result := FCurrentTOB.GetValue('EG_CREDITDEV');
  9 : result := FCurrentTOB.GetValue('EG_ARRET');
 10 : result := FCurrentTOB.GetValue('GU_LIBELLE');
 11 : result := FCurrentTOB.GetValue('EG_AUXILIAIRE');
 12 : result := FCurrentTOB.GetValue('GU_ETABLISSEMENT');
 end; // case

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Suppression du guide de la tob memoire du guide
Suite ........ : vStGUIDE
Mots clefs ... :
*****************************************************************}
procedure TZEtebacGuide.DeleteGuide( vStGUIDE : string);
var
 lTOBLigneGuide : TOB;
begin

 lTOBLigneGuide := FTOB.FindFirst(['EG_GUIDE'],[vStGUIDE],false);

 while assigned ( lTOBLigneGuide ) do
  begin
   lTOBLigneGuide.free;
   lTOBLigneGuide := FTOB.FindNext(['EG_GUIDE'],[vStGUIDE],false);
  end; // while

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 31/08/2001
Modifié le ... : 11/09/2007
Description .. : -Chargement des guides au démarrage.
Suite ........ : - Suppression des guides comportant un point d'arrêt sur la
Suite ........ : dernière ligne
Suite ........ : - Calcul du sens des guides et affectation du solde sur
Suite ........ : toutes les lignes
Suite ........ : - Stockage des montants, des libelles de recherche,
Suite ........ : 
Suite ........ : -24/05/2002 - ajout du flag indiquant un guide pour la saisie
Suite ........ : de tresorerie
Suite ........ : - LG - 17/07/2007 - FB 20561 - gestion des etablissement 
Suite ........ : ds les guides
Suite ........ : - LG 11/09/2007 - Fb 21408 - on affiche plus le msg si les 
Suite ........ : crites n'ont pas chanegr
Mots clefs ... : 
*****************************************************************}
function TZEtebacGuide.Load : boolean;
var
 Q           : TQuery;
 lStCrit     : string;
 lIndex      : integer;
 i,j         : integer;
 lTOB        : TOB;
 lStValeur   : string;
 lStSolde    : string;
 lPRecLib    : PRecLib ;
 lp1,lp2     : integer ;
 lStSQL      : string ;
begin

 result := false;

 if ( FStJournalContrepartie = '' ) or ( FStSQL = ( FStJournalContrepartie + FStEtablissement ) ) then exit ;

 // creation de la TOB contenant toutes les guides FTOB
 // TOB base sur la table ECRGUI. struture :
 // -> FTOB
 //    -> TOBligne
 //    -> TOBLigne
 // toutes les lignes de guides sont au meme niveau. On utilise des findfirst pour trouvé un guide et des findnext pour passer
 // a la ligne suivante.
 if not assigned(FTOB) then
  FTOB := TOB.Create('ECRGUI', nil , -1)
   else
    ClearGuide;
 Q := nil;

 try

 lStSQL := 'select EG_GUIDE , '                                        +
              'EG_NUMLIGNE, '                                          +
              'EG_GENERAL, '                                           +
              'EG_REFINTERNE, '                                        +
              'EG_REFEXTERNE, '                                        +
              'EG_LIBELLE, '                                           +
              'EG_DEBITDEV, '                                          +
              'EG_CREDITDEV, '                                         +
              'EG_ARRET, '                                             +
              'GU_LIBELLE,'                                            +
              'GU_ETABLISSEMENT, '                                      +
              'EG_AUXILIAIRE '                                         +
              'from ECRGUI, GUIDE '                                    +
              'where EG_TYPE = "NOR" '                                 +
              ' and GU_JOURNAL = "' + FStJournalContrepartie + '" '    +
              ' and GU_GUIDE = EG_GUIDE '                              +
              ' and GU_TYPE = EG_TYPE '                                +
              ' and GU_DEVISE = "' + FStDevise + '" '                  +
              ' and GU_TRESORERIE= "X" ' ;

 if FStEtablissement <> '' then
  lStSQL := lStSQL + ' and (GU_ETABLISSEMENT = "' + FStEtablissement  + '" '  +
                     ' OR GU_ETABLISSEMENT = "" OR GU_ETABLISSEMENT IS NULL) ' ;

 lStSQL := lStSQL + 'order by EG_GUIDE , EG_NUMLIGNE' ;


 Q := OpenSQL(lStSQL , true ) ;
 // on charge l'ensemble des guides
 FTOB.LoadDetailDB('ECRGUI','','',Q,true);

 if ( FTOB.Detail.Count = 0 ) then
  begin
   FCurrentTOB := nil;
   if not FSilentMode and ( FStSQL <> ( FStJournalContrepartie + FStEtablissement ) ) then  // FB 21408 
    PGIInfo( cStTexteAucunGuide , cStTexteTitreFenetre);
   FStSQL := FStJournalContrepartie + FStEtablissement ;
   exit;
  end; // if

 FCurrentTOB := FTOB.Detail[0];
 // on stocke pour chaque ligne le sens du guide et le solde total du guide
 // pour les verif de validité du guide
 FCurrentTOB.AddChampSup('SENS'   , true);
 FCurrentTOB.AddChampSup('SOLDE'  , true);
 FCurrentTOB.AddChampSup('RECH'   , true);

 // 1 -  suppression des lignes comportant un point d'arrêt sur la dernière ligne
 // 2 - calcul du sens des guides et affectation du solde sur toutes les lignes
 // 3-  stockage des montants de recherche

 // les dernières lignes sont forcement sur le compte de contrepartie
 FCurrentTOB  := FTOB.FindFirst(['EG_GENERAL'],[StCompteContrepartie],false);

 while Assigned(FCurrentTOB) do
  begin

   // Suppression des lignes ayants des points d'arret sur la dernire ligne
   if Pos( 'X', StARRET ) > 0 then
    DeleteGuide(StGUIDE)
     else
      begin

       if ( StDEBIT <> '' ) then
        begin

         if ( isNumeric(StDEBIT) )  then
          FListeMontant.AddObject(VarToStr(StDEBIT),TObject(FCurrentTOB));

         FCurrentTOB.PutValue('SENS','0');
         FCurrentTOB.PutValue('SOLDE',StDEBIT);

        end; // if

       if ( StCREDIT <> '' ) then
        begin

         if ( isNumeric(StCREDIT) ) then
          FListeMontant.AddObject(VarToStr(StCREDIT),TObject(FCurrentTOB));

         FCurrentTOB.PutValue('SENS','1');
         FCurrentTOB.PutValue('SOLDE',StCREDIT);

        end; // if
      end; // else

   FCurrentTOB := FTOB.FindNext(['EG_GENERAL'],[StCompteContrepartie],false);

 end; // while

 // recontrole de la presence de guide ( il avait peut etre tous des point d'arret sur la der ligne
 if FTOB.Detail.Count = 0 then
  begin
   FCurrentTOB := nil;
   if not FSilentMode then
    PGIInfo( cStTexteAucunGuide , cStTexteTitreFenetre);
   exit;
  end; // if


 // affectation du sens et du solde à toutes les lignes
 // on part du premier element
 lIndex := 0;

 // calcul du sens du guide
 for i := 0  to ( FTOB.Detail.Count - 1 ) do
  begin

    lTOB       := FTOB.Detail[i];
    lStValeur  := lTOB.GetValue('SENS');
    lStSolde   := lTOB.GetValue('SOLDE');
    // est ce que la valeur est définie
    if lStValeur <> #0 then
     begin
      // si oui on remonte et on remet à jour les lignes
      for j := lIndex to i do
       begin
        lTOB := FTOB.Detail[j];
        lTOB.PutValue('SENS',lStValeur);
        lTOB.PutValue('SOLDE',lStSolde);
       end; // for
      // on stocke la derniere ligne mise à jour
      lIndex := i + 1;

     end; // if

  end; // for

 FCurrentTOB  := FTOB.FindFirst(['EG_NUMLIGNE'],[1],false); // on se place sur le première ligne des filtres

 while Assigned(FCurrentTOB) do
  begin

   // stockage des libelles de recherche
   if ( StLIBELLE <> '' )  then
    begin

     lStCrit         := copy( StLIBELLE,1,pos(';',StLIBELLE ) - 1  );
     New(lPRecLib) ;
     lPRecLib^.Value := '' ;
     lPRecLib^.Prio  := 99 ;
     lPRecLib^.PTOB  := FCurrentTOB ;

     {JP 19/01/07 : gestion des libellés enrichis}
     if (lStCrit = '') or (Pos(cStLibEnrichi, lStCrit) > 0) then
      lPRecLib^.Value := UpperCase(StLIBELLE)
       else
        lPRecLib^.Value := UpperCase(copy( StLIBELLE,1,pos(';',StLIBELLE ) - 1  )) ;

     lP1 :=  pos(cStIdCodeAFB,lPRecLib^.Value) ;
     lP2 :=  pos(cStIdCodeAFBSupp,lPRecLib^.Value) ;

     if ( lP1 > 0)  and ( lP2  > 0)  then lPRecLib^.Prio  := 0 ;  // $ + &
     if ( lP1 = 0)  and ( lP2  = 0 ) then lPRecLib^.Prio  := 1 ;  // libelle sans code afb et caractere de recherche
     if ( lP1 > 0 ) and ( lP2 = 0 ) and ( Pos(cStJocker,lPRecLib^.Value) = 0 ) then lPRecLib^.Prio  := 2 ;  // $
     if ( lP1 > 0 ) and ( lP2 = 0 ) and ( Pos(cStJocker,lPRecLib^.Value) > 0 ) then lPRecLib^.Prio  := 3 ;  // $ + *

     FListeLibelle.Add(lPRecLib) ;

    end; // if

   // stockage des comptes
  // if StGENERAL <> '' then
    FListeGeneral.AddObject(VarToStr(StGENERAL),TObject(FCurrentTOB));

   FCurrentTOB := FTOB.FindNext(['EG_NUMLIGNE'],[1],false);

  end; // while

  result := true;

 finally
  {$IFDEF TEST}
   FTOB.SaveToFile('c:\FTOBGuide.txt',false,true,true);
 {$ENDIF}

  Ferme(Q);
 end; // try

end;

procedure TZEtebacGuide.FindNext;
begin

 if not assigned(FTOB) then exit;

 FCurrentTOB := FTOB.FindNext(['EG_GUIDE'], [FStEG_GUIDE],false);

 end;


function TZEtebacGuide.FirstGuide( Value : string) : integer;
begin

 result       := -1;

 // l'imputation doit être trouvé sur la première ligne
 FCurrentTOB  := FTOB.FindFirst(['EG_NUMLIGNE','EG_GUIDE'],[1,Value],false);
 FStEG_GUIDE  := StGUIDE;

 if FCurrentTOB <> nil then
  result := FTOB.GetIndex;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Retourne la ligne d'une formule
Suite ........ : ex : [E_LIBELLE:L1] -> 1
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.GetRowFormule ( var StFormule : hstring ) : Integer ;
var
 Pos1,
 Pos2 : Integer ;
begin

 Result := 0;

 Pos2 := Pos(':L', StFormule);
 if Pos2 > 0 then
  System.Delete( StFormule, Pos2+1 , 1 );

 Pos1 := Pos(':',  StFormule);
 if ( Pos1 <= 0 ) or ( Pos1 = Length(StFormule)) then
  Exit ;

 Result := Round(Valeur(Copy(StFormule, Pos1+1, 5))) ;

 System.Delete(StFormule, Pos1, 5) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 09/11/2001
Description .. : Retoune la valeur la valeur du champs correspondant à la
Suite ........ : formule du guide.
Suite ........ : Les champs en pris compte sont :
Suite ........ : SOLDE,E_DEBIT,E_CREDIT,SOLDE,E_LIBELLE,
Suite ........ : E_REFINTERNE,E_REFEXTERNE,E_DATECOMPTABLE,
Suite ........ : E_TVA
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.GetFormule( StFormule : hstring) : Variant ;
var
 lValue : variant;
 lInRow : integer; // ligne sur laquel s'applique la formule
 lTOB   : TOB;
begin

 result := #0;
 lValue := #0;

 StFormule := UpperCase(Trim(StFormule));

 if StFormule = '' then Exit ;

 // on recupere la ligne d'application de la formule
 lInRow := GetRowFormule(StFormule);
 if ( lInRow = 0 ) or ( Abs(lInRow) > ( FTOBEcr.Detail.Count ) ) or ( (FInNumLigne - 1 + lInRow) < 0) then // le n° de ligne de la formule ne peut être > n° ligne courant
  lInRow := 1;

 if StFormule = 'SOLDE' then
  begin
    lValue := FStSolde ;
  end
   else
    if StFormule = 'E_DEBIT' then
     begin
      if ( lInRow < 0 ) then  // on remonte sur la ligne precedente
       lTOB   := FTOBEcr.Detail[ FInNumLigne - 1 + lInRow]
        else
         lTOB   := FTOBEcr.Detail[lInRow - 1];
      lValue := VarToStr( lTOB.GetValue( FStPrefixe + '_DEBIT') );
     end
      else
       if StFormule = 'E_CREDIT' then
        begin
         if ( lInRow < 0 ) then  // on remonte sur la ligne precedente
          lTOB   := FTOBEcr.Detail[ FInNumLigne - 1 + lInRow]
           else
            lTOB   := FTOBEcr.Detail[lInRow - 1];
         lValue := VarToStr( lTOB.GetValue( FStPrefixe + '_CREDIT') );
        end
   else
    if StFormule = 'E_TVA' then
      lValue := FTabRech[TColTVA].Value
    else if StFormule = 'E_DATECOMPTABLE' then
      lValue := FTabRech[TColDate].Value
    else if StFormule = 'E_REFINTERNE' then
      lValue := FTabRech[TColRefInterne].Value
    else if StFormule = 'E_REFEXTERNE' then
      lValue := FTabRech[TColRefExterne].Value
      
    {JP 02/03/07 : Gestion des libellés enrichis}
    else if StFormule = 'E_LIBRETEXTE1' then
      lValue := FTabRech[TColLibEnrichi1].Value
    else if StFormule = 'E_LIBRETEXTE2' then
      lValue := FTabRech[TColLibEnrichi2].Value
    else if StFormule = 'E_LIBRETEXTE3' then
      lValue := FTabRech[TColLibEnrichi3].Value

    else if StFormule = 'E_LIBELLE' then begin
      if ( lInRow = 1 ) then
        lValue := FTabRech[TColLibelle].Value
      else begin
        if ( lInRow < 0 ) then  // on remonte sur la ligne precedente
          lTOB   := FTOBEcr.Detail[ FInNumLigne - 1 + lInRow]
        else
          lTOB   := FTOBEcr.Detail[lInRow - 1];
        lValue := VarToStr( lTOB.GetValue( FStPrefixe + '_LIBELLE') );
      end;
    end; // libelle

 result := lValue;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Calcule du guide pour la ligne vTOBLigneEcr
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.MoteurCalculGuide ( vTOBLigneEcr : TOB ; vInNumLigne : integer) : boolean;
var
 lStNumeroPiece     : string;
 lStLibelle         : string;
 lRdDebit           : double;
 lRdCredit          : double;
 lStCritRech        : string;
 lStCritLibelle     : string;
 lStValeurCell      : string;
 lStLibelleGuide    : string;
 lStARRET           : string;
 lRdTaux            : double ;
 lStNatureAuxi      : string ;
 {JP 19/01/07 : Gestion des libellés enrichis}
 lStLibelleEnr      : string;


   {LGE- 23/06/2004 - FB 13668 - correction de la nature ds la fct RecupInfoTVA}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 23/08/2007
Modifié le ... :   /  /    
Description .. : - LG - 23/08/2007 - on stocke le taux de tva qd on l'a 
Suite ........ : calculer
Mots clefs ... : 
*****************************************************************}
   function IsNumericLG ( Value : string ) : boolean;
   begin
     result := ( pos('-', Copy( Value, 2 , length(Value) ) ) = 0 ) and ( pos('/',Value) = 0 ) and ( pos('*',Value) = 0 ) and ( pos('+',Value) = 0 );
   end;

begin

 result                  := false;
 FTabRech[TColTVA].Value := FStTauxTvaInitial ;

 try

  if vTOBLigneEcr.GetValue( FStPrefixe +  '_GENERAL') <> ''  then
   FTabRech[TColGeneral].Value := vTOBLigneEcr.GetValue( FStPrefixe +  '_GENERAL')
    else
     FTabRech[TColGeneral].Value := StGeneral ;

  // calcul de la TVA
  if ZCompte.IsTvaAutorise(FTabRech[TColGeneral].Value) then
   begin
    if FBoPasTVASaisie then
     begin
      // chargement des info necessaire a la recuperation de la tva
      Info.LoadCompte(FTabRech[TColGeneral].Value) ;
      Info.LoadJournal(StJournalContrepartie) ;
      lStNatureAuxi := '' ;
      if Info.LoadAux(vTOBLigneEcr.GetValue( FStPrefixe +  '_AUXILIAIRE') ) then
       lStNatureAuxi := Info.Aux.GetValue('T_NATUREAUXI') ;

      ZCompte.RecupInfoTVA(FTabRech[TColGeneral].Value,
                           Valeur(FTabRech[TColDebit].Value) ,
                           FTabRech[TColNaturePiece].Value,
                           Info.Journal.GetValue('J_NATUREJAL') ,
                           lStNatureAuxi ,
                           FStCompteTVA,
                           lRdTaux  );
      vTOBLigneEcr.PutValue( FStPrefixe + '_TAUXTVA',lRdTaux ) ;
      FTabRech[TColTVA].Value:=FloatToStr(vTOBLigneEcr.GetValue( FStPrefixe + '_TAUXTVA')) ;
      FStTauxTvaInitial      := FTabRech[TColTVA].Value ;  // FB 20055
     end // on n'a pas saisie de TVA
      else
       begin // on a saisie un taux de tva et la tva et autorise pour ce compte
        vTOBLigneEcr.PutValue( FStPrefixe + '_TAUXTVA' ,Valeur(FTabRech[TColTVA].Value)) ;
       end
   end ; // la tva est autorise pour ce compte

   // on stocke le libelle du guide dans une string pour pouvoir le modifier
   // ( supprimer les caractères de recherche )
   lStLibelleGuide := StLIBELLE;
   lStLibelle      := '';

   // gestion des libelles dans le guide
   lStCritRech     := UpperCase(copy( StLIBELLE,1,pos(';',StLIBELLE ) - 1  ));
   lStCritLibelle  := trim(copy( StLIBELLE,pos(';',StLIBELLE) + 1 , length(StLIBELLE) )) ;

   // suppression des caractères de recherche dans la chaine de recherche
   if Copy( lStCritLibelle, 1 , 1) = '*' then
    lStCritLibelle := Copy(lStCritLibelle, 2, length( lStCritLibelle ));

   if Copy( lStCritLibelle , length(lStCritLibelle) , 1) = '*' then
    lStCritLibelle := Copy(lStCritLibelle, 1, length( lStCritLibelle ) - 1);

   if Copy( lStLibelleGuide, 1 , 1) = '*' then
    lStLibelleGuide := Copy(lStLibelleGuide, 2, length( lStLibelleGuide ));

   if Copy( lStLibelleGuide , length(lStLibelleGuide) , 1) = '*' then
     lStLibelleGuide := Copy(lStLibelleGuide, 1, length( lStLibelleGuide ) - 1);


   if lStCritRech <> ''  then
   begin
     {JP 19/01/07 : Gestion du libellé enrichi dans la recherche du guide}
     if (Pos(cStLibEnrichi, lStCritRech) > 0) then begin
            if (Pos(cStLibEnrichi + '1', lStCritRech) > 0) then lStLibelleEnr := FTabRech[TColLibEnrichi1].Value
       else if (Pos(cStLibEnrichi + '2', lStCritRech) > 0) then lStLibelleEnr := FTabRech[TColLibEnrichi2].Value
       else if (Pos(cStLibEnrichi + '3', lStCritRech) > 0) then lStLibelleEnr := FTabRech[TColLibEnrichi3].Value;
       {Si par hasard, la formule de lStCritLibelle renvoie vide, on met le contenu du libellé pour
        éviter de voir apparaître dans le libellé de la pièce quelque chose de la forme #2*GGGG*}
       lStCritRech := FTabRech[TColLibelle].Value + '(Def)';
     end;

     if ( lStCritLibelle = '>' ) then
     begin
       if ( length(trim(FTabRech[TColLibelle].Value)) <= length(lStCritRech) ) then
         lStLibelle := FTabRech[TColLibelle].Value
       else
         lStLibelle := trim ( copy(
                                   FTabRech[TColLibelle].Value,
                                   pos( lStCritRech , UpperCase(FTabRech[TColLibelle].Value )) + length(lStCritRech) ,
                                   length(FTabRech[TColLibelle].Value )
                             ));
     end

     else if lStCritLibelle <> '' then
     begin
       lStValeurCell := trim(GFormule(lStCritLibelle, GetFormule, nil, 1));
       if ( lStValeurCell <> '' ) then
         lStLibelle  := lStValeurCell
       else
         lStLibelle  := lStCritRech;
     end; //if
   end // critere de recherche

   else if trim(lStLibelleGuide) <> '' then
   begin
     lStValeurCell           := GFormule(lStLibelleGuide, GetFormule, nil, 1);
     if ( lStValeurCell <> '' ) and ( pos(cStIdCodeAFB,lStValeurCell) > 0 )  then
       lStLibelle             := FTabRech[TColLibelle].Value
     else
     if ( lStValeurCell <> '' ) and ( pos(cStIdCodeAFB,lStValeurCell) = 0 ) then
       lStLibelle          := trim(lStValeurCell)
     else
       lStLibelle        := lStValeurCell; //trim(lStLibelleGuide);
   end; //if


   if StDEBIT <> '' then
    begin
     lStValeurCell  := GFormule(StDEBIT, GetFormule, nil, 1);
     if IsNumeric(lStValeurCell) and not IsNumericLG(lStValeurCell) then
      begin
       NotifyError( TraduireMemoire('Erreur de formule dans le Guide n°')          + FStEG_GUIDE
                    + ' ' + RechDom('TTGUIDEECR', FStEG_GUIDE ,false)              + #13#10
                    + TraduireMemoire('La formule ') + StDEBIT + TraduireMemoire(' est incorrecte') ,
                    '',
                    '' );
       exit;
      end; // if

     lRdDebit       := Valeur(lStValeurCell) ;
    end //if
     else
      if vTOBLigneEcr.GetValue( FStPrefixe +  '_DEBIT') <> 0 then
       lRdDebit                          := vTOBLigneEcr.GetValue( FStPrefixe +  '_DEBIT')
        else
         lRdDebit                        := 0;

   if StCREDIT<> '' then
    begin
     lStValeurCell  := GFormule(StCREDIT, GetFormule, nil, 1);
     if IsNumeric(lStValeurCell) and not IsNumericLG(lStValeurCell) then
      begin
       NotifyError( TraduireMemoire('Erreur de formule dans le Guide n°')          + FStEG_GUIDE
                    + ' ' + RechDom('TTGUIDEECR', FStEG_GUIDE ,false)  + #13#10
                    + TraduireMemoire('La formule ') + StCREDIT + TraduireMemoire(' est incorrecte') ,
                    '',
                    '' );
       exit;
      end; // if
      lRdCredit     := Valeur(lStValeurCell);
    end //if
     else
      if vTOBLigneEcr.GetValue( FStPrefixe + '_CREDIT') <> 0 then
       lRdCredit                         := vTOBLigneEcr.GetValue( FStPrefixe +  '_CREDIT')
        else
         lRdCredit                       := 0;

   // gestion des montants negatifs
   if ( lRdDebit < 0 ) and not( VH^.MontantNegatif ) then
    begin
     lRdCredit := lRdDebit * (-1);
     lRdDebit  := 0;
    end // lRdDebit > 0
     else
      if ( lRdCredit < 0 ) then
       begin
         lRdDebit  := lRdCredit * (-1);
         lRdCredit := 0;
       end;

   if StREFINTERNE <> '' then
    lStNumeroPiece                       := GFormule(StREFINTERNE, GetFormule, nil, 1)
     else
      lStNumeroPiece                     := '';

   lStARRET := StARRET;

   vTOBLigneEcr.PutValue('GUIDE'        , StGUIDE);
   vTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , vInNumLigne);
   if ( StGENERAL <> '' ) and (vTOBLigneEcr.GetValue( FStPrefixe + '_GENERAL')= '' )  then
    vTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'     , StGENERAL );
   if ( StAUXILIAIRE <> '' ) then
    vTOBLigneEcr.PutValue( FStPrefixe + '_AUXILIAIRE'  , StAUXILIAIRE );
   if lStNumeroPiece <> '' then
    vTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , lStNumeroPiece);
   if lStLibelle <> '' then
    vTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , lStLibelle);
   vTOBLigneEcr.PutValue( FStPrefixe + '_DEBIT'        , lRdDebit);
   vTOBLigneEcr.PutValue( FStPrefixe + '_CREDIT'       , lRdCredit);
   vTOBLigneEcr.PutValue( FStPrefixe + '_ETABLISSEMENT' , StETABLISSGUID);
   {JP 19/01/07 : Gestion du libellé enrichi}
   if lStLibelleEnr <> '' then
     vTOBLigneEcr.PutValue( FStPrefixe + '_REFEXTERNE', lStLibelleEnr);

   result          := true;

  except
   on E:Exception do
    begin
      NotifyError( TraduireMemoire('Erreur de formule dans le Guide n°')          + FStEG_GUIDE
                    + ' ' + RechDom('TTGUIDEECR', FStEG_GUIDE ,false)  + #13#10 + E.message , '' , '' );
    end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 09/11/2001
Description .. : Retoune une TOB de type FStLeNomTable calcule pour le
Suite ........ : guide FStEG_GUIDE.
Suite ........ : Attention : FStLeNomTable doit avoir des champs suivant
Suite ........ : _GENERAL,_REFINTERNE,_LIBELLE,_DEBIT,_CREDIT,
Suite ........ : _NUMLIGNE
Suite ........ :
Mots clefs ... :
*****************************************************************}
function  TZEtebacGuide.CalculeGuide : boolean;
var
 lTOBLigneEcr : TOB;
begin

  result       := false;

  try

  FInNumLigne  := 1;
  // on se place sur la première ligne du guide
  FCurrentTOB  := FTOB.FindFirst(['EG_NUMLIGNE','EG_GUIDE'],[1,FStEG_GUIDE],false);

  while assigned(FCurrentTOB) do
  begin

    lTOBLigneEcr  := TOB.Create( FStLeNomTable , FTOBEcr , -1);
    lTOBLigneEcr.AddChampSup( 'GUIDE',true );

    if not MoteurCalculGuide( lTOBLigneEcr , FInNumLigne ) then
     begin
      result  := false;
      FTOBEcr := nil;
      exit;
     end;

    FindNext;

    Inc(FInNumLigne);

  end; // while

  result := true;
  //TobDebug(FTOBEcr);
 except
  on E : Exception do
   begin
    MessageAlerte('Erreur lors de l''affectation des guides '+ #10#13 + E.message);
   end;
 end; // if

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Recalcul de la TOB ecriture vTOBEcr a partir de la ligne
Suite ........ : vInNumLigne
Mots clefs ... :
*****************************************************************}
function  TZEtebacGuide.RecalculGuide ( vTOBEcr : TOB ; vInNumLigne : integer)   : TOB;
var
 lTOBLigneEcr : TOB;
begin

  FInNumLigne := vInNumLigne;

  FTOBEcr     := vTOBEcr;
  RecupereValeurTOB(vTOBEcr.Detail[vInNumLigne-1]);

  // on se replace sur la ligne courante
  FCurrentTOB := FTOB.FindFirst(['EG_NUMLIGNE','EG_GUIDE'],[FInNumLigne,FStEG_GUIDE],false);

  while assigned(FCurrentTOB) do
   begin

    lTOBLigneEcr := vTOBEcr.Detail[FInNumLigne - 1];

    MoteurCalculGuide(lTOBLigneEcr,FInNumLigne);

     Inc(FInNumLigne);
   // FindNext;
    FCurrentTOB := FTOB.FindFirst(['EG_NUMLIGNE','EG_GUIDE'],[FInNumLigne,FStEG_GUIDE],false);

  end; // while

  result := vTOBEcr;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 18/10/2007
Description .. : Retourne une tob ecriture genere ne fonction du compte
Suite ........ : d'imputation :
Suite ........ : 1 ligne : compte d'imput
Suite ........ : 2 ligne : compte de tva du compte
Suite ........ : 3 ligne : compte de contrepartie
Suite ........ : - LG - 18/10/2007 - on perdais l'etablissement en 
Suite ........ : provenance de l'ecriture
Mots clefs ... : 
*****************************************************************}
function TZEtebacGuide.GenerationGuide : boolean;
var
 lTOBLigneEcr   : TOB;
 lRdTva         : double;
 lRdVal         : double;
begin

//  result       := false;

  // 1 ligne : compte d'imputation saisie
  lTOBLigneEcr := TOB.Create( FStLeNomTable , FTOBEcr , -1 );
  lTOBLigneEcr.AddChampSup( 'GUIDE',true );
  lTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , 1);
  lTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'      , FTabRech[TColGeneral].Value);
  lTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , FTabRech[TColRefInterne].Value);
  lTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , FTabRech[TColLibelle].Value);
  lTOBLigneEcr.PutValue( FStPrefixe + '_DATECOMPTABLE' ,StrToDateTime( FTabRech[TColDate].Value));
  lTOBLigneEcr.PutValue( FStPrefixe + '_ETABLISSEMENT' , FStEtablissement);

  if ( ZCompte.IsTvaAutorise( FTabRech[TColGeneral].Value , 0 )) and
     ( FTabRech[TColTva].Value <> '') and ( FTabRech[TColTva].Value <> '0') then
       begin
        // 2 ligne : compte de tva
        lTOBLigneEcr := TOB.Create( FStLeNomTable , FTOBEcr , -1);
        lTOBLigneEcr.AddChampSup( 'GUIDE',true );
        lTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , 2);
        lTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'      , StCompteTVA);
        lTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , FTabRech[TColRefInterne].Value);
        lTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , FTabRech[TColLibelle].Value);
        lTOBLigneEcr.PutValue( FStPrefixe + '_DATECOMPTABLE' , StrToDateTime(FTabRech[TColDate].Value));
        lTOBLigneEcr.PutValue( FStPrefixe + '_ETABLISSEMENT' , FStEtablissement);

        // 3 ligne : compte de contrepartie
        lTOBLigneEcr := TOB.Create( FStLeNomTable , FTOBEcr , -1);
        lTOBLigneEcr.AddChampSup( 'GUIDE',true );
        lTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , 3);
        lTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'      , StCompteContrepartie);
        lTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , FTabRech[TColRefInterne].Value);
        lTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , FTabRech[TColLibelle].Value);
        lTOBLigneEcr.PutValue( FStPrefixe + '_DATECOMPTABLE' , StrToDateTime(FTabRech[TColDate].Value));
        lTOBLigneEcr.PutValue( FStPrefixe + '_ETABLISSEMENT' , FStEtablissement);

       end
        else
         begin
          // 2 ligne : compte de contrepartie
          lTOBLigneEcr := TOB.Create( FStLeNomTable , FTOBEcr ,-1);
          lTOBLigneEcr.AddChampSup( 'GUIDE',true );
          lTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , 2);
          lTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'      , StCompteContrepartie);
          lTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , FTabRech[TColRefInterne].Value);
          lTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , FTabRech[TColLibelle].Value);
          lTOBLigneEcr.PutValue( FStPrefixe + '_DATECOMPTABLE' ,StrToDateTime(FTabRech[TColDate].Value));
          lTOBLigneEcr.PutValue( FStPrefixe + '_ETABLISSEMENT' , FStEtablissement);
          
         end;

      // gestion des colonnes debit / credit
      if ZCompte.IsTvaAutorise(FTabRech[TColGeneral].Value, 0)
         and ( FTabRech[TColTva].Value <> '')
         and ( FTabRech[TColTva].Value <> '0') then
       begin
        if FTabRech[TColDebit].Value <> '' then
         begin

          lRdVal := Valeur(FTabRech[TColDebit].Value) / ( 1 + ( Valeur(FTabRech[TColTva].Value) / 100 ) ) ;
          lRdTva := Valeur(FTabRech[TColDebit].Value) - lRdVal;
          TOB(FTOBEcr.Detail[0]).PutValue( FStPrefixe + '_CREDIT'     , lRdVal);
          TOB(FTOBEcr.Detail[1]).PutValue( FStPrefixe + '_CREDIT'     , lRdTva );
          TOB(FTOBEcr.Detail[2]).PutValue( FStPrefixe + '_DEBIT'     , Valeur(FTabRech[TColDebit].Value) );
         end
          else
           begin

            lRdVal := Valeur(FTabRech[TColCredit].Value) / ( 1 + ( Valeur(FTabRech[TColTva].Value) / 100 ) ) ;
            lRdTva := Valeur(FTabRech[TColCredit].Value) - lRdVal;
            TOB(FTOBEcr.Detail[0]).PutValue( FStPrefixe + '_DEBIT'     ,  lRdVal );
            TOB(FTOBEcr.Detail[1]).PutValue( FStPrefixe + '_DEBIT'     ,  lRdTva );
            TOB(FTOBEcr.Detail[2]).PutValue( FStPrefixe + '_CREDIT'    ,  Valeur( FTabRech[TColCredit].Value ) );
          end;
        end
         else
          begin
          // HGImputation.RowCount := 3;
           if FTabRech[TColDebit].Value <> '' then
            begin
             TOB(FTOBEcr.Detail[0]).PutValue( FStPrefixe + '_CREDIT'     ,  Valeur(FTabRech[TColDebit].Value ) );
             TOB(FTOBEcr.Detail[1]).PutValue( FStPrefixe + '_DEBIT'     ,  Valeur(FTabRech[TColDebit].Value ) );
         end
             else
              begin
                TOB(FTOBEcr.Detail[0]).PutValue( FStPrefixe + '_DEBIT'  , Valeur(FTabRech[TColCredit].Value) );
                TOB(FTOBEcr.Detail[1]).PutValue( FStPrefixe + '_CREDIT' , Valeur(FTabRech[TColCredit].Value) );
              end;
          end;

 result := true;

end;

function _SuppCarJocker( const Value : string ; var vInPos : integer ) : string ;
begin

 vInPos := 0 ;
 Result := Value ;

 // suppression des caractères de recherche dans la chaine de recherche
 if Copy( Value, 1 , 1) = cStJocker then
  begin
   Result := Copy(Result, 2, length(Result));
   vInPos := 1; // on a trouver une fois le jocker
  end;

 if Copy( Value, length(Value) , 1) = cStJocker then
  begin
   Result := Copy(Result, 1, length( Result ) - 1);
   vInPos        := vInPos + 1; // on a trouver deux fois le jocker
  end; // if

end ;

function TZEtebacGuide.TestCarJocker ( Const vStTexte : string ; vStChaine : string = '') : boolean ;
var
 lInPosCrit  : integer ;
 lInPos      : integer ;
 lStCrit     : string ;
begin

  if vStChaine = '' then vStChaine := trim(FTabRech[TColLibelle].Value) ;

  lStCrit := _SuppCarJocker( vStTexte , lInPos ) ;

 // place du critere de recherche dans la grille
  if pos(UpperCase( lStCrit ),UpperCase(FTabRech[TColLibelle].Value)) > 0 then
   lInPosCrit := pos(UpperCase( lStCrit ),UpperCase(FTabRech[TColLibelle].Value)) + length(UpperCase(lStCrit)) - 1
    else
     lInPosCrit := -1 ;

  result := ( ( pos('*',vStTexte) > 1 ) and ( pos( UpperCase( lStCrit ) , vStChaine) = 1 ) )   or // cas TPE*
            ( ( pos('*',vStTexte) = 1 ) and ( lInPosCrit = length(vStChaine ) ) )              or // cas *TPE
            ( ( pos('*',vStTexte) = 0 ) and ( UpperCase( lStCrit ) = vStChaine  ) )            or // cas TPE
            ( ( lInPos            = 2 ) and ( pos( UpperCase( lStCrit ) , vStChaine ) > 0 ) ) ;// cas *TPE*
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 07/06/2004
Modifié le ... :   /  /
Description .. : - 07/06/2004 - FB 12652 - LG - gestion du $30 &
Suite ........ : *SERVICE*
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.RechercheGuideParAFB( vStLibelle : string ) : boolean;
var
 lInIndex     : integer ;
 lStTexteRech : string ;
begin

 result := false ;

 if  ( pos( cStIdCodeAFB , vStLibelle ) = 0 ) then exit ;

 lInIndex := Pos(cStIdCodeAFBSupp,vStLibelle) ;

 if lInIndex = 0 then
  begin // pas de &
   lStTexteRech := Trim( Copy( vStLibelle , 2 , length(vStLibelle) ) );
   result       := TestCarJocker(lStTexteRech, FTabRech[TColCodeAFB].Value) ;
   exit ;
  end
   else
    begin
     // test de la partie code AFB
     lStTexteRech := Trim( Copy( vStLibelle , 2 , lInIndex - 2 ) ) ;
     result       := TestCarJocker(lStTexteRech, FTabRech[TColCodeAFB].Value) ;
     if not result then exit ;
     // test du libelle qui suit le code AFB
     lStTexteRech := Trim( Copy( vStLibelle , lInIndex + 1 , length(vStLibelle) ) );
     {JP 13/04/07 : gestion des libellés enrichis}
     result       := TestCarJocker(lStTexteRech) or HasLibelleEnrichi(lStTexteRech);
    end;

end;

{JP 18/01/07 : Gestion des libellés enrichis :
               la présence d'un libellé enrichi est marqué de la sorte µLIB1, µLIB2, µLIB3
{---------------------------------------------------------------------------------------}
function  TZEtebacGuide.HasLibelleEnrichi(aLibelle : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  StP : string;
  StN : string;
  StF : string;
begin
  StP := aLibelle;
  {Il y a une formule sur les libellés enrichis}
  Result := Pos(cStLibEnrichi, StP) > 0;
  if Result then begin
    {Récupération de la formule sur le libellé enrichi
    On conserve que ce qui concerne le libellé enrichi, le reste éventuel concerne le libellé}
    StP := ReadTokenSt(StP);
    StN := Copy(StP, 1, 2); {'#1', '#2', '#3'}
    StF := Copy(StP, 3, Length(StP)); {Condition sur le libellé enrichi}

    {On regarde  si la formule du guide s'applique à l'un des guides
     19/03/07 : Ajout du test (FTabRech[TColLibEnrichiX].Value <> ''), car si vide TestCarJoker fait le test sur le libellé}
    Result := (TestCarJocker(StF, FTabRech[TColLibEnrichi1].Value) and (FTabRech[TColLibEnrichi1].Value <> '') and (StN = cStLibEnrichi + '1')) or
              (TestCarJocker(StF, FTabRech[TColLibEnrichi2].Value) and (FTabRech[TColLibEnrichi2].Value <> '') and (StN = cStLibEnrichi + '2')) or
              (TestCarJocker(StF, FTabRech[TColLibEnrichi3].Value) and (FTabRech[TColLibEnrichi3].Value <> '') and (StN = cStLibEnrichi + '3'));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 09/06/2004
Description .. : Recherche d'un guide par rapport au libelle de la grille ou au
Suite ........ : code AFB de la ligne ( si le guide est parametre avec $AFB )
Suite ........ : le caractère '*' represente le caractère jocker.
Suite ........ :
Suite ........ : - 20/09/2003 - correction du cas *TPE* et le test sur la
Suite ........ : reconnaissance etebas est fait en premier
Suite ........ : - 07/06/2004 - FB 12652 - LG - gestion du $30 &
Suite ........ : *SERVICE*
Suite ........ : - 09/06/2004 - FB 12653 - LG - gestion du $3*
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.RechercheGuideParLibelle : string;
var
 i               : integer ;
 j               : integer ;
 lTOB            : TOB ;
 lTOBEnteteGuide : TOB ;
 lTOBLigneGuide  : TOB ;
 lStTexteRech    : string ;
 lPRecLib        : PRecLib ;
 lInMax          : integer ;
begin

 lTOBEnteteGuide := TOB.Create('',nil,-1);
 lTOBLigneGuide  := nil;
 result          := '';

 try

  for i := 0 to FListeLibelle.Count - 1 do
   begin
    // on sauvegarde les critères de recherche
    lPRecLib     := FListeLibelle.Items[i] ;
    lStTexteRech := lPRecLib^.Value ;
      {JP 18/01/07 : gestion des libellés enrichis}
      if RechercheGuideParAFB( lStTexteRech ) or TestCarJocker(lStTexteRech) or HasLibelleEnrichi(lStTexteRech) then
       begin

       lTOB := lPRecLib^.PTOB ;

       FirstGuide(VarToStr(lTOB.GetValue('EG_GUIDE')));

       if VerifSensDuGuide(lTOB) then
        begin
         lTOBLigneGuide  := TOB.Create('ECRGUI',lTOBEnteteGuide,-1);
         lTOBLigneGuide.PutValue('EG_GUIDE'      , lTOB.GetValue('EG_GUIDE'));
         lTOBLigneGuide.PutValue('EG_LIBELLE'    , lTOB.GetValue('EG_LIBELLE'));
         lTOBLigneGuide.PutValue('EG_REFINTERNE' , lTOB.GetValue('GU_LIBELLE'));
         lTOBLigneGuide.AddChampSupValeur('PRIO' , lPRecLib^.Prio) ;
        end; // if

      end; // if

   end; // for

  if lTOBEnteteGuide.Detail.Count = 0 then
   exit;

  if assigned(lTOBLigneGuide) and ( lTOBEnteteGuide.Detail.Count = 1 ) then
   result := lTOBLigneGuide.GetValue('EG_GUIDE')
    else
     if assigned( FGrilleReleve ) then
      begin
       lTOB := LookUpTob ( FGrilleReleve ,
                           lTOBEnteteGuide ,
                           'Guide pour : ' + FTabRech[TColLibelle].Value ,
                           'EG_GUIDE;EG_REFINTERNE',
                           'Code Guide;Libellé' );
       if assigned(lTOB) then
        result := lTOB.GetValue('EG_GUIDE');
      end // if
       else
        begin
         result := lTOBEnteteGuide.Detail[0].GetValue('EG_GUIDE') ;
         lInMax := lTOBEnteteGuide.Detail[0].GetValue('PRIO') ;
         for j := 1 to lTOBEnteteGuide.Detail.Count - 1 do
          begin
           if lTOBEnteteGuide.Detail[j].GetValue('PRIO') < lInMax then
            begin
             result := lTOBEnteteGuide.Detail[j].GetValue('EG_GUIDE') ;
             lInMax := lTOBEnteteGuide.Detail[j].GetValue('PRIO') ;
            end // if
             else
              if lTOBEnteteGuide.Detail[j].GetValue('PRIO') = lInMax then
               begin
                result := '' ;
                break ;
               end ;
          end ; // for
        end ; // else

  FTypeRecherche := FTypeRecherche - [TColLibelle];

 finally
  if assigned(lTOBEnteteGuide) then lTOBEnteteGuide.Free;
 end; // try

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 10/06/2005
Description .. : Recherche d'un guide par rapport au credit de la ligne
Suite ........ : - LG - FB 13666 - correction de la reconnaissance des 
Suite ........ : montants
Mots clefs ... : 
*****************************************************************}
function TZEtebacGuide.RechercheGuideParCredit : string;
var
 i                : integer;
 lTOB             : TOB;
 lTOBEnteteGuide  : TOB;
 lTOBLigneGuide   : TOB;
begin

 result          := '';
 lTOBLigneGuide  := nil;

 if FTabRech[TColCredit].Value = '' then exit;

 lTOBEnteteGuide := TOB.Create('',nil,-1);

 try

 for i := 0 to FListeMontant.Count - 1 do
  if Valeur(FListeMontant.Strings[i])=  Valeur( FTabRech[TColCredit].Value ) then
    begin

     lTOB := TOB( FListeMontant.Objects[i] );

     FirstGuide(VarToStr(lTOB.GetValue('EG_GUIDE')));

     if VerifSensDuGuide(lTOB) then
      begin
       lTOBLigneGuide  := TOB.Create('ECRGUI',lTOBEnteteGuide,-1);
       lTOBLigneGuide.PutValue('EG_GUIDE'      , lTOB.GetValue('EG_GUIDE'));
       lTOBLigneGuide.PutValue('EG_LIBELLE'    , lTOB.GetValue('EG_LIBELLE'));
       lTOBLigneGuide.PutValue('EG_REFINTERNE' , lTOB.GetValue('GU_LIBELLE'));
      end; // if

    end; // if

  if lTOBEnteteGuide.Detail.Count = 0 then
   exit;

  if assigned(lTOBLigneGuide) and ( lTOBEnteteGuide.Detail.Count = 1 ) then
   result := lTOBLigneGuide.GetValue('EG_GUIDE')
    else
     if assigned( FGrilleReleve ) then
      begin
       lTOB := LookUpTob ( FGrilleReleve ,
                           lTOBEnteteGuide ,
                           'Guide pour : ' + FTabRech[TColCredit].Value ,
                           'EG_GUIDE;EG_REFINTERNE',
                           'Code Guide;Libellé' );
       if assigned(lTOB) then
        result := lTOB.GetValue('EG_GUIDE');
     end; // if

  FTypeRecherche := FTypeRecherche - [TColCredit];

 finally
  if assigned(lTOBEnteteGuide) then lTOBEnteteGuide.Free;
 end; // try

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 10/06/2005
Description .. : Recherche d'un guide par rapport au debit de la ligne
Suite ........ : - LG - FB 13666 - correction de la reconnaissance des 
Suite ........ : montants
Mots clefs ... : 
*****************************************************************}
function TZEtebacGuide.RechercheGuideParDebit : string;
var
 i                : integer;
 lTOB             : TOB;
 lTOBEnteteGuide  : TOB;
 lTOBLigneGuide   : TOB;
begin

 result          := '';
 lTOBLigneGuide  := nil;

 if FTabRech[TColDebit].Value = '' then exit;

 lTOBEnteteGuide := TOB.Create('',nil,-1);

 try

 for i := 0 to FListeMontant.Count - 1 do
  if Valeur( FListeMontant.Strings[i] ) = Valeur ( FTabRech[TColDebit].Value ) then
    begin

     lTOB := TOB( FListeMontant.Objects[i] );

     FirstGuide(VarToStr(lTOB.GetValue('EG_GUIDE')));

     if VerifSensDuGuide(lTOB) then
      begin
       lTOBLigneGuide  := TOB.Create('ECRGUI',lTOBEnteteGuide,-1);
       lTOBLigneGuide.PutValue('EG_GUIDE'      , lTOB.GetValue('EG_GUIDE'));
       lTOBLigneGuide.PutValue('EG_LIBELLE'    , lTOB.GetValue('EG_LIBELLE'));
       lTOBLigneGuide.PutValue('EG_REFINTERNE' , lTOB.GetValue('GU_LIBELLE'));
      end; // if

    end; // if

  if lTOBEnteteGuide.Detail.Count = 0 then
   exit;

  if assigned(lTOBLigneGuide) and ( lTOBEnteteGuide.Detail.Count = 1 ) then
   result := lTOBLigneGuide.GetValue('EG_GUIDE')
    else
     if assigned( FGrilleReleve ) then
      begin
       lTOB   := LookUpTob ( FGrilleReleve ,
                            lTOBEnteteGuide ,
                            'Guide pour : ' + FTabRech[TColDebit].Value ,
                            'EG_GUIDE;EG_REFINTERNE',
                            'Code Guide;Libellé' );
       if assigned(lTOB) then
        result := lTOB.GetValue('EG_GUIDE');
      end; // if

  FTypeRecherche := FTypeRecherche - [TColDebit];

 finally
  if assigned(lTOBEnteteGuide) then lTOBEnteteGuide.Free;
 end; // try

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Recherche d'un guide par rapport au compte de la ligne.
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.RechercheGuideParImputation : string;
var
 i                : integer;
 lTOB             : TOB;
 lTOBEnteteGuide  : TOB;
 lTOBLigneGuide   : TOB;
begin

 result          := '';
 lTOBLigneGuide  := nil;

 if FTabRech[TColGeneral].Value = '' then exit;

 lTOBEnteteGuide := TOB.Create('',nil,-1);

 try

 for i := 0 to FListeGeneral.Count - 1 do
  if pos( FListeGeneral.Strings[i] ,UpperCase(FTabRech[TColGeneral].Value) ) > 0 then
    begin

     lTOB := TOB( FListeGeneral.Objects[i] );

     FirstGuide(VarToStr(lTOB.GetValue('EG_GUIDE')));

     if VerifSensDuGuide(lTOB) then
      begin
       lTOBLigneGuide  := TOB.Create('ECRGUI',lTOBEnteteGuide,-1);
       lTOBLigneGuide.PutValue('EG_GUIDE'      , lTOB.GetValue('EG_GUIDE'));
       lTOBLigneGuide.PutValue('EG_LIBELLE'    , lTOB.GetValue('EG_LIBELLE'));
       lTOBLigneGuide.PutValue('EG_REFINTERNE' , lTOB.GetValue('GU_LIBELLE'));
      end; // if

    end; // if

  if lTOBEnteteGuide.Detail.Count = 0 then
   exit;

  if lTOBEnteteGuide.Detail.Count = 1 then
   result := lTOBLigneGuide.GetValue('EG_GUIDE')
    else
     if assigned( FGrilleReleve ) then
      begin
       lTOB := LookUpTob ( FGrilleReleve ,
                           lTOBEnteteGuide ,
                           'Guide pour : ' + FTabRech[TColGeneral].Value ,
                           'EG_GUIDE;EG_REFINTERNE',
                           'Code Guide;Libellé' );

       if assigned(lTOB) then
        result := lTOB.GetValue('EG_GUIDE');
      end; // if

  FTypeRecherche := FTypeRecherche - [TColGeneral];

 finally
  if assigned(lTOBEnteteGuide) then lTOBEnteteGuide.Free;
 end; // try

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Recherche d'un guide par rapport au compte de la ligne.
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.RechercheGuideTotal : string;
var
 i                : integer;
 lTOB             : TOB;
 lTOBEnteteGuide  : TOB;
 lTOBLigneGuide   : TOB;
begin

 result          := '';
 lTOBLigneGuide  := nil;

 lTOBEnteteGuide := TOB.Create('',nil,-1);

 try

 for i := 0 to FListeGeneral.Count - 1 do
  begin

     lTOB := TOB( FListeGeneral.Objects[i] );

     FirstGuide(VarToStr(lTOB.GetValue('EG_GUIDE')));

     if VerifSensDuGuide(lTOB) then
      begin
       lTOBLigneGuide  := TOB.Create('ECRGUI',lTOBEnteteGuide,-1);
       lTOBLigneGuide.PutValue('EG_GUIDE'      , lTOB.GetValue('EG_GUIDE'));
       lTOBLigneGuide.PutValue('EG_LIBELLE'    , lTOB.GetValue('EG_LIBELLE'));
       lTOBLigneGuide.PutValue('EG_REFINTERNE' , lTOB.GetValue('GU_LIBELLE'));
      end; // if

  end; // for

  if lTOBEnteteGuide.Detail.Count = 0 then
   exit;

  if lTOBEnteteGuide.Detail.Count = 1 then
   result := lTOBLigneGuide.GetValue('EG_GUIDE')
    else
     if assigned( FGrilleReleve ) then
      begin
       lTOB := LookUpTob ( FGrilleReleve ,
                           lTOBEnteteGuide ,
                           'Liste des guides',
                           'EG_GUIDE;EG_REFINTERNE',
                           'Code Guide;Libellé' );

       if assigned(lTOB) then
        result := lTOB.GetValue('EG_GUIDE');
      end; // if


 finally
  if assigned(lTOBEnteteGuide) then lTOBEnteteGuide.Free;
 end; // try

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 09/11/2001
Description .. : Fonction de recherche globale d'un guide par rapport à une
Suite ........ : ligne de grille.
Suite ........ : Si la grille est passer en parametre, en cas de choix
Suite ........ : multiples on presente une fenetre de choix.
Suite ........ : Attention : pour les guide de treso les colonnes debit et
Suite ........ : credit sont inversé
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.RechercheGuideEnBase( vTOB,vTOBResult : TOB ; G : THGrid = nil ) : boolean;
var
 lValue : string;
 i      : integer;
begin

 result          := false;

 if ( ZCompte = nil ) or ( vTOBResult = nil )  then exit;

 // récuperation des valeurs de la grille
 FGrilleReleve   := G;
 FTOBEcr         := vTOBResult;
 FTypeRecherche  := FTypeRechercheReference; //JP 16/01/07 : [TColGeneral,TColLibelle,TColDebit,TColCredit]; ré-initialisation des options de recherche

 RecupereValeurTOB(vTOB,true);

 // il n'y a pas de guide
 if not assigned(FTOB) then
  begin
  result := GenerationGuide;
  exit;
  end
   else
    if FTOB.Detail.Count = 0 then
     begin
      if FTabRech[TColGeneral].Value <> '' then // si le compte d'imput ets definie
       begin
        result := GenerationGuide; // on genere un guide en auto
        exit;
       end
         else
          exit;
     end; // if

 if lValue = '' then
 for i := 0 to 3 do
  if TTypeRecherche(i) in FTypeRecherche then
   begin
     lValue := FTabRech[TTypeRecherche(i)].ProcAction;
     if lValue <> '' then
      break;
   end; // if

 // on stocke le code du guide trouvé
 // valeur utilisé dans CalculeGuide
 FStEG_GUIDE := lValue;

 if lValue <> '' then
  result := CalculeGuide
   else
    if FTabRech[TColGeneral].Value <> '' then
     result := GenerationGuide
      else
       if assigned(G) then
       begin
        lValue      := RechercheGuideTotal;
        FStEG_GUIDE := lValue;
        if lValue <> '' then
         result := CalculeGuide;
       end;

end;

function TZEtebacGuide.RechercheDernierGuide ( vTOB,vTOBResult : TOB ; G : THGrid = nil ) : boolean ;
var
 lValue : string;
begin

 result := false;

 if ( ZCompte = nil ) or ( vTOBResult = nil )  then exit;

 // récuperation des valeurs de la grille
 FGrilleReleve  := G;

 FTOBEcr        := vTOBResult;

 FTypeRecherche := FTypeRechercheReference; //JP 16/01/07 [TColGeneral,TColLibelle,TColDebit,TColCredit]; ré-initialisation des options de recherche

 RecupereValeurTOB(vTOB,true);

 // il n'y a pas de guide
 if not assigned(FTOB) then
  begin
  result := GenerationGuide;
  exit;
  end
   else
    if FTOB.Detail.Count = 0 then
     begin
      if FTabRech[TColGeneral].Value <> '' then // si le compet d'imput ets definie
       begin
        result := GenerationGuide; // on genere un guide en auto
        exit;
       end
         else
          exit;
     end; // if

 if FStEG_GUIDE <> '' then
  result := CalculeGuide
   else
    if FTabRech[TColGeneral].Value <> '' then
     result := GenerationGuide
      else
       if assigned(G) then
       begin
        lValue      := RechercheGuideTotal;
        FStEG_GUIDE := lValue;
        if lValue <> '' then
         result := CalculeGuide;
       end;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/07/2001
Modifié le ... :   /  /
Description .. : Test si l'on doit s'arreter sur la colonne
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.IsColStop(ARow, ACol : integer): boolean;
var
 lStArret : string;
begin

 result := false;

 FCurrentTOB := FTOB.FindFirst(['EG_NUMLIGNE','EG_GUIDE'],[ARow,FStEG_GUIDE],false);

 if FCurrentTOB = nil then exit;

 lStArret    := copy(StARRET,1,6); // la grille de saisie ne comporte que 6 cases

 if ( length(lStArret) >= ACol ) and ( lStArret <> '' ) then
  result     := trim(lStArret[ACol]) = 'X';

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 23/07/2001
Modifié le ... :   /  /
Description .. : Test si le guide possède au moins une colonne d'arret
Mots clefs ... :
*****************************************************************}
function TZEtebacGuide.PossedeArret : boolean;
var
 i,j  : integer;
begin

 result := false;

 for i := 1 to 6 do
  for j := 1 to 6 do
   if IsColStop(i,j) then
    begin
     Result := true;
     exit;
    end; // if

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 10/06/2005
Description .. : Contrôle du sens du guide
Suite ........ : - LG - FB 13666 - correction de la reconnaissance des 
Suite ........ : montants
Mots clefs ... : 
*****************************************************************}
function TZEtebacGuide.VerifSensDuGuide( Value : TOB ) : boolean;
begin
 // contrôle du sens du guide
 result := ( ( FTabRech[TColCredit].Value = '' ) and ( Value.GetValue('SENS') = '0' ) ) or
           ( ( FTabRech[TColDebit].Value = '' ) and ( Value.GetValue('SENS') = '1' ) );

 if not result then exit ;

 // on contrôle que le compte d'imputation du guide est le même que celui de la grille de releve
 if result and ( FTabRech[TColGeneral].Value <> '' ) then
  result := ( StGENERAL = FTabRech[TColGeneral].Value );

 if not result then exit ;

 // si la derniere ligne contient un montant, il doit être egal au solde du guide
 if result and isNumeric( FTabRech[TColCredit].Value ) and isNumeric( Value.GetValue('SOLDE') ) then
  result := Valeur(FTabRech[TColCredit].Value) = Valeur(Value.GetValue('SOLDE'));

 if not result then exit ; 

 // si la derniere ligne contient un montant, il doit être egal au solde du guide
 if result and isNumeric( FTabRech[TColDebit].Value ) and isNumeric( Value.GetValue('SOLDE') ) then
  result := Valeur(FTabRech[TColDebit].Value)  = Valeur(Value.GetValue('SOLDE'));

end;


function TZEtebacGuide.EstVide : boolean;
begin
 result := true ;
 if FTOB <> nil then
 result := FTOB.Detail.Count = 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Recupere le prefixe de la table de travail en fonction de son
Suite ........ : nom
Mots clefs ... :
*****************************************************************}
procedure TZEtebacGuide.SetLeNomTable( Value : string );
begin
 FStPrefixe    := TableToPrefixe ( Value );
 FStLeNomTable := Value;
end;

procedure TZEtebacGuide.NotifyError( vMessage, vDelphiMessage , vFunctionName : string);
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

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 15/05/2002
Modifié le ... : 18/10/2007
Description .. : Suppression de tous les guides enregsitrer en memoire
Suite ........ : -28/11/2002 - correction du plantage qd on fermait la
Suite ........ : fenetre des guides sans que l'on soit rentrer en saisie
Suite ........ : - LG - 18/10/2007 - remise a zero de FStSQL pour relancer 
Suite ........ : le load du guide avec les meme parametre
Mots clefs ... : 
*****************************************************************}
procedure TZEtebacGuide.ClearGuide;
begin
 if FTOB<>nil then FTOB.ClearDetail;
 FListeLibelle.Clear;
 FListeMontant.Clear;
 FListeGeneral.Clear;
 FStSQL := '' ;
end;

procedure TZEtebacGuide.RecupereValeurTOB( vTOB : TOB ; vBoInitialisation : boolean = false );
begin

 if vTOB = nil then raise Exception.Create('Calcul du guide impossible, la récupération des données de la grille a échoué !');

 if vTOB.GetValue( FStPrefixe + '_DEBIT') <> 0 then
  FTabRech[TColCredit].Value        := vTOB.GetValue( FStPrefixe + '_DEBIT')
   else
    FTabRech[TColCredit].Value      := '';

 if vTOB.GetValue( FStPrefixe + '_CREDIT') <> 0 then
  FTabRech[TColDebit].Value        := vTOB.GetValue( FStPrefixe + '_CREDIT')
   else
    FTabRech[TColDebit].Value       := '';

 FTabRech[TColLibelle].Value        := vTOB.GetValue( FStPrefixe + '_LIBELLE') ;
 FTabRech[TColGeneral].Value        := vTOB.GetValue( FStPrefixe + '_GENERAL') ;
 //FTabRech[TColGeneral].Value        := vTOB.GetValue( FStPrefixe + '_AUXILIAIRE') ;
 FTabRech[TColDate].Value           := vTOB.GetValue( FStPrefixe + '_DATECOMPTABLE') ;
 FTabRech[TColRefInterne].Value     := vTOB.GetValue( FStPrefixe + '_REFINTERNE') ;
 FTabRech[TColRefExterne].Value     := vTOB.GetValue( FStPrefixe + '_REFEXTERNE') ;
 FTabRech[TColTVA].Value            := FloatToStr(vTOB.GetValue( FStPrefixe + '_TAUXTVA') ) ;
 FTabRech[TColCodeAFB].Value        := vTOB.GetValue( FStPrefixe + '_CODEAFB') ;
 FTabRech[TColNaturePiece].Value    := vTOB.GetValue( FStPrefixe + '_NATUREPIECE') ;
 if vBoInitialisation then
  begin
   FBoPasTVASaisie                  := vTOB.GetValue( FStPrefixe + '_TVASAISIE') = '-' ;
   FStTauxTvaInitial                := FTabRech[TColTVA].Value ;
   if FTabRech[TColCredit].Value <> '' then
    FStSolde := FTabRech[TColCredit].Value
     else
      FStSolde := FTabRech[TColDebit].Value;
  end;

  {JP 19/01/07 : Gestion des libellés enrichis}
  FTabRech[TColLibEnrichi1].Value := vTOB.GetValue( FStPrefixe + '_LIBELLE1') ;
  FTabRech[TColLibEnrichi2].Value := vTOB.GetValue( FStPrefixe + '_LIBELLE2') ;
  FTabRech[TColLibEnrichi3].Value := vTOB.GetValue( FStPrefixe + '_LIBELLE3') ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 18/10/2007
Modifié le ... :   /  /    
Description .. : - LG - 18/10/2007 - FB 21136 - creation d'un guide a partir 
Suite ........ : d'une ecriture
Mots clefs ... : 
*****************************************************************}
function CCreerCommeGuide( vTOB : TOB ; vStPref : string = 'E_' ) : string ;
var
 lTOBGuide     : TOB ;
 lTOB          : TOB ;
 i             : integer ;
 lStCode       : string ;
 lTOBRef       : TOB ;
 lBoMvtAuDebit : boolean ;
 lSt1          : string ;
 lSt2          : string ;
begin

 result := '' ;

 if ( vTOB = nil ) or ( vTOB.Detail = nil )  then exit ;

 lStCode       :=  CCalculeCodeGuide('NOR')  ;

 lBoMvtAuDebit :=  vTOB.Detail[vTOB.Detail.Count - 1].GetValue(vStPref + 'DEBIT') = 0 ;

 lTOBRef       :=  vTOB.Detail[0] ;

 lTOBGuide    := TOB.Create('GUIDE' , nil , -1 ) ;
 lTOBGuide.PutValue('GU_TYPE'           , 'NOR' ) ;
 lTOBGuide.PutValue('GU_GUIDE'          , lStCode ) ;
 lTOBGuide.PutValue('GU_UTILISATEUR'    , V_PGI.User) ;
 lTOBGuide.PutValue('GU_SOCIETE'        , V_PGI.CodeSociete) ;
 lTOBGuide.PutValue('GU_DEVISE'         , V_PGI.DevisePivot ) ;
 lTOBGuide.PutValue('GU_JOURNAL'        , lTOBRef.GetValue(vStPref + 'JOURNAL' ) ) ;
 lTOBGuide.PutValue('GU_NATUREPIECE'    , lTOBRef.GetValue(vStPref + 'NATUREPIECE' ) ) ;
 lTOBGuide.PutValue('GU_DEVISE'         , 'EUR' ) ;
 lTOBGuide.PutValue('GU_ETABLISSEMENT'  , lTOBRef.GetValue(vStPref + 'ETABLISSEMENT' ) ) ;
 lTOBGuide.PutValue('GU_UTILISATEUR'    , lTOBRef.GetValue(vStPref + 'UTILISATEUR' ) ) ;
 lTOBGuide.PutValue('GU_LIBELLE'        , '<<automatique>>' ) ;
 if vStPref = 'E_' then
  lTOBGuide.PutValue('GU_TRESORERIE'    , '-' )
   else
    lTOBGuide.PutValue('GU_TRESORERIE'    , 'X' ) ;


 for i := 0 to vTOB.Detail.Count - 1 do
  begin
   lTOBRef :=  vTOB.Detail[i] ;
   lTOB    := TOB.Create('ECRGUI', lTOBGuide , -1 ) ;
   lTOB.PutValue('EG_TYPE'       , 'NOR' ) ;
   lTOB.PutValue('EG_GUIDE'      , lStCode ) ;
   lTOB.PutValue('EG_NUMLIGNE'   , i + 1 ) ;
   lTOB.PutValue('EG_GENERAL'    , lTOBRef.GetValue(vStPref + 'GENERAL' ) ) ;
   lTOB.PutValue('EG_AUXILIAIRE' , lTOBRef.GetValue(vStPref + 'AUXILIAIRE' ) ) ;
   if length(lTOBRef.GetValue(vStPref + 'GENERAL' )) = GetInfoCpta( fbGene ).Lg then
    lSt1 := '-'
     else
      lSt1 := 'X' ;
   if i < vTOB.Detail.Count - 1 then
    begin
     if lBoMvtAuDebit then
      lSt2 := 'X--'
      else
       lSt2 := '-X-' ;
    end
     else
      lSt2 := '---' ;

   lTOB.PutValue('EG_ARRET' , lSt1 + '---' + lSt2 ) ;

  end ; // for

 if lBoMvtAuDebit then
  lTOBGuide.Detail[vTOB.Detail.Count - 1].PutValue('EG_CREDITDEV', '[SOLDE]')
   else
    lTOBGuide.Detail[vTOB.Detail.Count - 1].PutValue('EG_DEBITDEV', '[SOLDE]') ;


 lTOBGuide.InsertDB(nil) ;

 result := lStCode ;

 AvertirTable('ttGuideEcr') ; 

end ;



end.
