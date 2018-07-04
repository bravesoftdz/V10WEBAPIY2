{***********UNITE*************************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : Classe gérant la recherche en guide pour la saisie de
Suite ........ : trésorie.
Mots clefs ... :
*****************************************************************}
unit ZGuideTreso;

interface

uses

 Classes,
 {$IFNDEF EAGLCLIENT}
 dbTables,
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
 UCstEcriture,
 ZReleveBanque,
 ZCompte, // pour le TZCompte
 // AGL
 LookUp,
 SaisUtil,
 UTOFLOOKUPTOB,
 UTOB ;


type

 TTypeRecherche  = ( TColGeneral,TColDebit,TColCredit,TColLibelle,TColTVA, TColDate, TColNumeroPiece, TColPiece , TColCodeAFB);
 TTypeRecherches = set of TTypeRecherche;

 TProcAction     = function : string of object;

 TRecRech       = record
  Value         : string;
  ProcAction    : TProcAction;
 end;


 TZGuideTreso = Class
  private
   FTypeContexte             : TTypeContexte;
   FTOB                      : TOB;             // TOB contenant l'ensemble des guides
   FCurrentTOB               : TOB;             // pointeur sur la ligne de guide courante
   FTOBEcr                   : TOB;             // TOB genere par le moteur des guides
   FStEG_GUIDE               : string;          // code du guide courant
   FGrilleReleve             : THGrid;          // pointeur sur la grille de saisie
   FListeLibelle             : TStringList;     // Liste des libelles de recherches pour les guides
   FListeMontant             : TStringList;     // Liste des montant de recherche pour les guides
   FListeGeneral             : TStringList;     // Liste des comptes de recherche pour les guides
   FTypeRecherche            : TTypeRecherches; // Liste des recherche a effectuer
   FStJournalContrepartie    : string;          // code du journal de la contrepartie
   FStCompteTVA              : string;          // Compte de TVA
   FStEtablissement          : string;          // Code de l'etablissement pour les recherche des guides
   FStCompteContrepartie     : string;          // Compte de contrpartie : compte de banque du journal de tresorie
   FInNumLigne               : integer;
   FTabRech                  : array[TTypeRecherche] of TRecRech;
   //FZReleveBanque            : TZReleveBanque;  // Pointeur sur un objet de releve bancaire qui sera remplit par un guide
   FStDevise                 : string;          // code de la devise pour la recherche des guides
   FStLeNomTable             : string;          // Nom de la table -> utilisé pour cree la TOB de retour
   FStPrefixe                : string;
   FZCompte                  : TZCompte;  // objet de gestion des comptes

   procedure SetLeNomTable ( Value : string );
   function  GetFieldProperty ( vIndex : integer ) : variant;
   function  GetFormule( StFormule : string ) : variant;
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
   function  GetRowFormule(var StFormule: string): Integer;
   procedure DeleteGuide(vStGUIDE: string);

 //  property StTYPE                       : variant index 0  read GetFieldProperty;
   property StGUIDE                      : variant index 1  read GetFieldProperty;
 //  property InNUMLIGNE                   : variant index 2  read GetFieldProperty;
   property StGENERAL                    : variant index 3  read GetFieldProperty;
   property StREFINTERNE                 : variant index 4  read GetFieldProperty;
 //  property StREFEXTERNE                 : variant index 5  read GetFieldProperty;
   property StLIBELLE                    : variant index 6  read GetFieldProperty;
   property StDEBITDEV                   : variant index 7  read GetFieldProperty;
   property StCREDITDEV                  : variant index 8  read GetFieldProperty;
   property StARRET                      : variant index 9  read GetFieldProperty;
   procedure NotifyError(vMessage, vDelphiMessage, vFunctionName: string);

  public

   constructor Create;
   destructor Destroy; override;

   function  RecalculGuide( vTOBEcr : TOB ; vInNumLigne : integer) : TOB;
   function  RechercheGuideEnBase( vStCredit ,
                                   vStDebit ,
                                   vStLibelle ,
                                   vStGeneral,
                                   vStTVA,
                                   vStDate,
                                   vStNumeroPiece,
                                   vStPiece,
                                   vStCodeAFB : string;
                                   vTOBResult : TOB ;
                                   G : THGrid = nil
                                 ) : boolean;
   function  IsColStop(ARow,ACol : integer) : boolean;
   function  FirstGuide( Value : string) : integer;
   function  PossedeArret : boolean;

   function  Load : boolean;
   function  EstVide : boolean;
   procedure ClearGuide;

   property StJournalContrepartie        : string           read FStJournalContrepartie     write FStJournalContrepartie;
   property StEtablissement              : string           read FStEtablissement           write FStEtablissement;
   property StCompteContrepartie         : string           read FStCompteContrepartie      write FStCompteContrepartie;
   property StCompteTVA                  : string           read FStCompteTVA               write FStCompteTVA;
   property ZCompte                      : TZCompte         read FZCompte                   write FZCompte;
   property StDevise                     : string           read FStDevise                  write FStDevise;
   property StLeNomTable                 : string           read FStLeNomTable              write SetLeNomTable;

 end;

implementation

const
 cStIdCodeAFB = '$';

constructor TZGuideTreso.Create;
begin
 FListeLibelle                        := TStringList.Create;
 FListeMontant                        := TStringList.Create;
 FListeGeneral                        := TStringList.Create;
 FTypeRecherche                       := [TColGeneral,TColDebit,TColCredit,TColLibelle];

 FTabRech[TColLibelle].ProcAction     := RechercheGuideParLibelle;
 FTabRech[TColDebit].ProcAction       := RechercheGuideParDebit;
 FTabRech[TColCredit].ProcAction      := RechercheGuideParCredit;
 FTabRech[TColGeneral].ProcAction     := RechercheGuideParImputation;

 FTypeContexte                        := TModeGuide;

end;

destructor TZGuideTreso.Destroy;
begin

  //VideTStringListe(FListeLibelle);
  //VideTStringListe(FListeMontant);
  //VideTStringListe(FListeGeneral);

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

function TZGuideTreso.GetFieldProperty ( vIndex : integer ) : variant;
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
procedure TZGuideTreso.DeleteGuide( vStGUIDE : string);
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
Modifié le ... : 09/11/2001
Description .. : -Chargement des guides au démarrage.
Suite ........ : - Suppression des guides comportant un point d'arrêt sur la
Suite ........ : dernière ligne
Suite ........ : - Calcul du sens des guides et affectation du solde sur
Suite ........ : toutes les lignes
Suite ........ : - Stockage des montants, des libelles de recherche,
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TZGuideTreso.Load : boolean;
var
 Q           : TQuery;
 lStCrit     : string;
 lIndex      : integer;
 i,j         : integer;
 lTOB        : TOB;
 lStValeur   : string;
 lStSolde  : string;
begin

 result := false;

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

 Q := OpenSQL('select EG_GUIDE , '                                     +
              'EG_NUMLIGNE, '                                          +
              'EG_GENERAL, '                                           +
              'EG_REFINTERNE, '                                        +
              'EG_REFEXTERNE, '                                        +
              'EG_LIBELLE, '                                           +
              'EG_DEBITDEV, '                                          +
              'EG_CREDITDEV, '                                         +
              'EG_ARRET, '                                             +
              'GU_LIBELLE '                                            +
              'from ECRGUI, GUIDE '                                    +
              'where EG_TYPE = "NOR" '                                 +
              ' and GU_JOURNAL = "' + FStJournalContrepartie + '" '    +
              ' and GU_GUIDE = EG_GUIDE '                              +
              ' and GU_TYPE = EG_TYPE '                                +
              ' and GU_ETABLISSEMENT = "' + FStEtablissement  + '" '   +
              ' and GU_DEVISE = "' + FStDevise + '" '                  +
              'order by EG_GUIDE , EG_NUMLIGNE',
              true );
 // on charge l'ensemble des guides
 FTOB.LoadDetailDB('ECRGUI','','',Q,true);

 if FTOB.Detail.Count = 0 then
  begin
   FCurrentTOB := nil;
   PGIInfo( cStTexteAucunGuide , cStTexteTitreFenetre);
   exit;
  end; // if

 FCurrentTOB := FTOB.Detail[0];
 // on stocke pour chaque ligne le sens du guide et le solde total du guide
 // pour les verif de validité du guide
 FCurrentTOB.AddChampSup('SENS'   , true);
 FCurrentTOB.AddChampSup('SOLDE'  , true);

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

       if ( StDEBITDEV <> '' ) then
        begin

         if ( isNumeric(StDEBITDEV) )  then
          FListeMontant.AddObject(VarToStr(StDEBITDEV),TObject(FCurrentTOB));

         FCurrentTOB.PutValue('SENS','0');
         FCurrentTOB.PutValue('SOLDE',StDEBITDEV);

        end; // if

       if ( StCREDITDEV <> '' ) then
        begin

         if ( isNumeric(StCREDITDEV) ) then
          FListeMontant.AddObject(VarToStr(StCREDITDEV),TObject(FCurrentTOB));

         FCurrentTOB.PutValue('SENS','1');
         FCurrentTOB.PutValue('SOLDE',StCREDITDEV);

        end; // if
      end; // else

   FCurrentTOB := FTOB.FindNext(['EG_GENERAL'],[StCompteContrepartie],false);

 end; // while

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

     lStCrit      := copy( StLIBELLE,1,pos(';',StLIBELLE ) - 1  );
     if lStCrit = '' then
      FListeLibelle.AddObject( UpperCase(StLIBELLE) ,TObject(FCurrentTOB) )
       else
        FListeLibelle.AddObject(UpperCase(copy( StLIBELLE,1,pos(';',StLIBELLE ) - 1  )) ,TObject(FCurrentTOB) );

    end; // if

   // stockage des comptes
   if StGENERAL <> '' then
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

procedure TZGuideTreso.FindNext;
begin

 if not assigned(FTOB) then exit;

 FCurrentTOB := FTOB.FindNext(['EG_GUIDE'], [FStEG_GUIDE],false);

 end;


function TZGuideTreso.FirstGuide( Value : string) : integer;
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
function TZGuideTreso.GetRowFormule ( var StFormule : string ) : Integer ;
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
function TZGuideTreso.GetFormule( StFormule : string) : Variant ;
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
   if FTabRech[TColCredit].Value <> '' then
    lValue := FTabRech[TColCredit].Value
     else
      lValue := FTabRech[TColDebit].Value;
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
      else
       if StFormule = 'E_DATECOMPTABLE' then
        lValue := FTabRech[TColDate].Value
         else
          if StFormule = 'E_REFINTERNE' then
            lValue := FTabRech[TColNumeroPiece].Value
             else
              if StFormule = 'E_REFEXTERNE' then
               lValue := FTabRech[TColPiece].Value
                else
                 if StFormule = 'E_LIBELLE' then
                  begin
                   if ( lInRow = 1 ) then
                    lValue := FTabRech[TColLibelle].Value
                     else
                      begin
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
function TZGuideTreso.MoteurCalculGuide ( vTOBLigneEcr : TOB ; vInNumLigne : integer) : boolean;
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

 function IsNumericLG ( Value : string ) : boolean;
 begin
  result := ( pos('-', Copy( Value, 2 , length(Value) ) ) = 0 ) and ( pos('/',Value) = 0 ) and ( pos('*',Value) = 0 ) and ( pos('+',Value) = 0 );
 end;

begin

 result           := false;

 try

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
       else
        if lStCritLibelle <> '' then
         begin
          lStValeurCell                       := trim(GFormule(lStCritLibelle, GetFormule, nil, 1));
           if ( lStValeurCell <> '' ) then
              lStLibelle          := lStValeurCell
               else
               lStLibelle         := lStCritRech;
         end; //if
    end // critere de recherche
     else
      if trim(lStLibelleGuide) <> '' then
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


   if StDEBITDEV <> '' then
    begin
     lStValeurCell  := GFormule(StDEBITDEV, GetFormule, nil, 1);
     if IsNumeric(lStValeurCell) and not IsNumericLG(lStValeurCell) then
      begin
       NotifyError( TraduireMemoire('Erreur de formule dans le Guide n°')          + FStEG_GUIDE
                    + ' ' + RechDom('TTGUIDEECR', FStEG_GUIDE ,false)              + #13#10
                    + TraduireMemoire('La formule ') + StDEBITDEV + TraduireMemoire(' est incorrecte') ,
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

   if StCREDITDEV <> '' then
    begin
     lStValeurCell  := GFormule(StCREDITDEV, GetFormule, nil, 1);
     if IsNumeric(lStValeurCell) and not IsNumericLG(lStValeurCell) then
      begin
       NotifyError( TraduireMemoire('Erreur de formule dans le Guide n°')          + FStEG_GUIDE
                    + ' ' + RechDom('TTGUIDEECR', FStEG_GUIDE ,false)  + #13#10
                    + TraduireMemoire('La formule ') + StCREDITDEV + TraduireMemoire(' est incorrecte') ,
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
   if ( StGENERAL <> '' ) then
    vTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'     , StGENERAL );
   if lStNumeroPiece <> '' then
    vTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , lStNumeroPiece);
   if lStLibelle <> '' then
    vTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , lStLibelle);
   vTOBLigneEcr.PutValue( FStPrefixe + '_DEBIT'        , lRdDebit);
   vTOBLigneEcr.PutValue( FStPrefixe + '_CREDIT'       , lRdCredit);

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
function  TZGuideTreso.CalculeGuide : boolean;
var
 lTOBLigneEcr : TOB;
begin

  result       := false;

  try

  FInNumLigne  := 1;
  // on se place sur la première ligne du guide
  FCurrentTOB  := FTOB.FindFirst(['EG_NUMLIGNE','EG_GUIDE'],[1,FStEG_GUIDE],false);

  // si le taux de tva est nulle, on recherche par rapport a celui du compte
  if ( StGENERAL <> '' ) and ( ( FTabRech[TColTVA].Value = '' ) or ( FTabRech[TColTVA].Value = '0' ) ) then
   FTabRech[TColTVA].Value := FloatToStr( ZCompte.GetTauxTVA( StGENERAL, Valeur(FTabRech[TColDebit].Value) ));


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
function  TZGuideTreso.RecalculGuide ( vTOBEcr : TOB ; vInNumLigne : integer)   : TOB;
var
 lTOBLigneEcr : TOB;
begin

  FInNumLigne := vInNumLigne;

  FTOBEcr := vTOBEcr;

  FTabRech[TColLibelle].Value        := vTOBEcr.Detail[0].GetValue( FStPrefixe + '_LIBELLE');
  FTabRech[TColGeneral].Value        := vTOBEcr.Detail[0].GetValue( FStPrefixe + '_GENERAL');

  FCurrentTOB := FTOB.FindFirst(['EG_NUMLIGNE','EG_GUIDE'],[1,FStEG_GUIDE],false);
  // on recherche le taux de tva si celui ci n'etait pas defini
  if ( FTabRech[TColTVA].Value = '' ) or ( FTabRech[TColTVA].Value = '0' ) then
   FTabRech[TColTVA].Value := FloatToStr( ZCompte.GetTauxTVA( StGENERAL, Valeur(FTabRech[TColDebit].Value) ));
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
Modifié le ... : 09/11/2001
Description .. : Retourne une tob ecriture genere ne fonction du compte 
Suite ........ : d'imputation :
Suite ........ : 1 ligne : compte d'imput
Suite ........ : 2 ligne : compte de tva du compte
Suite ........ : 3 ligne : compte de contrepartie
Mots clefs ... : 
*****************************************************************}
function TZGuideTreso.GenerationGuide : boolean;
var
 lTOBLigneEcr   : TOB;
 lRdTva         : double;
 lRdVal         : double;
begin

  //FTOBEcr      := TOB.Create( '' , nil , -1);

  result       := false;

  // 1 ligne : compte d'imputation saisie
  lTOBLigneEcr := TOB.Create( FStLeNomTable , FTOBEcr , -1 );
  lTOBLigneEcr.AddChampSup( 'GUIDE',true );
  lTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , 1);
  lTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'      , FTabRech[TColGeneral].Value);
  lTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , FTabRech[TColNumeroPiece].Value);
  lTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , FTabRech[TColLibelle].Value);
  lTOBLigneEcr.PutValue( FStPrefixe + '_DATECOMPTABLE' ,StrToDateTime( FTabRech[TColDate].Value));

  if ( ZCompte.IsTvaAutorise( FTabRech[TColGeneral].Value , 0 )) and
     ( FTabRech[TColTva].Value <> '') and ( FTabRech[TColTva].Value <> '0') then
       begin
        // 2 ligne : compte de tva
        lTOBLigneEcr := TOB.Create( FStLeNomTable , FTOBEcr , -1);
        lTOBLigneEcr.AddChampSup( 'GUIDE',true );
        lTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , 2);
        lTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'      , StCompteTVA);
        lTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , FTabRech[TColNumeroPiece].Value);
        lTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , FTabRech[TColLibelle].Value);
        lTOBLigneEcr.PutValue( FStPrefixe + '_DATECOMPTABLE' , StrToDateTime(FTabRech[TColDate].Value));

        // 3 ligne : compte de contrepartie
        lTOBLigneEcr := TOB.Create( FStLeNomTable , FTOBEcr , -1);
        lTOBLigneEcr.AddChampSup( 'GUIDE',true );
        lTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , 3);
        lTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'      , StCompteContrepartie);
        lTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , FTabRech[TColNumeroPiece].Value);
        lTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , FTabRech[TColLibelle].Value);
        lTOBLigneEcr.PutValue( FStPrefixe + '_DATECOMPTABLE' , StrToDateTime(FTabRech[TColDate].Value));

       end
        else
         begin
          // 2 ligne : compte de contrepartie
          lTOBLigneEcr := TOB.Create( FStLeNomTable , FTOBEcr ,-1);
          lTOBLigneEcr.AddChampSup( 'GUIDE',true );
          lTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , 2);
          lTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'      , StCompteContrepartie);
          lTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'   , FTabRech[TColNumeroPiece].Value);
          lTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'      , FTabRech[TColLibelle].Value);
          lTOBLigneEcr.PutValue( FStPrefixe + '_DATECOMPTABLE' ,StrToDateTime(FTabRech[TColDate].Value));

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
             TOB(FTOBEcr.Detail[0]).PutValue( FStPrefixe + '_DEBIT'     ,  Valeur(FTabRech[TColDebit].Value ) );
             TOB(FTOBEcr.Detail[1]).PutValue( FStPrefixe + '_CREDIT'    ,  Valeur(FTabRech[TColDebit].Value ) );
         end
             else
              begin
                TOB(FTOBEcr.Detail[0]).PutValue( FStPrefixe + '_CREDIT' , Valeur(FTabRech[TColCredit].Value) );
                TOB(FTOBEcr.Detail[1]).PutValue( FStPrefixe + '_DEBIT'  , Valeur(FTabRech[TColCredit].Value) );
              end;
          end;

 result := true;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 09/11/2001
Description .. : Recherche d'un guide par rapport au libelle de la grille ou au
Suite ........ : code AFB de la ligne ( si le guide est parametre avec $AFB )
Suite ........ : le caractère '*' represente le caractère jocker.
Suite ........ : 
Mots clefs ... : 
*****************************************************************}
function TZGuideTreso.RechercheGuideParLibelle : string;
var
 i               : integer;
 lTOB            : TOB;
 lTOBEnteteGuide : TOB;
 lTOBLigneGuide  : TOB;
 lStCritRech     : string;
 lStTexteRech    : string;
 lPos            : integer;
 lInPosCrit      : integer;
begin

 lTOBEnteteGuide := TOB.Create('',nil,-1);
 lTOBLigneGuide  := nil;
 result          := '';
 lPos            := 0; // flag de position du caractères jocker

 try

  for i := 0 to FListeLibelle.Count - 1 do
   begin
    // on sauvegarde les critères de recherche
    lStCritRech  := FListeLibelle.Strings[i];
    lStTexteRech := FListeLibelle.Strings[i];

    // suppression des caractères de recherche dans la chaine de recherche
    if Copy( FListeLibelle.Strings[i], 1 , 1) = '*' then
     begin
      lStCritRech := Copy(lStCritRech, 2, length( lStCritRech ));
      lPos        := 1; // on a trouver une fois le jocker
     end;

    if Copy( FListeLibelle.Strings[i], length(FListeLibelle.Strings[i]) , 1) = '*' then
     begin
      lStCritRech := Copy(lStCritRech, 1, length( lStCritRech ) - 1);
      lPos        := lPos + 1; // on a trouver deux fois le jocker
     end; // if

    // place du critere de recherche dans la grille
    lInPosCrit := pos(UpperCase( lStCritRech ),UpperCase(FTabRech[TColLibelle].Value)) + length(UpperCase(lStCritRech)) - 1 ;

    if ( pos('*',lStTexteRech) > 1 ) and ( pos( UpperCase( lStCritRech ) , FTabRech[TColLibelle].Value) = 1 )  or // cas TPE*
       ( pos('*',lStTexteRech) = 1 ) and ( lInPosCrit = ( length(FTabRech[TColLibelle].Value) ) )              or // cas *TPE
       ( pos('*',lStTexteRech) = 0 ) and ( UpperCase( lStCritRech ) = FTabRech[TColLibelle].Value  )           or // cas TPE
       ( lPos                  = 2 ) and ( pos( UpperCase( lStCritRech ) , FTabRech[TColLibelle].Value ) > 0 ) or // cas *TPE*
       ( pos( FListeLibelle.Strings[i] , cStIdCodeAFB + FTabRech[TColCODEAFB].Value ) > 0 ) then                  // code AFB
      begin

       lTOB := TOB( FListeLibelle.Objects[i] );

       FirstGuide(VarToStr(lTOB.GetValue('EG_GUIDE')));

       if VerifSensDuGuide(lTOB) then
        begin
         lTOBLigneGuide  := TOB.Create('ECRGUI',lTOBEnteteGuide,-1);
         lTOBLigneGuide.PutValue('EG_GUIDE'      , lTOB.GetValue('EG_GUIDE'));
         lTOBLigneGuide.PutValue('EG_LIBELLE'    , lTOB.GetValue('EG_LIBELLE'));
         lTOBLigneGuide.PutValue('EG_REFINTERNE' , lTOB.GetValue('GU_LIBELLE'));
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
     end; // if

  FTypeRecherche := FTypeRecherche - [TColLibelle];

 finally
  if assigned(lTOBEnteteGuide) then lTOBEnteteGuide.Free;
 end; // try

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /    
Description .. : Recherche d'un guide par rapport au credit de la ligne
Mots clefs ... : 
*****************************************************************}
function TZGuideTreso.RechercheGuideParCredit : string;
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
  if pos( FListeMontant.Strings[i] ,UpperCase(FTabRech[TColCredit].Value) ) > 0 then
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
Modifié le ... :   /  /    
Description .. : Recherche d'un guide par rapport au debit de la ligne
Mots clefs ... : 
*****************************************************************}
function TZGuideTreso.RechercheGuideParDebit : string;
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
  if pos( FListeMontant.Strings[i] ,UpperCase(FTabRech[TColDebit].Value) ) > 0 then
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
function TZGuideTreso.RechercheGuideParImputation : string;
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
function TZGuideTreso.RechercheGuideTotal : string;
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
function TZGuideTreso.RechercheGuideEnBase( vStCredit,
                                              vStDebit,
                                              vStLibelle,
                                              vStGeneral,
                                              vStTVA,
                                              vStDate,
                                              vStNumeroPiece,
                                              vStPiece,
                                              vStCodeAFB : string;
                                              vTOBResult : TOB ;
                                              G : THGrid = nil
                                              ) : boolean;
var
 lValue : string;
 i      : integer;
begin

 result         := false;

 if ( ZCompte = nil ) or ( vTOBResult = nil )  then exit;

 // récuperation des valeurs de la grille
 FGrilleReleve  := G;

 FTOBEcr        := vTOBResult;

 FTypeRecherche := [TColGeneral,TColLibelle,TColDebit,TColCredit];  // ré-initialisation des options de recherche

 if vStDebit <> '0' then
  FTabRech[TColCredit].Value        := vStDebit
   else
    FTabRech[TColCredit].Value      := '';

 if vStCredit <> '0' then
  FTabRech[TColDebit].Value        := vStCredit
   else
    FTabRech[TColDebit].Value      := '';

 if IsNumeric(vStTVA) and (StrToFloat(vStTVA) = 0 ) then
  vStTVA := '0';

 FTabRech[TColLibelle].Value        := vStLibelle;
 FTabRech[TColGeneral].Value        := vStGeneral;
 FTabRech[TColTVA].Value            := vStTVA;
 FTabRech[TColDate].Value           := vStDate;
 FTabRech[TColNumeroPiece].Value    := vStNumeroPiece;
 FTabRech[TColPiece].Value          := vStPiece;
 FTabRech[TColCodeAFB].Value        := vStCodeAFB;

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
function TZGuideTreso.IsColStop(ARow, ACol : integer): boolean;
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
function TZGuideTreso.PossedeArret : boolean;
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
Modifié le ... :   /  /    
Description .. : Contrôle du sens du guide
Mots clefs ... : 
*****************************************************************}
function TZGuideTreso.VerifSensDuGuide( Value : TOB ) : boolean;
begin
 // contrôle du sens du guide
 result := ( ( FTabRech[TColCredit].Value = '' ) and ( Value.GetValue('SENS') = '0' ) ) or
           ( ( FTabRech[TColDebit].Value = '' ) and ( Value.GetValue('SENS') = '1' ) );

 // on contrôle que le compte d'imputation du guide est le même que celui de la grille de releve
 if result and ( FTabRech[TColGeneral].Value <> '' ) then
  result := ( StGENERAL = FTabRech[TColGeneral].Value );

 // si la derniere ligne contient un montant, il doit être egal au solde du guide
 if result and isNumeric( FTabRech[TColCredit].Value ) and isNumeric( Value.GetValue('SOLDE') ) then
  result := FTabRech[TColCredit].Value = Value.GetValue('SOLDE');

 // si la derniere ligne contient un montant, il doit être egal au solde du guide
 if result and isNumeric( FTabRech[TColDebit].Value ) and isNumeric( Value.GetValue('SOLDE') ) then
  result := FTabRech[TColDebit].Value = Value.GetValue('SOLDE');

end;


function TZGuideTreso.EstVide : boolean;
begin
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
procedure TZGuideTreso.SetLeNomTable( Value : string );
begin
 FStPrefixe    := TableToPrefixe ( Value );
 FStLeNomTable := Value;
end;

procedure TZGuideTreso.NotifyError( vMessage, vDelphiMessage , vFunctionName : string);
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
Modifié le ... :   /  /    
Description .. : Suppression de tous les guides enregsitrer en memoire
Mots clefs ... : 
*****************************************************************}
procedure TZGuideTreso.ClearGuide;
begin
 FTOB.ClearDetail;
 FListeLibelle.Clear;
 FListeMontant.Clear;
 FListeGeneral.Clear;
end;

end.
