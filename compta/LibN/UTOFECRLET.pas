{***********UNITE*************************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 15/11/2001
Modifié le ... : 18/09/2002
Description .. : TOF des ecritures simplifies
Mots clefs ... : TOF;ECRLET
JP 14/02/07 : Création d'un ancêtre commune à TOF_ECRLET et à CPSAISIEPIECE_TOF
              dans ULibPieceCompta : A chacune des ces deux classes correspond
              une fiche, mais ces deux fiches doivent rester identiques =>
              TOUTE MODIFICATION DANS L'UNE DES FICHES DOIT ÊTRE RÉPERCUTÉ DANS L'AUTRE !!
*****************************************************************}
Unit UTOFECRLET ;

Interface

Uses    utobdebug,
StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF , buttons, Dialogs,
     {$IFDEF EAGLCLIENT}
     MaineAGL,
     {$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF},
     FE_Main, 
     {$ENDIF}
     // ALG
     UTOB,
     Ent1,     // pourle VH
     LettUtil,
     paramsoc,
     AGLInit, // pour TheData

     SaisUtil,  // pour le GetNewNumJal
     UtilSais, // pour le MajSoldeEcrituresTOB
     LookUp,
     Vierge,
     Windows,
     HTB97,
     Graphics,
     ULibEcriture,
     uLibAnalytique,  // pour AlloueAxe
     SaisComm,        // pour les tsmDevise...
     ed_tools,
     extctrls,
     Messages,
     HPanel,
     UtilPgi,
     ULibEcrSimplifiee, {Contient la classe ancêtre de la tof}
     {$IFDEF COMPTA}
     SaisComp ,
     {$ENDIF}
     ULibWindows
     ;

Type
  {JP 14/02/07 : création d'une classe ancêtre commune aux deux TOF, car les deux fiches
                 doivent rester IDENTIQUES}
  TOF_ECRLET = Class (TofAncetreEcr)
   private
    FIFL                  : RINFOSLETT;
    FRLETTR               : RLETTR;
    FInLastNumero         : integer;
    FDtLastDateComptable  : TDateTime;
    FStLastJournal        : string;
    FStLastLibelle        : string;
    FStLastGeneral        : string;
    FStRetour             : string;
    FTOBEcrLett           : TOB;  // tob contenant l'ensemble des lignes de l'ecriture de departs
    FTOBEcrGene           : TOB;
    FTOBLigne             : TOB ; // tob poinatant sur la ligne de contrepartie
    FDevise               : RDevise;
    FR                    : TAGLLanceFiche_TOFECRLET;
    FBoMvtAuDebit         : boolean;
    FRechCompte           : TRechercheGrille;
    FEcr                  : TTOBEcriture;
    FBoLettrable          : boolean;
    FBoPointable          : boolean;
    FBoCollectif          : boolean;
    FBoVentilable         : boolean;
    FBoBordereau          : boolean;
    FStGeneral1           : string;
    FStGeneral2           : string;
    FStGeneral3           : string;
    FDtDatePaquetMax      : TDateTime;
    FDtDatePaquetMin      : TDateTime;
    FStNatureJal          : string ;
    FMessCompta           : TMessageCompta ;
    FStCompteTVA          : string ;
    procedure BValiderClick(Sender: TObject); // Bouton Valider
    procedure BFermeClick(Sender: TObject); // Bouton Annuler
    procedure EdtJournalExit(Sender: TObject);
    procedure EdtGeneralExit(Sender: TObject);
    procedure EdtAuxiExit(Sender: TObject);
    procedure EdtDateComptableChange( Sender : TObject );
    procedure EdtNatureExit(Sender: TObject);

    function  Generation              : boolean ;
    function  InitControl             : boolean; // initialise les grilles
    function  CreateControl           : boolean;

    function  IsValidEcr           : boolean; // validate toutes les info saisies par l'utilisateurs
    function  IsValidJournal( Parle: boolean = false ): boolean;
    function  IsValidGeneral( Parle : boolean=true ): boolean;
    function  IsValidNature( Parle : boolean=false ): boolean;
    function  IsValidAuxi(AutoriseCellExit : boolean = false ; Parle : boolean=false ): boolean;
    function  IsValidEtablissement( Parle : boolean=false ): boolean;
    function  IsValidNumeroPiece( Parle : boolean=false ): boolean;
    procedure RemplirLesLignes ;
    {$IFDEF COMPTA}
//    procedure GereCutoff ( vTOB : TOB ) ;
    {$ENDIF}
    procedure OnError(sender: TObject; Error: TRecError);

   protected
     procedure EdtJournalElipsisClick(Sender: TObject); override;
     function  AssignEvent             : boolean; override;// Affecte les evenements au contrôle
     function  IsValidDateComptable( Parle : boolean=false ): boolean; override;
     procedure EdtGeneralElipsisClick(Sender: TObject); 
     procedure EdtDateComptableExit(Sender: TObject); override;
     procedure BtnVoirClick(Sender: TObject); override;
     procedure EdtFolioExit(Sender: TObject); override;

   public
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

 TEcrRegul = Class(TObjetCompta)
  private
   FTOBEcrGene     : TOB;   // TOB des ecritures de regul pour le mode PCL
   procedure RecupereLigneBor(vTOB: TOB);
   procedure RecupereLignePiece(vTOB: TOB);
  public
   constructor Create( vInfoEcr : TInfoEcriture ) ; override ;
   destructor  Destroy ; override;

   function  Save : boolean;
   procedure RecupereLigne( vTOB : TOB );
   procedure Delete( vStCodeLettrage : string ) ;
//   procedure ALettrerAvec( vTOB : TOB ) ;

   property TOBEcrGene : TOB read FTOBEcrGene write FTOBEcrGene;

 end;


function CPLanceFiche_ECRLET( const vTOBEcr : TOB ; var vTOBResult : TOB ; vStParam : string = '' ; vInfoEcr : TInfoEcriture = nil ) : string ;
//function CCreateTOBECRLETFromQuery( vQ : TQuery ) : TOB ;
//function CCreateTOBECRLETFromTOB( vTOB : TOB ) : TOB ;
function CSaveEcr( vTOBSimpl , vTOBEcr : TOB  ) : boolean ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ZCompte;

const
 cStJournalInexistant      = 'Ce journal n''existe pas !';
 cStJournalPasMultiDevise  = 'Ce journal n''est pas multidevise !';
 cStCompteContrepartie     = 'le compte doit être égal au compte de contrepartie du compte !';
 cStCompteInterdit         = 'Ce compte est interdit pour ce journal!';
 cStPasSurCompteBilan      = 'Vous ne pouvez pas saisir sur le compte d''ouverture de bilan';
 cStManqueEtablissement    = 'Il manque le code établissement';
 cStNumeroPieceIncorrecte  = 'Le numéro de folio est incorrect';
 cStNumeroPieceEnDev       = 'Vous ne pouvez pas faire d''écriture compte à compte sur un journal libre en devise';
 cStNumeroPieceValide      = 'La pièce est validée' ;
 cStCompteInterditPourNat  = 'Ce compte général est interdit pour cette nature de pièce';
 cStPasCompteRegul         = 'Ce compte n''est pas un compte de régularisation';


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 07/11/2006
Modifié le ... :   /  /    
Description .. : - LG - 07/06/2006 - on bloque la generation d'ecriture 
Suite ........ : simplfie en devise
Mots clefs ... : 
*****************************************************************}
function CPLanceFiche_ECRLET(  const vTOBEcr : TOB ; var vTOBResult : TOB ; vStParam : string = '' ; vInfoEcr : TInfoEcriture = nil ) : string ;
var
 lBoCreateInfo                : boolean ;
 lAGLLanceFiche_TOFECRLET     : TAGLLanceFiche_TOFECRLET ;
 i                            : integer ;
begin

 if not ExJaiLeDroitConcept(TConcept(ccSaisEcritures), true) then
  begin
   Result := '' ;
   Exit ;
  end ;

 if ( vTOBEcr.Detail <> nil ) then
  for i := 0 to vTOBEcr.Detail.Count - 1 do
   if ( vTOBEcr.Detail[i].GetString('E_DEVISE') <> V_PGI.DevisePivot ) then
    begin
     PGIInfo('Vous ne pouvez pas faire d''écriture compte à compte en devise') ;
     exit ;
    end ;

 if vStParam = '' then vStParam:='AL6' ;

 if vInfoEcr = nil then
  begin
   vInfoEcr      := TInfoEcriture.Create ;
   lBoCreateInfo := true ;
  end
   else
    lBoCreateInfo := false ;

 lAGLLanceFiche_TOFECRLET            := TAGLLanceFiche_TOFECRLET.Create ;
 lAGLLanceFiche_TOFECRLET.Info       := vInfoEcr ;
 lAGLLanceFiche_TOFECRLET.TOBEcr     := vTOBEcr ;
 lAGLLanceFiche_TOFECRLET.TOBResult  := vTOBResult ;

 TheData                             := lAGLLanceFiche_TOFECRLET ;

 try

  result := AGLLanceFiche('CP','CECRSIMPL','','',vStParam) ;

 finally
  if lBoCreateInfo                      then vInfoEcr.Free ;
  if assigned(lAGLLanceFiche_TOFECRLET) then lAGLLanceFiche_TOFECRLET.Free;
  TheData                         := nil ;
 end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 12/04/2007
Modifié le ... : 02/05/2007
Description .. : - FB 19801 - LG - 12/04/2007 - reecriture de la fct pour 
Suite ........ : l'ameliorer ( le bug etait ds ULibEcriture - CalculMonoEche )
Suite ........ : - LG - 02/05/2007 - FB 20201 - l'ajout ds un bordereau 
Suite ........ : existant plantait
Mots clefs ... : 
*****************************************************************}
function CSaveEcr( vTOBSimpl , vTOBEcr : TOB ) : boolean ;
var
 i        : integer ;
 lTOB     : TOB ;
 lTOBLett : TOB ;
begin

 result := false ;
 if not CEnregistreSaisie( vTOBSimpl ,(vTOBSimpl.Detail[0].GetValue('E_MODESAISIE') = '-') ,true ) then exit ;
 result:=true ;

 if vTOBEcr = nil then exit ;

 lTOBLett := TOB.Create('', nil , -1) ;

 for i := 0 to vTOBEcr.Detail.Count - 1 do
  begin
   lTOB     := TOB.Create('', nil , -1) ;
   lTOB.Dupliquer(vTOBEcr.Detail[i],false,true) ;
   lTOB.ChangeParent(lTOBLett,-1) ;
  end ;

 try

 // la 2 ligne est la ligne à lettrer
 vTOBSimpl.Detail[vTOBSimpl.Detail.Count - 1].ChangeParent(lTOBLett,-1 ) ;

 {$IFNDEF EAGLSERVER}
 {$IFDEF COMPTA}
  CLettrageTOB(lTOBLett) ;
 {$ENDIF}
 {$ENDIF}

 finally
  lTOBLett.Free ;
 end ;

end;


procedure TOF_ECRLET.OnArgument (S : String ) ;
begin

 FStRetour         := '0';
 FIFL              := ChargeInfosLett ;
 FIFL.ChoixValid   := ReadTokenST(S) ;
 FR                := TAGLLanceFiche_TOFECRLET(TheData);
 // recuperation des parametres
 if FIFL.ChoixValid = 'AL3' then
  begin // on ne recupere les info que pour les ecritures simplifies
   FInLastNumero                        := StrToInt(ReadTokenST(S));
   if (S<>'') then FDtLastDateComptable := StrToDateTime(ReadTokenST(S)) else FDtLastDateComptable := iDate1900 ;
   if (S<>'') then FStLastJournal       := ReadTokenST(S);
   if (S<>'') then FStLastLibelle       := ReadTokenST(S);
   if (S<>'') then FStLastGeneral       := ReadTokenST(S);
  end; // if

 {FQ19053 11/10/2007 YMO Journal défini dans les param sociétés}
 If FIFL.ChoixValid = 'AL6' then //Passage d'une écriture de compte à compte
    FStLastJournal:=GetParamSocSecur('SO_CPJALCOMPENSATION', False);

 if not GetControlTOF then exit;
 if not CreateControl then exit;

 if FIFL.ChoixValid = 'AL1' then
  begin
   FStLastJournal := FIFL.Journal;
   FStLastLibelle := FIFL.Libelle;
  end; // if

 if not InitControl   then exit;
 if not AssignEvent   then exit;

 Inherited ;

end ;

procedure TOF_ECRLET.OnClose ;
begin

 try
  if assigned(FRechCompte)   then FRechCompte.Free;
  if assigned(FMessCompta)   then FMessCompta.Free;
  if assigned(FEcr)          then FEcr.Free;
 finally
  FRechCompte := nil ;
  FMessCompta := nil ;
  FEcr        := nil ;
 end;

 Inherited ;

end ;
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... :   /  /
Description .. : Recuperation des controles de la fiche
Mots clefs ... :
*****************************************************************}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 05/07/2006
Description .. : Initialisation de tous les contrôles de la grille
Suite ........ : 
Suite ........ : -05/06/2002 - affichage du libelle de l'etablissement ( bug
Suite ........ : 10112 )
Suite ........ : 
Suite ........ : - 17/06/2002 - on test si c'est un lettrage en devise pour
Suite ........ : des ecritures en devise
Suite ........ : 
Suite ........ : - FB 10763 - on ne recupere pas le compte de regul
Suite ........ : - FB13273 - LG - 05/07/2006 - on reprends le libelle de la 
Suite ........ : dernier ligne de l'ecriture de depatr
Mots clefs ... : 
*****************************************************************}
function TOF_ECRLET.InitControl: boolean;
var
 i          : integer;
begin

 // recuperation des info de lettrage
 FRLETTR.LettrageDevise                        := FTOBLigne.GetValue('E_LETTRAGEDEV')='X' ;
 FRLETTR.RefLettrage                           := FTOBLigne.GetValue('E_REFLETTRAGE') ;
 // paramètrage de la grille
 G.ListeParam                                  := 'LETTECR';
 G.ColWidths[G.ColCount-1]                     := - 1; // la grille a une ligne en trop -> suppression par code
 EdtEtablissement.ElipsisButton                := false; // on ne peut plus modifier l'etablissement
 EdtEtablissement.Enabled                      := false;
 EdtTva.enabled                                := false; //(FR.IFL.ChoixValid = 'AL3') and (GetParamSocSecur('SO_REGIMEDEFAUT','')<>'');
 EdtAuxiliaire.Enabled                         := FIFL.ChoixValid <> 'AL1'; // pas de tva pour les ecritures de regul
 lblAuxiliaire.Enabled                         := FIFL.ChoixValid <> 'AL1';
 THLabel(GetControl('TE_TVA')).Enabled         := EdtTva.Enabled;
 EdtAuxiliaire.DataType                        := '' ;

// ColorControl(Tedit(EdtTva));
// ColorControl(Tedit(EdtAuxiliaire));

 // recherche de la plus grande date comptable de l'ensemble d'ecriture
 FDtDatePaquetMax                              := FTOBLigne.GetValue('E_DATECOMPTABLE');
 FDtDatePaquetMin                              := FTOBLigne.GetValue('E_DATECOMPTABLE');
 for i := 0 to FTOBEcrLett.Detail.Count - 1 do
  begin
   if FTOBEcrLett.Detail[i].GetValue('E_DATECOMPTABLE') > FDtDatePaquetMax then
    FDtDatePaquetMax:=FTOBEcrLett.Detail[i].GetValue('E_DATECOMPTABLE');
   if FTOBEcrLett.Detail[i].GetValue('E_DATECOMPTABLE') < FDtDatePaquetMin then
    FDtDatePaquetMin:=FTOBEcrLett.Detail[i].GetValue('E_DATECOMPTABLE');
  end; // for

 // recherche des info sur la devise
 if FRLETTR.LettrageDevise then
  FRechCompte.Info.Devise.Load( [FTOBLigne.GetValue('E_DEVISE')] )
   else
     FRechCompte.Info.Devise.Load( [V_PGI.DevisePivot] ) ;

 // contrôle de la coherence de la date avec la derniere periode close
 if FDtDatePaquetMin < (GetParamSocSecur('SO_DATECLOTUREPER',iDate1900)+1) then
  FDtDatePaquetMin:=(GetParamSocSecur('SO_DATECLOTUREPER',iDate1900)+1);

 EdtDateComptable.Text                         := DateToStr(FDtDatePaquetMax);

 if ( FStLastLibelle <> '' ) then
  EdtLibelle.Text := FStLastLibelle
   else
    EdtLibelle.Text := FTOBEcrLett.Detail[FTOBEcrLett.Detail.count-1].GetString('E_LIBELLE') ;

 FBoMvtAuDebit      := FTOBLigne.GetValue('ECARTREGUL') > 0;

 if FIFL.ChoixValid = 'AL1' then
  begin // on ne recupere les info que pour les ecritures de regul
   if not FBoMvtAuDebit then
//     FStLastGeneral := FIFL.GeneralDebit  CA - 06/02/2003 - FQ 10763
      FStLastGeneral := FIFL.GeneralCredit
      else
//       FStLastGeneral := FIFL.GeneralCredit; CA - 06/02/2003 - FQ 10763
        FStLastGeneral := FIFL.GeneralDebit;
  end; // if

 // affichage du libelle de l'etabllissement ( bug 10112 )
 EdtEtablissement.Text := RechDom('TTETABLISSEMENT',FTOBLigne.GetString('E_ETABLISSEMENT'),false);

 if FStLastJournal <> '' then
  begin
   EdtJournal.Text := FStLastJournal;
   EdtJournalExit(nil);
  end; // if

 G.Enabled                := false;
 EdtNaturePiece.Text      := 'OD';

 if FIFL.ChoixValid = 'AL1' then
  begin
   EdtNaturePiece.Enabled := false;
  // ColorControl(Tedit(EdtNaturePiece));
  end; // if

 EdtJournal.SetFocus;

 EdtGeneral.MaxLength := VH^.Cpta[fbGene].Lg ; // longeur d'un compte general

 P2.Visible                                  := false ;
 TTabSheet(GetControl('PDETAIL')).visible    := false ;
 TTabSheet(GetControl('PDETAIL')).TabVisible := false ;
 Ecran.Height         := 320;
 SetControlText('LGEN', '' ) ;
 SetControlText('LAUX', '' ) ;

 result               := true;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... :   /  /
Description .. : Affectation des evenements
Mots clefs ... :
*****************************************************************}
function TOF_ECRLET.AssignEvent : boolean;
begin
  Result := inherited AssignEvent;
  if not Result then Exit;

  // affectation des evenements
 EdtJournal.OnExit             := EdtJournalExit;
 EdtGeneral.OnExit             := EdtGeneralExit;
 EdtDateComptable.OnChange     := EdtDateComptableChange;
 EdtNaturePiece.OnExit         := EdtNatureExit;
 EdtAuxiliaire.OnExit          := EdtAuxiExit;

 TToolbarButton97(GetControl('BValider')).Onclick  := BValiderClick;
 TToolbarButton97(GetControl('BFerme')).Onclick    := BFermeClick;

 TFVierge(Ecran).HMTrad.ResizeGridColumns(G);

 result                        := true;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... :   /  /
Description .. : Creation des objets de gestion de la TOF
Mots clefs ... :
*****************************************************************}
function TOF_ECRLET.CreateControl: boolean;
begin
 // creation des objets
 FTOBEcrLett                    := FR.TOBEcr;
 FTOBLigne                      := FTOBEcrLett.Detail[0] ;
 FTOBEcrGene                    := FR.TOBResult ;
 // Objet de recherche des comptes
 FRechCompte                    := TRechercheGrille.Create(TInfoEcriture(FR.Info)) ;
 FRechCompte.OnError            := OnError ;
 // objet de validation des ecritures
 FEcr                           := TTOBEcriture.Create('ECRITURE',nil,-1) ;
 FEcr.Info                      := FRechCompte.Info ;
 FEcr.OnError                   := OnError ;
 // Objet d'affichage des messages
 FMessCompta                    := TMessageCompta.Create(Ecran.Caption) ;
 result                         := true;
end;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 11/10/2007
Modifié le ... :   /  /    
Description .. : LG - 11/07/2007 - FB 21616 - les lignes de tva etaient ss 
Suite ........ : montant pour des regul = 0.1
Mots clefs ... : 
*****************************************************************}
function TOF_ECRLET.Generation : boolean ;
var
 i       : integer;
 lTError : TRecError ;
 lTOB    : TOB ;
begin

 result := false ;

 // controle des champs saisies ds l'interface
 if not IsValidEcr then exit;


 RemplirLesLignes ;

 i := 0 ;

 // FB 21616 - on supprime les lignes vides, genere pour des montant de regul trop petit. le monatnt de tva est agale a zero
 while i < FTOBEcrGene.Detail.Count - 1 do
  begin
   lTOB := FTOBEcrGene.Detail[i] ;
   if ( lTOB.GetDouble('E_DEBIT') = 0 ) and ( lTOB.GetDouble('E_CREDIT') = 0 ) then
    lTOB.Free
     else
      inc(i) ;
  end; // while


 // controle globale de la piece
 lTError := CIsValidSaisiePiece( FTOBEcrGene , FRechCompte.Info , true) ;
 if ( lTError.RC_Error <> RC_PASERREUR ) then
  begin
   OnError(nil,lTError) ;
   exit ;
  end ;

 result := true ;

 G.VidePile(false);
 for i := 0 to FTOBEcrGene.Detail.Count - 1 do
  begin
   G.Row := G.RowCount-1;
   CTOBVersTHGrid(G,FTOBEcrGene.Detail[i],FRLETTR.LettrageDevise);
   G.RowCount:=G.RowCount+1 ;
  end; // for

end;

{$IFDEF COMPTA}
(*
procedure TOF_ECRLET.GereCutoff ( vTOB : TOB ) ;
var
 lRComp          : R_COMP ;
 lTOBCompl       : TOB ;
 lStCpt          : String ;
 lAction         : TActionFiche ;
 bModifBlocNote  : Boolean ;
begin


  lStCpt := vTOB.GetValue( 'E_GENERAL' ) ;
  if (lStCpt = '') or not ( FRechCompte.Info.LoadCompte( lStCpt ) )  then Exit ;

  if FRechCompte.Info.Compte_GetValue('G_CUTOFF') <> 'X' then
   exit ;

  lAction            := taCreat ;
  lRComp.StComporte  := '' ;
  lRComp.StLibre     := '' ;
  lRComp.Conso       := False ;
  lRComp.DateC       := vTOB.GetValue( 'E_GENERAL' ) ;
  lRComp.Attributs   := true ;
  lRComp.MemoComp    := nil ; // MemoComp := FindField('SC_ATTRCOMP')).AsString
  lRComp.Origine     := 0 ;
  lRComp.StLibre     := '---CUTXXXXXXXXXXXXXXXXXXXXXXXX' ;
  lRComp.StComporte  := '--XXXXXXXX' ;
  bModifBlocNote     := true ;
  lRComp.CutOffPer   := FRechCompte.Info.Compte.GetValue('G_CUTOFFPERIODE') ;
  lRComp.CutOffEchue := FRechCompte.Info.Compte.GetValue('G_CUTOFFECHUE') ;
  lTOBCompl          := CGetTOBCompl(vTOB) ;
  if lTOBCompl = nil then
   lTOBCompl := CCreateTOBCompl(vTOB,nil) ;
  lTOBCompl.PutValue('EC_DATECOMPTABLE', vTOB.GetValue('E_DATECOMPTABLE')) ;

  CCalculDateCutOff( lTOBCompl, FRechCompte.Info.Compte_GetValue('G_CUTOFFPERIODE') , FRechCompte.Info.Compte_GetValue('G_CUTOFFECHUE') ) ;
  lRComp.TOBCompl := lTOBCompl ;

  if not SaisieComplement(TOBM(vTOB), EcrGen, lAction, bModifBlocNote, lRComp, False, True) then
   CFreeTOBCompl(vTOB) ;

end;
*)
{$ENDIF}


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 11/06/2002
Description .. : Generation de l'ecriture de regul ou simplifie
Suite ........ :
Suite ........ : - LG - 25/04/2002 - affectation du code exercice par
Suite ........ : rapport à la date comptable
Suite ........ :
Suite ........ : - 07/06/2002 - bug 10134 - recuperation du champs
Suite ........ : reflettrage
Suite ........ :
Suite ........ : - 10/06/2002 - FB 10110 - affectation du champs E_RIB et
Suite ........ : E_MODEPAIE
Description .. : - 02/04/2003 - FB 12234 - correction de l'affectation de la
Suite ........ : date paquet max et min
Description .. : - 11/04/2003 - FB 12269 - ventilation de tous les comptes
*****************************************************************}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 24/08/2004
Modifié le ... : 04/05/2006
Description .. : - LG - FB 13834 - 24/08/2004 - affectation des champs
Suite ........ : e_numeche ( avec la fct cgetech )
Suite ........ : - LG - FB 13678 - suppression des fct de calcul de la tva,
Suite ........ : deplacer ds le isValid du general
Suite ........ : - LG - FB 13834 - 11/10/2005 - la date echeance etait
Suite ........ : efface par initialisation du pointage
Suite ........ : - LG - FB 17625 - 04/05/206 - Lors de la génération d'une
Suite ........ : écriture simpifiée d'od de lettrage le champ e_qualifpiece est
Suite ........ : alimenté avec la valeur du code utilisateur connecté
Mots clefs ... :
*****************************************************************}
procedure TOF_ECRLET.RemplirLesLignes ;
var
 lTOBLigneEcr      : TOBM;
 C                 : double;
 lRdVal            : double;
 lRdTva            : double;
 lStCompte         : string;
 lInIndex          : integer;
 j, k              : integer;
 ValMax            : double;
 RegTva, TxTva     : string;
 Eche              : array[1..4] of double;
 EcheD             : double;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/04/2007
Modifié le ... : 20/11/2007
Description .. : - 23/04/2007 - FB 19569 - la date de creation n'etait pas a
Suite ........ : jour
Suite ........ : - 20/11/2007 - FB 21840 - on ne duplique pas le champs 
Suite ........ : e_exporte. on lui remets la valeur par defaut '---'
Mots clefs ... : 
*****************************************************************}
 procedure RempliLigne;
  begin
    CDupliquerTOBEcr( FTOBLigne , lTOBLigneEcr);
    if FBoBordereau then
     lTOBLigneEcr.PutValue( 'E_NUMEROPIECE'     , StrToInt( EdtFolio.Text ));
    lTOBLigneEcr.PutValue( 'E_EXPORTE'         , '---') ;
    lTOBLigneEcr.PutValue( 'E_ETABLISSEMENT'   , FTOBLigne.GetValue('E_ETABLISSEMENT'));
    lTOBLigneEcr.PutValue( 'E_NATUREPIECE'     , EdtNaturePiece.Text);
    lTOBLigneEcr.PutValue( 'E_QUALIFPIECE'     , 'N' );
    lTOBLigneEcr.PutValue( 'E_LIBELLE'         , EdtLibelle.Text);
    lTOBLigneEcr.PutValue( 'E_MODESAISIE'      , FRechCompte.Info.Journal.GetValue('J_MODESAISIE'));
    lTOBLigneEcr.PutValue( 'E_REGIMETVA'       , FTOBLigne.GetValue('E_REGIMETVA'));
    lTOBLigneEcr.PutValue( 'E_JOURNAL'         , EdtJournal.Text);
    lTOBLigneEcr.PutValue( 'E_LIBELLE'         , EdtLibelle.Text);
    lTOBLigneEcr.PutValue( 'E_REFLETTRAGE'     , FRLETTR.RefLettrage);
    lTOBLigneEcr.PutValue( 'E_VALIDE'          , '-');
    lTOBLigneEcr.PutValue( 'E_TYPEMVT'         , 'DIV');
    lTOBLigneEcr.PutValue( 'E_ECRANOUVEAU'     , 'N') ;
    lTOBLigneEcr.PutValue( 'E_NUMGROUPEECR'    , 1) ;
    lTOBLigneEcr.PutValue( 'E_DATECREATION'    , Now) ;
    lTOBLigneEcr.PutValue( 'E_IO'              , 'X');
    lTOBLigneEcr.DateComptable:=StrToDate( EdtDateComptable.Text );
    lTOBLigneEcr.PutValue( 'E_EXERCICE'        , QuelExo(EdtDateComptable.Text));
     if FIFL.ChoixValid = 'AL1' then
      lTOBLigneEcr.PutValue('E_QUALIFORIGINE'   , 'REG')
       else
        lTOBLigneEcr.PutValue('E_QUALIFORIGINE' , 'SIM');
     CAjouteChampsSuppLett(lTOBLigneEcr);

    if FIFL.ChoixValid = 'AL6' then //YMOO
    begin
      lTOBLigneEcr.PutValue( 'E_TVA'       , TxTva);
      lTOBLigneEcr.PutValue( 'E_REGIMETVA' , RegTva);
    end;
   end;
begin

 FTOBEcrGene.ClearDetail ;

 // calcul du taux par rapport a la date comptable
 FDevise                        := FRechCompte.Info.Devise.Dev;
 FDevise.Taux                   := GetTaux(FDevise.Code , FDevise.DateTaux, StrToDate( EdtDateComptable.Text ));

 // 1 ligne : compte d'imputation saisie
 lTOBLigneEcr := TOBM.Create(EcrGen,'',true,FTOBEcrGene);

 if FIFL.ChoixValid = 'AL6' then {YMO 10/2008 Reprise informations TVA}
 begin

   TxTva:='';
   RegTva:='';
   For k:=1 to 4 do Eche[k]:=0.0;
   ValMax:=0;

   For j:=0 to FTOBEcrLett.Detail.Count-1 do
   begin

     For k:=1 to 4 do
     begin
        Eche[k] := Eche[k]+FTOBEcrLett.Detail[j].GetValue('E_ECHEENC'+inttostr(k)) ;
        If FTOBEcrLett.Detail[j].GetValue('E_ECHEENC'+inttostr(k))>ValMax then
        begin
          TxTva:=FTOBEcrLett.Detail[j].GetValue('E_TVA');
          RegTva:=FTOBEcrLett.Detail[j].GetValue('E_REGIMETVA');
          ValMax:=FTOBEcrLett.Detail[j].GetValue('E_ECHEENC'+inttostr(k));
        end;
     end;

     EcheD := EcheD+FTOBEcrLett.Detail[j].GetValue('E_ECHEDEBIT') ;
   end;

   For k:=1 to 4 do
       lTOBLigneEcr.PutValue( 'E_ECHEENC'+inttostr(k), Eche[k]);

   lTOBLigneEcr.PutValue( 'E_ECHEDEBIT', EcheD);

 end;

 RempliLigne;
 lTOBLigneEcr.PutValue('E_NUMLIGNE'         , 1);
 lTOBLigneEcr.PutValue('E_GENERAL'          , EdtGeneral.Text);
 lTOBLigneEcr.PutValue('E_AUXILIAIRE'       , EdtAuxiliaire.Text);
 lTOBLigneEcr.PutValue('E_RIB'              , CGetRIB(lTOBLigneEcr));

  if FBoLettrable then
  CRemplirInfoLettrage(lTOBLigneEcr)
   else
    CSupprimerInfoLettrage(lTOBLigneEcr);

 if FBoPointable then
  CRemplirInfoPointage(lTOBLigneEcr) ;

 // il faut que la saisie de tva soit autorise sur le compte, que l'on est pas saisie 0 et que le compte de tva soit definie
 if ( edtTva.Enabled ) and ( Valeur(edtTva.text)<>0 ) and ( trim(FStCompteTVA) <> '' ) then
  begin
   // 2 ligne : compte de tva
   lTOBLigneEcr := TOBM.Create(EcrGen,'',true,FTOBEcrGene);
   RempliLigne;
   lTOBLigneEcr.PutValue( 'E_NUMLIGNE'        , 2);
   lTOBLigneEcr.PutValue( 'E_GENERAL'         , FStCompteTVA);
   lInIndex  := FRechCompte.Info.Compte.GetCompte(FStCompteTVA);

   if FRechCompte.Info.Compte.IsLettrable(lInIndex) then
    CRemplirInfoLettrage(lTOBLigneEcr)
     else
      if FRechCompte.Info.Compte.IsPointable(lInIndex) then
       CRemplirInfoPointage(lTOBLigneEcr);

   // 3 ligne : compte de depart
   lTOBLigneEcr := TOBM.Create(EcrGen,'',true,FTOBEcrGene);
   RempliLigne;
   CRemplirInfoLettrage(lTOBLigneEcr);
   lTOBLigneEcr.PutValue( 'E_DATEPAQUETMAX'   , FDtDatePaquetMax);
   lTOBLigneEcr.PutValue( 'E_DATEPAQUETMIN'   , FDtDatePaquetMin);
   lTOBLigneEcr.PutValue( 'E_NUMLIGNE'        , 3);
   lTOBLigneEcr.PutValue( 'E_GENERAL'         , FTOBLigne.GetValue('E_GENERAL'));
   CDupliquerInfoAux( FTOBLigne,lTOBLigneEcr );
   lTOBLigneEcr.PutValue( 'E_RIB'             , CGetRIB(lTOBLigneEcr));
   lTOBLigneEcr.PutValue( 'SELECTIONNER'      , 'X' );
   lTOBLigneEcr.PutValue( 'PASENREGISTRER'    , 'X' );
   {$IFDEF COMPTA}
   FRechCompte.Info.Load(lTOBLigneEcr.GetValue('E_GENERAL'),lTOBLigneEcr.GetValue('E_AUXILIAIRE'),EdtJournal.Text) ;
   if FRechCompte.Info.Compte.IsCollectif and ( FRechCompte.Info.GetString('T_NATUREAUXI') <> 'SAL' ) and ( length(trim( lTOBLigneEcr.GetValue('E_REGIMETVA') )) = 0 )   then
   CGetRegimeTVA ( lTOBLigneEcr , FRechCompte.Info ) ;
   //GereCutoff( lTOBLigneEcr ) ;
   {$ENDIF}
  end
   else
    begin
     // 2 ligne : compte de contrepartie
     lTOBLigneEcr := TOBM.Create(EcrGen,'',true,FTOBEcrGene);
     RempliLigne;
     CRemplirInfoLettrage(lTOBLigneEcr);
     lTOBLigneEcr.PutValue( 'E_DATEPAQUETMAX'   , FDtDatePaquetMax);
     lTOBLigneEcr.PutValue( 'E_DATEPAQUETMIN'   , FDtDatePaquetMin);
     lTOBLigneEcr.PutValue( 'E_NUMLIGNE'        , 2);
     lTOBLigneEcr.PutValue( 'E_GENERAL'         , FTOBLigne.GetValue('E_GENERAL'));
     CDupliquerInfoAux( FTOBLigne,lTOBLigneEcr );
     lTOBLigneEcr.PutValue( 'E_RIB'             , CGetRIB(lTOBLigneEcr));
     lTOBLigneEcr.PutValue( 'SELECTIONNER'      , 'X' );
     lTOBLigneEcr.PutValue( 'PASENREGISTRER'    , 'X' );
     {$IFDEF COMPTA}
     FRechCompte.Info.Load(lTOBLigneEcr.GetValue('E_GENERAL'),lTOBLigneEcr.GetValue('E_AUXILIAIRE'),EdtJournal.Text) ;
     if FRechCompte.Info.Compte.IsCollectif and ( FRechCompte.Info.GetString('T_NATUREAUXI') <> 'SAL' ) and ( length(trim( lTOBLigneEcr.GetValue('E_REGIMETVA') )) = 0 )   then
      CGetRegimeTVA ( lTOBLigneEcr , FRechCompte.Info ) ;
     //GereCutoff( lTOBLigneEcr ) ;
     {$ENDIF}
    end;

 if FBoMvtAuDebit then
  C:=FTOBLigne.GetValue('ECARTREGUL')
   else
    C:=FTOBLigne.GetValue('ECARTREGUL')*(-1);

 // gestion des colonnes debit / credit
 if ( edtTva.Enabled ) and ( Valeur(edtTva.text)<>0 ) and ( trim(FStCompteTVA) <> '' ) then
     begin

     lRdVal := Arrondi(C / ( 1 + ( Valeur(EdtTva.Text) / 100 ) ),FDevise.Decimale) ;
     lRdTva := Arrondi(C - lRdVal,FDevise.Decimale);

      if not FBoMvtAuDebit then
         begin
          CSetMontants(FTOBEcrGene.Detail[0],0,lRdVal,FDevise,true);
          CSetMontants(FTOBEcrGene.Detail[1],0,lRdTva,FDevise,true);
          CSetMontants(FTOBEcrGene.Detail[2],C,0,FDevise,true);
      end
          else
           begin
            CSetMontants(FTOBEcrGene.Detail[0],lRdVal,0,FDevise,true);
            CSetMontants(FTOBEcrGene.Detail[1],lRdTva,0,FDevise,true);
            CSetMontants(FTOBEcrGene.Detail[2],0,C,FDevise,true);
          end;
        end
         else
          begin
           if FBoMvtAuDebit then
            begin
             CSetMontants(FTOBEcrGene.Detail[0],C,0,FDevise,true);
             CSetMontants(FTOBEcrGene.Detail[1],0,C,FDevise,true);
            end
             else
              begin
               CSetMontants(FTOBEcrGene.Detail[0],0,C,FDevise,true);
               CSetMontants(FTOBEcrGene.Detail[1],C,0,FDevise,true);
              end; // if
          end;

  if FIFL.ChoixValid<>'AL6' then
   begin
    if FRLETTR.LettrageDevise then
     begin
      FTOBEcrGene.Detail[FTOBEcrGene.Detail.Count-1].PutValue('E_COUVERTUREDEV',C) ;
      ConvertCouverture(TOBM(FTOBEcrGene.Detail[FTOBEcrGene.Detail.Count-1]),tsmDevise) ;
     end
    else
        begin
         FTOBEcrGene.Detail[FTOBEcrGene.Detail.Count-1].PutValue('E_COUVERTURE',C) ;
         ConvertCouverture(TOBM(FTOBEcrGene.Detail[FTOBEcrGene.Detail.Count-1]),tsmPivot) ;
        end ;
   end;

  // generation de l'analytique apres l'affectation des montants
  for j := 0 to FTOBEcrGene.Detail.Count - 1 do
   begin
    lStCompte := FTOBEcrGene.Detail[j].GetValue('E_GENERAL') ;
    FRechCompte.Info.Load(FTOBEcrGene.Detail[j].GetValue('E_GENERAL'),FTOBEcrGene.Detail[j].GetValue('E_AUXILIAIRE'),EdtJournal.Text) ;
    lInIndex  := FRechCompte.Info.Compte.InIndex ;
    {$IFDEF COMPTA}
    CGetEch(FTOBEcrGene.Detail[j],FRechCompte.Info ) ;
    {$ENDIF}
    if ( lInIndex > -1 ) and ( FRechCompte.Info.Compte.IsVentilable(-1,lInIndex) )   then
     begin
      AlloueAxe( FTOBEcrGene.Detail[j] ) ; // SBO 25/01/2006
      CVentilerTOB(FTOBEcrGene.Detail[j],FRechCompte.Info) ;
      FTOBEcrGene.Detail[j].PutValue('E_ANA' , 'X') ;
     end; // if
   end; // for

 if CEquilibrePiece(FTOBEcrGene) then
  begin
   CAjouteChampsSuppLett(FTOBEcrGene.Detail[FTOBEcrGene.Detail.Count-1]);
   CAffectCompteContrePartie(FTOBEcrGene,FRechCompte.Info);
   CAffectRegimeTva(FTOBEcrGene) ;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. : - 04/05/2003 - FB 12234 - on passe la date paquet min et
Suite ........ : max a la fenetre de lettrage
Mots clefs ... :
*****************************************************************}
procedure TOF_ECRLET.BValiderClick(Sender: TObject);
begin
 if not IsValidEcr then
  begin
   TToolbarButton97(GetControl('BValider')).ModalResult := 0;
   exit;
  end// if
   else
    TToolbarButton97(GetControl('BValider')).ModalResult := 1;
 if not Generation then exit ;
 FStRetour := '1;'+EdtFolio.Text+';'+EdtDateComptable.Text+';'+EdtJournal.Text+';'+EdtLibelle.Text+';'+EdtGeneral.Text+';'+DateToStr(FDtDatePaquetMax)+';'+DateToStr(FDtDatePaquetMin);
 TFVierge(Ecran).retour := FStRetour;
 TheData                := FTOBEcrGene ;
 Ecran.Close;
end;

procedure TOF_ECRLET.BFermeClick(Sender: TObject); // Bouton fermer
begin
 TFVierge(Ecran).retour := '';
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/06/2002
Modifié le ... : 20/01/2003
Description .. : - FB 10111 - si le journal n'est pas valide on ne passe pas a
Suite ........ : la suite
Suite ........ : - FB 10763 - on ne recupere pas le compte de regul
Mots clefs ... :
*****************************************************************}
procedure TOF_ECRLET.EdtJournalExit(Sender: TObject);
var
 lStNomChamp     : string;
 lBoUnSeulCompte : boolean;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 FStGeneral1          := '';
 FStGeneral2          := '';
 FStGeneral3          := '';

 if not IsValidJournal then exit;

 FBoBordereau := ( FRechCompte.Info.Journal.GetValue('J_MODESAISIE') = 'LIB' ) or ( FRechCompte.Info.Journal.GetValue('J_MODESAISIE') = 'BOR' );

 EdtFolio.Visible     := FBoBordereau;
 lblFolio.Visible     := EdtFolio.Visible;
 lBoUnSeulCompte      := true;

 if FIFL.ChoixValid = 'AL1' then
  begin

   if not FBoMvtAuDebit then
    lStNomChamp := 'J_CPTEREGULCREDIT'
     else
      lStNomChamp := 'J_CPTEREGULDEBIT';

   if (FRechCompte.Info.Journal.GetValue(lStNomChamp + '1') <> '') then
    begin
     FStGeneral1     := FRechCompte.Info.Journal.GetValue(lStNomChamp + '1');
     lBoUnSeulCompte := true;
     //lBoPasDeCompte  := false;
    end;

   if (FRechCompte.Info.Journal.GetValue(lStNomChamp + '2') <> '') then
    begin
     FStGeneral2     := FRechCompte.Info.Journal.GetValue(lStNomChamp + '2');
     lBoUnSeulCompte := false;
    end; //if

   if (FRechCompte.Info.Journal.GetValue(lStNomChamp + '3') <> '') then
    FStGeneral3 := FRechCompte.Info.Journal.GetValue(lStNomChamp + '3');

   EdtGeneral.Text := FStLastGeneral ;
   if EdtGeneral.Text = '' then
    EdtGeneral.Text := FStGeneral1 ;


   EdtGeneral.Enabled := not lBoUnSeulCompte;

  end;// if

 {FQ19053 11/10/2007 YMO message d'alerte si jal différent de celui des param société}
 if (FIFL.ChoixValid = 'AL6')
 and GetParamSocSecur('SO_CPGESTCOMPENSATION', False)
 and (EdtJournal.Text<>GetParamSocSecur('SO_CPJALCOMPENSATION', False)) then
    PGIInfo('L''OD ne sera pas suivie au niveau de la TVA sur encaissement !','Attention');

 EdtNaturePiece.DataType  := FRechCompte.Info.Journal.GetTabletteNature;

 if FIFL.ChoixValid <> 'AL1' then // pour les ecritures simplifies
  EdtGeneral.Enabled := FRechCompte.Info.Journal.GetValue('J_CONTREPARTIE') = '' ; // pour les journal avec compte de contrêpartie on ne peut pas changer celui ci

// ColorControl(Tedit(EdtGeneral));

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 16/04/2002
Modifié le ... : 02/10/2003
Description .. : - on ne doit pas choisir les journaux de type a-nouveaux ou
Suite ........ : de cloture
Suite ........ :
Suite ........ : -FB 10112 - 10/06/2002 - correction de la valeur utilisee
Suite ........ : pour le MenuDisp du lookupList : 4 ouverture la fenetre des
Suite ........ : journaux
Suite ........ : -FB 12483 - 02/10/2003 - la fenetre s'ouvrait sur les 
Suite ........ : journaux fermes
Mots clefs ... : 
*****************************************************************}
procedure TOF_ECRLET.EdtJournalElipsisClick(Sender: TObject);
var
 lStWhere : string;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 lStWhere := 'J_FERME="-" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"ANA" AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA"';

 if (FIFL.ChoixValid = 'AL1') then
  begin
   if FRLETTR.LettrageDevise then
    lStWhere := lStWhere + 'AND J_NATUREJAL="REG" AND J_MULTIDEVISE="X" AND J_MODESAISIE="-"'
     else
      lStWhere := lStWhere + 'AND J_NATUREJAL="REG" AND J_MODESAISIE<>"LIB"';
  end;
  // CA - 05/02/2003 - FQ 11642
 if (FIFL.ChoixValid = 'AL3') then
  begin
   if FRLETTR.LettrageDevise then
    lStWhere := lStWhere + 'AND J_MULTIDEVISE="X" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO"'
   else lStWhere := lStWhere + 'AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO"';
  end;

 if LookupList(EdtJournal,TraduireMemoire('Journaux'),'JOURNAL','J_JOURNAL','J_LIBELLE',lStWhere,'J_JOURNAL',true, 4) then
  EdtJournalExit(nil);

end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 07/09/2004
Modifié le ... :   /  /    
Description .. : - 07/09/2004 - LG  - FB 14250 - on passe le code nature en 
Suite ........ : majuscule s'il est correct
Mots clefs ... : 
*****************************************************************}
procedure TOF_ECRLET.EdtNatureExit(Sender: TObject);
begin
 if IsValidNature then
  EdtNaturePiece.Text := UpperCase(EdtNaturePiece.Text) ;
end;

procedure TOF_ECRLET.EdtGeneralElipsisClick(Sender: TObject);
var
 lStWhere    : string;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 if FIFL.ChoixValid = 'AL1' then
  begin

   if (FStGeneral1 <> '') then
    lStWhere    := 'G_GENERAL="' + FStGeneral1 + '" ';

   if (FStGeneral2 <> '') then
    lStWhere    := lStWhere + 'or G_GENERAL="' + FStGeneral2 + '" ';

   if (FStGeneral3 <> '') then
    lStWhere    := lStWhere + 'or G_GENERAL="' + FStGeneral3 + '" ';

  end;// if

 if LookupList(EdtGeneral,TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',lStWhere,'G_GENERAL',true, 1) then
  EdtGeneralExit(nil);

end;


procedure TOF_ECRLET.EdtGeneralExit(Sender: TObject);
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 if (FIFL.ChoixValid = 'AL3') then
  begin
    if (EdtGeneral.Text = '')  then exit;
    IsValidGeneral;
  end
   else
    IsValidGeneral;

 
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : - 27/07/2005 - LG -  FB 15205 - on affiche dirstment le plus 
Suite ........ : grand element pour le numero de folio
Mots clefs ... : 
*****************************************************************}
procedure TOF_ECRLET.EdtFolioExit(Sender: TObject);
begin

  if ( csDestroying in Ecran.ComponentState ) then Exit ;

  IsValidNumeroPiece(true) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 06/07/2005
Modifié le ... :   /  /    
Description .. : -FB 15822 - 06/07/2005 - LG - Compte tenu de la 
Suite ........ : modification de l'ergonomie de la fiche et la création de 
Suite ........ : deux onglets "Détails" et "Ecriture", l'utilsation du bouton 
Suite ........ : "Voir Ecriture" (les lunettes) doit positionner sur l'onglet 
Suite ........ : "Détail"
Mots clefs ... : 
*****************************************************************}
procedure TOF_ECRLET.BtnVoirClick(Sender: TObject);
begin
 PgcDetail.ActivePageIndex                   := 1 ;
 TTabSheet(GetControl('PDETAIL')).visible    := true ;
 if TTabSheet(GetControl('PDETAIL')).CanFocus then
  TTabSheet(GetControl('PDETAIL')).SetFocus ;
 TTabSheet(GetControl('PDETAIL')).TabVisible := true ;
 G.Enabled                                   := true;
 Generation ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 01/04/2003
Modifié le ... : 12/05/2006
Description .. : - 01/04/2003 - FB 12234 - mise a jour de la date paquet 
Suite ........ : max
Suite ........ : - 18/10/2004 - LG - on reste sur ds la case date s'il n'est 
Suite ........ : pas correcte
Suite ........ : - 27/07/2005 - LG -  FB 15205 - on affiche dirstment le plus 
Suite ........ : grand element pour le numero de folio
Suite ........ : - LG - FB 17740  - on affecte le type d'exo pour le controel 
Suite ........ : des comptes vise 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ECRLET.EdtDateComptableExit(Sender: TObject);
begin

 if csDestroying in Ecran.ComponentState then Exit;

 if IsValidDateComptable then
  begin
   if StrToDate(EdtDateComptable.Text) >= FDtDatePaquetMax then
    FDtDatePaquetMax:=StrToDate(EdtDateComptable.Text) ;
   if StrToDate(EdtDateComptable.Text) < FDtDatePaquetMin then
    FDtDatePaquetMin:=StrToDate(EdtDateComptable.Text) ;
   {JP 14/02/07 : Mis dans l'ancêtre}
   RemplitComboFolio;

   FRechCompte.Info.TypeExo := CGetTypeExo(StrToDate(EdtDateComptable.Text)) ;
  end // if
   else
    if EdtDateComptable.CanFocus then EdtDateComptable.Setfocus ;

end;

procedure TOF_ECRLET.EdtDateComptableChange( Sender : TObject );
begin
 EdtDateComptableExit(nil);
end;

procedure TOF_ECRLET.EdtAuxiExit(Sender: TObject);
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 IsValidAuxi(true);

end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/06/2002
Modifié le ... : 17/10/2002
Description .. : - 17/10/2002 - le test sur le compte de contrepartie etait 
Suite ........ : faux
Mots clefs ... : 
*****************************************************************}
function TOF_ECRLET.IsValidJournal ( Parle : boolean = false ) : boolean;
begin

 FEcr.GetEcran(Ecran) ;
 FEcr.Initialize ;

 FRechCompte.Info.Journal.Load([EdtJournal.Text]) ;

 result := FEcr.IsValidJournal ; 

 if result and (FIFL.ChoixValid = 'AL1') and ( UpperCase(FRechCompte.Info.Journal.GetValue('J_NATUREJAL')) <> 'REG' ) then
  begin
   PGIInfo(cStPasCompteRegul,Ecran.Caption) ;
   result := false;
  end;

 if not result then
  begin
   EdtJournal.Text := '' ;
   if EdtJournal.CanFocus then EdtJournal.SetFocus;
   EdtJournalElipsisClick(nil);
   exit ;
  end;

 FStNatureJal := FRechCompte.Info.Journal.GetValue('J_NATUREJAL') ;

 if ( FRechCompte.Info.Journal.GetValue('J_CONTREPARTIE') <> '' ) and (EdtGeneral.Text <> FRechCompte.Info.Journal.GetValue('J_CONTREPARTIE')) then
  begin
   // contrôle que le compte de contrepartie est le compte de contrepartie du journal de banque
   if EdtGeneral.Text <> '' then PGIInfo('le compte doit être égale au compte de contrepartie du compte !','Attention'); // si le compte n'etaient pas definie on le remplace simplement
   EdtGeneral.Text           := FRechCompte.Info.Journal.GetValue('J_CONTREPARTIE');
   IsValidGeneral(Parle);
  end; // if


 EdtJournal.Text            := FEcr.Info.StJournal ;

 result                     := true;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 21/01/2003
Modifié le ... :   /  /    
Description .. : - FB 11800 - on pouvait mettre n'importe quoi dans la nature
Mots clefs ... : 
*****************************************************************}
function TOF_ECRLET.IsValidNature ( Parle : boolean = false ) : boolean;
begin

 FEcr.GetEcran(Ecran) ;
 FEcr.Initialize ;

 result         := FEcr.IsValidNat ;

 if not result then
  begin
  // declencher manuellement l'affichage de la tablette
   EdtNaturePiece.ElipsisClick(nil) ;
   if EdtNaturePiece.CanFocus then EdtNaturePiece.SetFocus;
   exit;
  end; //if

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 18/10/2002
Modifié le ... :   /  /    
Description .. : - 18/10/2002 - la gestion du taux de tva ne fct pour les 
Suite ........ : ecritures simplifiees depuis la consultation des comptes
Mots clefs ... : 
*****************************************************************}
function TOF_ECRLET.IsValidGeneral ( Parle : boolean = true ) : boolean;
var
 lInIndex          : integer;
 lRdTauxTva        : double ;
 lStNatureAuxi     : string ;
 lStRegimeTva      : string ;
 lStCompte         : string ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 04/01/2005
Modifié le ... : 15/01/2008
Description .. : - 04/01/2004 - LG  - FB 13678 - gestion de la tva pour les 
Suite ........ : tic ou tid, et pas pour les auxi de type sal
Suite ........ : - LG - 15/01/2008 - FB 13678 - idem
Mots clefs ... : 
*****************************************************************}
 procedure AffTablette;
 begin
  if EdtGeneral.CanFocus then EdtGeneral.SetFocus;
  if Parle then EdtGeneralElipsisClick(nil);
 end;

begin

 lStCompte := EdtGeneral.Text;
 lInIndex  := FRechCompte.Info.Compte.GetCompte(lStCompte);

 FEcr.GetEcran(Ecran) ;
 FEcr.Initialize ;

 result := FEcr.IsValidCompte ;

 // le compte doit etre defini est existe
 if not result then
  begin
   if FEcr.LastError.RC_Error in [ RC_AUXINEXISTANT , RC_AUXOBLIG ] then
    result := true 
     else
      begin
       EdtGeneral.Text := '';
       FStGeneral1     := '';
       AffTablette;
       exit;
      end ;
  end; // if

   //pour les ecritures de regul
 if (FIFL.ChoixValid = 'AL1') then
  begin
   if FStGeneral1 <> '' then
    result := ( EdtGeneral.Text = FStGeneral1 );
   if not result and (FStGeneral2 <> '') then
    result := ( EdtGeneral.Text = FStGeneral2 );
   if not result and (FStGeneral3 <> '') then
    result := ( EdtGeneral.Text = FStGeneral3 );
   if not result then
    begin
     if Parle then PGIError(cStPasCompteRegul,Ecran.Caption);
     AffTablette;
     exit;
    end; // if
  end; // if

 // afectation du taux de tva pour le compte
 lStNatureAuxi  := '' ;
 lStRegimeTVA   := '' ;
  if FRechCompte.Info.Aux.Load( [FTOBLigne.getValue('E_AUXILIAIRE')] ) <> -1 then
   begin
    lStNatureAuxi  := FRechCompte.Info.GetString('T_NATUREAUXI') ;
    lStRegimeTVA   := FRechCompte.Info.GetString('T_REGIMETVA') ;
    lblTva.Enabled := FRechCompte.Info.GetString('T_NATUREAUXI') <> 'SAL' ;
   end
   else
    if lInIndex > -1 then
    begin
     lStRegimeTVA := FRechCompte.Info.GetString('G_REGIMETVA') ;
     lblTva.Enabled := true ;
    end ;

 EdtGeneral.Text       := lStCompte;


 // la tva n'est pas accessible pour les ecretures de regul
 lblTva.Enabled        := lblTva.Enabled and FRechCompte.Info.Compte.IsTvaAutorise(edtGeneral.Text,lInIndex) and (FIFL.ChoixValid <> 'AL1') and (GetParamSocSecur('SO_REGIMEDEFAUT','')<>'') ;
 edtTva.Enabled        := lblTva.Enabled;

 if ( edtTva.Enabled ) and ( edtTva.Text = '' ) then
  begin


    FRechCompte.Info.Compte.RecupInfoTVA( edtGeneral.Text ,
                                          FTOBLigne.GetValue('ECARTREGUL'),
                                          EdtNaturePiece.Text,
                                          FStNatureJal ,
                                          lStNatureAuxi ,
                                          FStCompteTVA,
                                          lRdTauxTva,
                                          lStRegimeTva );
   edtTva.Text := StrS0(lRdTauxTva);
  end; // if

 FBoLettrable           := FRechCompte.Info.Compte.IsLettrable(lInIndex);
 FBoPointable           := FRechCompte.Info.Compte.IsPointable(lInIndex);
 FBoVentilable          := FRechCompte.Info.Compte.IsVentilable(-1,lInIndex);
 FBoCollectif           := FRechCompte.Info.Compte.IsCollectif(lInIndex);

 EdtAuxiliaire.Enabled  := FBoCollectif;
 lblAuxiliaire.Enabled  := FBoCollectif;
 FStNatGene             := FRechCompte.Info.GetString('G_NATUREGENE',lInIndex) ;

 if FBoCollectif then
  begin
   if EdtAuxiliaire.CanFocus then
    EdtAuxiliaire.SetFocus ;
  end
   else
    begin
     EdtAuxiliaire.Text := '' ;
     if EdtLibelle.CanFocus  then
      EdtLibelle.SetFocus;
    end ;
    
 result := true;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 11/06/2002
Modifié le ... : 21/11/2002
Description .. : -FB 10110 - correction du contrôle de la date
Suite ........ : - 21/11/2002 - suppression de warning
Mots clefs ... :
*****************************************************************}
function TOF_ECRLET.IsValidDateComptable ( Parle : boolean = false ) : boolean;
begin
  Result:= inherited IsValidDateComptable(Parle);
  if Result then begin
    FEcr.GetEcran(Ecran);
    FEcr.Initialize;
    Result := FEcr.IsValidDateComptable;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/09/2005
Modifié le ... :   /  /    
Description .. : - LG - 29/09/2005 -FB 15204 - on ne pouvait pas supprimer 
Suite ........ : un auxi
Mots clefs ... : 
*****************************************************************}
function TOF_ECRLET.IsValidAuxi (AutoriseCellExit : boolean = false ; Parle : boolean = false ) : boolean;
var
 lStNumCompte : string;
 lInIndex     : integer;
begin

 result        := true;

 if ( trim(EdtAuxiliaire.Text) = '' ) and not FBoCollectif then exit;

 result       := false;
 lStNumCompte := EdtAuxiliaire.Text;
 lInIndex     := FRechCompte.Info.Aux.GetCompte(lStNumCompte);

 if ( lInIndex < 0 ) then
  begin
   //PGIInfo('Cet auxiliaire n''existe pas !','Attention');
   EdtAuxiliaire.Text := '';
   if AutoriseCellExit then exit ;
   if EdtAuxiliaire.CanFocus then
    begin
     EdtAuxiliaire.SetFocus;
     CLookupListAux( EdtAuxiliaire , Copy(EdtGeneral.Text,1,1),EdtNaturePiece.Text, FStNatGene  ) ;
    end ; // if
   exit;
  end // if
   else
    begin
     EdtAuxiliaire.Text := FRechCompte.Info.GetString('T_AUXILIAIRE',lInIndex);
     if EdtGeneral.Text = '' then
      begin
       EdtGeneral.Text := FRechCompte.Info.GetString('T_COLLECTIF',lInIndex);
       result := IsValidGeneral(Parle);
       if not result then exit ;
      end;
      if not NATUREGENAUXOK(FStNatGene,FRechCompte.Info.GetString('T_NATUREAUXI',lInIndex)) then
       EdtAuxiliaire.Text := '' ;
    end; // if

 // test si l'auxiliare est lettrable
 if EdtAuxiliaire.Text <> '' then
  FBoLettrable := FRechCompte.Info.GetString('T_LETTRABLE',lInIndex) = 'X';

 result := true;

end;

function TOF_ECRLET.IsValidEtablissement ( Parle : boolean = false ) : boolean;
begin

 result := false;

 if EdtEtablissement.Text = '' then
  begin
   if Parle then PGIError(cStManqueEtablissement,Ecran.Caption);
   if EdtEtablissement.CanFocus then EdtEtablissement.SetFocus;
 //  TSpeedButton(EdtEtablissement.controls[0]).Click;
   exit;
  end;

 result := true;

end;

function TOF_ECRLET.IsValidNumeroPiece ( Parle : boolean = false ) : boolean;
var
 lQ     : TQuery ;
begin

 result := false;

 if FBoBordereau then
  begin
   if ( trim(EdtFolio.Text) = '' ) then
    begin
     if Parle then PGIError(cStNumeroPieceIncorrecte,Ecran.Caption);
     if EdtFolio.CanFocus then EdtFolio.SetFocus;
     exit;
    end
     else
     // if FRechCompte.Info.Journal.GetValue('J_MODESAISIE') = 'LIB' then
     //  begin
        lQ := OpenSql( ' select E_DEVISE,E_VALIDE from ecriture ' +
                       ' WHERE E_JOURNAL="' + EdtJournal.Text     + '"' +
                       ' AND E_EXERCICE="' + QuelExo(EdtDateComptable.Text)    + '" ' +
                       ' AND E_DATECOMPTABLE>="'+USDateTime(StrToDate( EdtDateComptable.Text )) + '" ' +
                       ' AND E_DATECOMPTABLE<="'+USDateTime(StrToDate( EdtDateComptable.Text )) + '" ' +
                       ' AND E_NUMEROPIECE=' + EdtFolio.Text +
                       ' AND E_QUALIFPIECE="N" AND E_NUMLIGNE=1 ' ,true ) ;
        if not lQ.Eof then
         begin
          if ( lQ.FindField('E_VALIDE').asString = 'X' ) then
           begin
            if Parle then PGIError(cStNumeroPieceValide,Ecran.Caption);
            if EdtFolio.CanFocus then EdtFolio.SetFocus;
            result := false ;
            exit ;
           end ;
          if ( FRechCompte.Info.Journal.GetValue('J_MODESAISIE') = 'LIB' ) and ( lQ.FindField('E_DEVISE').asString <> V_PGI.DevisePivot ) then
           begin
            if Parle then PGIError(cStNumeroPieceEnDev,Ecran.Caption);
            if EdtFolio.CanFocus then EdtFolio.SetFocus;
            result := false ;
            exit ;
           end ;
         end ; // if
       //end ;
  end ;
  
 result := true ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... :   /  /
Description .. : Contrôle la validite de la piece genere
Mots clefs ... :
*****************************************************************}
function TOF_ECRLET.IsValidEcr  : boolean;
begin
 result := false;
 if not IsValidNumeroPiece(true)     then exit;
 if not IsValidJournal               then exit;
 if not IsValidEtablissement         then exit;
 if not IsValidNature                then exit;
 if not IsValidGeneral               then exit;
 if not IsValidAuxi                  then exit;
 if not IsValidDateComptable         then exit;
 result := true;
end;

procedure TOF_ECRLET.OnError (sender : TObject; Error : TRecError ) ;
begin
 if trim(Error.RC_Message)<>'' then PGIInfo(Error.RC_Message, Ecran.Caption )
  else
   if not (Error.RC_Error in [RC_COMPTEINEXISTANT,RC_NATAUX, RC_AUXINEXISTANT, RC_AUXOBLIG , RC_PASCOLLECTIF ] ) then
    FMessCompta.Execute(Error.RC_Error);
end;

////////////////////////////////////////////////////////////////////////////////

constructor TEcrRegul.Create( vInfoEcr : TInfoEcriture ) ;
begin
 inherited ;
 FTOBEcrGene   := TOB.Create('',nil,-1);
end;

destructor TEcrRegul.Destroy;
begin
 FTOBEcrGene.Free;
end;

procedure TEcrRegul.RecupereLigneBor( vTOB : TOB );
var
 lTOBPiece : TOB;
 lStCrit : string;
 lInNumGroupeEcr : integer;
begin
 lStCrit:=vTOB.Detail[0].GetValue('E_JOURNAL')+vTOB.Detail[0].GetValue('E_MODESAISIE')+
          intToStr(vTOB.Detail[0].GetValue('E_PERIODE'))+intToStr(vTOB.Detail[0].GetValue('E_NUMEROPIECE')) ;
 // recherche s'il n'existe pas une piece avec ces critères
 lTOBPiece:=FTOBEcrGene.FindFirst(['CRIT'],[lStCrit],true) ;
 if lTOBPiece=nil then
  BEGIN
   lTOBPiece:=TOB.Create('',FTOBEcrGene,-1) ;
   lTOBPiece.AddChampSup('CRIT',true) ; lTOBPiece.PutValue('CRIT',lStCrit) ;
   lTOBPiece.AddChampSup('LETTRAGE',true) ; lTOBPiece.PutValue('LETTRAGE',vTOB.Detail[vTOB.Detail.count-1].GetValue('E_LETTRAGE')) ;
   lTOBPiece.AddChampSup('MODESAISIE',true) ; lTOBPiece.PutValue('MODESAISIE',vTOB.Detail[vTOB.Detail.count-1].GetValue('E_MODESAISIE')) ;
  END; // if
 if vTOB.Detail[0].GetValue('E_MODESAISIE')<>'BOR' then lInNumGroupeEcr:=1
 else if (lTOBPiece.Detail.Count=0) then lInNumGroupeEcr:=1
 else lInNumGroupeEcr:=lTOBPiece.Detail[lTOBPiece.Detail.Count-1].GetValue('E_NUMGROUPEECR')+1;
 while (vTOB.Detail.Count>0) do
  BEGIN
   vTOB.Detail[0].PutValue('E_NUMGROUPEECR',lInNumGroupeEcr) ;
   vTOB.Detail[0].ChangeParent(lTOBPiece, -1) ;
  END; // while
end;

procedure TEcrRegul.RecupereLignePiece( vTOB : TOB );
var
 lTOBPiece : TOB;
begin
 lTOBPiece:=TOB.Create('',FTOBEcrGene,-1) ;
 lTOBPiece.AddChampSup('CRIT',true) ; lTOBPiece.PutValue('CRIT','') ;
 lTOBPiece.AddChampSup('LETTRAGE',true) ; lTOBPiece.PutValue('LETTRAGE',vTOB.Detail[vTOB.Detail.count-1].GetValue('E_LETTRAGE')) ;
 lTOBPiece.AddChampSup('MODESAISIE',true) ; lTOBPiece.PutValue('MODESAISIE',vTOB.Detail[vTOB.Detail.count-1].GetValue('E_MODESAISIE')) ;
 while ( vTOB.Detail.Count > 0 ) do
  BEGIN
   vTOB.Detail[0].PutValue('E_NUMGROUPEECR', 1 ) ;
   vTOB.Detail[0].ChangeParent(lTOBPiece, -1) ;
  END; // while
end;


procedure TEcrRegul.RecupereLigne( vTOB : TOB );
begin
 if vTOB=nil then exit ; if vTOB.Detail.Count=0 then exit ;

 if vTOB.Detail[0].GetValue('E_MODESAISIE') = '-' then
  RecupereLignePiece( vTOB )
   else
    RecupereLigneBor( vTOB );

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/01/2004
Modifié le ... :   /  /
Description .. : - 22/01/2004 - le param vNumerote est passe a false ( le
Suite ........ : numero choisit ds la fentre de saisie bor n'etait pas pris en
Suite ........ : compte
Mots clefs ... :
*****************************************************************}
function TEcrRegul.Save : boolean;
var
 i : integer ;
begin
 result := false ;
 for i:= 0 to ( FTOBEcrGene.Detail.Count - 1 ) do
  if not CEnregistreSaisie( FTOBEcrGene.Detail[i],(FTOBEcrGene.Detail[i].Detail[0].GetValue('E_MODESAISIE') = '-'), True , true , Info ) then exit;
 result:=true ;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/04/2002
Modifié le ... :   /  /
Description .. : Supprime une piece suite à la suppression du lettrage
Suite ........ : associée
Mots clefs ... :
*****************************************************************}
procedure TEcrRegul.Delete( vStCodeLettrage : string );
var
 lInIndex : integer;
begin
 lInIndex := 0;

 while lInIndex <= ( FTOBEcrGene.Detail.Count-1 ) do
  begin
   if FTOBEcrGene.Detail[lInIndex].GetValue('LETTRAGE')= Copy(vStCodeLettrage,1,4) then
     FTOBEcrGene.Detail[lInIndex].Free;
   Inc(lInIndex);
  end; // while
end;


Initialization
  registerclasses ( [ TOF_ECRLET ] ) ;

end.

