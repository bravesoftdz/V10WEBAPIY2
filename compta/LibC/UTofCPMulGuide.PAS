{***********UNITE*************************************************
Auteur  ...... : BOUSSERT Stéphane
Créé le ...... : 18/02/2003
Modifié le ... : 21/02/2003
Description .. : Source TOF de la FICHE : CPMULGUIDE ()
Suite ........ :
Suite ........ : Gestion du multi-sélection des guides.
Suite ........ : Utilisée à la fois pour le paramètrage et la selection en vue
Suite ........ : d'une saisie guidée (C.F. CPLanceFiche_MulGuide...)
Mots clefs ... : TOF;CPMULGUIDE
*****************************************************************}
Unit uTOFCPMulGuide ;

Interface

Uses 
     StdCtrls,
     Controls,
     Classes,
     Windows, //VK_F10
{$IFDEF EAGLCLIENT}
     uTob,
     eMul,
     Maineagl,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul,
     Fe_Main,
     HDB,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Ent1,      // pour le PositionneEtabUser, et QuelExoDT
     SaisUtil,  // Pout le RMVT
     AglInit,  //
     lookup,
     CPANCETREMUL,{JP 15/10/07 : FQ 16149 : gestion des restrictions utilisateur}
     HTB97
     ;

function CPLanceFiche_MulGuide( vStRange, vStLequel, vStArguments : String ) : String ;

Type
  TOF_CPMULGUIDE = Class (TOFANCETREMUL)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure BGuideClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure FormKeyDown( Sender : TObject; var Key: Word; Shift: TShiftState );
    procedure SuppressionGuide ;
    procedure EditionGuide ;
    procedure OuvreSaisie ;
    procedure OnKeyDownEcran (Sender : TObject; var Key : Word; Shift : TShiftState) ;
    procedure OnExitE_General             ( Sender : TObject );
    procedure OnElipsisClickE_General     ( Sender : TObject );

  private
    FEcran      : TFMul ;
    FcbJournal  : THMultiValComboBox ;
    FcbNatPiece : THMultiValComboBox ;
    FcbDevise   : THValComboBox ;
    FcbEtab     : THMultiValComboBox ;
    E_General   : THEdit;
    GLibelle    : THEdit;
    stMode      : String ;

    {$IFDEF EAGLCLIENT}
      FListe : THGrid ;
    {$ELSE}
      FListe : THDBGrid ;
    {$ENDIF}
  end ;

Implementation

uses
{$IFDEF MODENT1}
  CPProcMetier,
{$ENDIF MODENT1}
{$IFDEF CCSTD}

{$ELSE}
  Saisie, // Pour le lanceSaisie
{$ENDIF}
// Guide ;
  ULibWindows , 
  CPGUIDE_TOM;  // Pour le paramGuide



{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/02/2003
Modifié le ... :   /  /
Description .. : Ouvre le mul de sélection d'un guide de saisie.
Suite ........ : Existe dans le seul but de garder la compatibilité avec
Suite ........ : l'existant.
Mots clefs ... :
*****************************************************************}
function SelectGuide( vStJal, vStNatP, vStDev, vStEtab : string  ; vStMode : string = 'SELECT;') : String ;
var
  lStRange  : String ;
  lStLequel : String ;
  lStArgs   : String ;
begin

  lStRange  := '' ;
  lStArgs   := vStMode + vStJal + ';' + vStNatP + ';' + vStDev + ';' + vStEtab ;
  Result    := CPLanceFiche_MulGuide( lStRange, lStLequel, lStArgs ) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/02/2003
Modifié le ... : 19/02/2003
Description .. : Lancement de la fiche de sélection d'un guide de saisie
Suite ........ : Le dernier arguments de lancement doit comporter
Suite ........ : l'indication du mode de fonctionnement :
Suite ........ :  - 'CLASSIC' -> mode de sélection classique (la mouette bleu
Suite ........ :  ouvre la fiche du guide)
Suite ........ :  - 'SELECT'  -> mode de fonctionnement pour sélection (la
Suite ........ : mouette bleu ferme la fiche et on retourne le code du guide
Suite ........ : sélectionné)
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_MulGuide( vStRange, vStLequel, vStArguments : String ) : String ;
begin
 if vStArguments = '' then
  vStArguments := 'CLASSIC;' ;
 Result := AGLLanceFiche('CP', 'CPMULGUIDE', vStRange, vStLequel, vStArguments ) ;
end ;



// =============================================================================
// ==========               CPMULGUIDE_TOF                        ==============
// =============================================================================

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 06/04/2006
Modifié le ... : 24/09/2007
Description .. : - FB 17701 - 06/04/2006 - on affecte a Grayed de la
Suite ........ : checkbox 'guide de tresorerie'
Suite ........ : - 22/06/2006 - FB 18167 - on supprime l'onglet contenue
Suite ........ : - LG - FB 19719 - 10/07/2007 - acces a la fct en mode
Suite ........ : classic
Suite ........ : - LG - 23/07/2007 - FB 21137 - on desactive le filtre par
Suite ........ : defaut si on a des criteres de selection
Suite ........ : - LG - 24/09/2007 - FB 21495/FB 21495 - correction des 
Suite ........ : appel du mul ds guides ds les diff modules
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULGUIDE.OnArgument(S : String);
Var
    lStArg   : String;
    lStJal   : String;
    lStNatP  : String;
    lStDev   : String;
    lStEtab  : String;
    lStTreso : string ;
begin
    Inherited;

    // Récup Ecran
    FEcran                                       := TFMul(Ecran);
    FListe                                       := FEcran.FListe;
    FcbJournal                                   := THMultiValComboBox(GetControl('GU_JOURNAL'));
    FcbNatPiece                                  := THMultiValComboBox(GetControl('NATUREPIECE'));
    FcbDevise                                    := THValComboBox(GetControl('GU_DEVISE'));
    FcbEtab                                      := THMultiValComboBox(GetControl('ETABLISSEMENT'));
    Ecran.OnKeyDown                              := FormKeyDown;
    E_General                                    := THEdit(GetControl('G_GENERAL', True)) ;
    GLibelle                                     := THEdit(GetControl('EG_LIBELLE' , true));
    E_General.OnElipsisClick                     := OnElipsisClickE_General;
    E_General.OnDblClick                         := OnElipsisClickE_General;
    E_General.OnExit                             := OnExitE_General;

    TTabSheet(GetControl('PCONTENUE')).TabVisible := false ;

    // Ajout Evènements
    THGrid(GetControl('FListe')).OnDblClick      := FListeDblClick;
    TButton(GetControl('BGuide')).OnClick        := BGuideClick;
    TButton(GetControl('BInsert')).OnClick       := BInsertClick;
    TButton(GetControl('BDelete')).OnClick       := BDeleteClick;
//    Ecran.OnKeyDown                              := OnKeyDownEcran;

    // Par défaut, valcombobox sur tous :
//    FcbJournal.ItemIndex := 0;
//    FcbNatPiece.ItemIndex := 0;
    FcbDevise.ItemIndex := 0;

    TCheckBox(GetControl('GU_TRESORERIE', True)).AllowGrayed := true ;
    TCheckBox(GetControl('GU_TRESORERIE', True)).State       := cbGrayed ;

    // Décomposition des arguments :
    lStArg  := S ;
    if ( ReadTokenSt(lStArg) = 'SELECT') then
     stMode := 'SELECT'
      else
       stMode := 'CLASSIC' ;

    SetControlText('NATUREPIECE','<<Tous>>') ;
    {JP 15/10/07 : FQ 16149 : gestion des restrictions utilisateur}
    if GetControlText('ETABLISSEMENT') = '' then
      SetControlText('ETABLISSEMENT','<<Tous>>') ;

    // Journal ?
    if (Trim(lStArg) <> '') then
    begin
     TFMul(Ecran).FiltreDisabled := true ;
     lStJal := ReadTokenSt(lStArg);
     SetControlText('GU_JOURNAL',lStJal);
     stMode := 'SELECT' ;
    end;
    // Nature de pièce ?
    if (Trim(lStArg) <> '') then
    begin
        lStNatP := ReadTokenSt(lStArg);
        if (Trim(lStNatP) <> '') then SetControlText('NATUREPIECE',lStNatP) ;
    end;
    // Devise ?
    if (Trim(lStArg) <> '') then
    begin
        lStDev := ReadTokenSt(lStArg);
        if (Trim(lStDev) <> '') then SetControlText('GU_DEVISE',lStDev);
    end;
    // Etablissement ?
    if (Trim(lStArg) <> '') then
    begin
        lStEtab := ReadTokenSt(lStArg);
        if (Trim(lStEtab) <> '') then SetControlText('ETABLISSEMENT', lStEtab)
        else PositionneEtabUser(THValComboBox(GetControl('ETABLISSEMENT')));
    end
    else PositionneEtabUser(THValComboBox(GetControl('ETABLISSEMENT')));

    // guide de tresorerie
    SetControlEnabled('GU_TRESORERIE', stMode = 'CLASSIC' );
    SetControlChecked('GU_TRESORERIE' , false );

    if (Trim(lStArg) <> '') then
     begin
      lStTreso := ReadTokenSt(lStArg) ;
      SetControlEnabled('GU_TRESORERIE', false );
      SetControlChecked('GU_TRESORERIE' , lStTreso = 'X' );
      stMode := 'CLASSIC' ;
     end ;

    if ( stMode = 'SELECT' ) and FListe.CanFocus then
     FListe.setFocus ;

    FListe.SortEnabled := true ;

end;

procedure TOF_CPMULGUIDE.OnClose ;
begin
  Inherited ;
  if ( (not isInside(FEcran)) and (FEcran.ModalResult = mrOk) )
    then TFMul(Ecran).Retour := GetField('GU_GUIDE')
    else TFMul(Ecran).Retour := '' ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 18/07/2003
Modifié le ... : 17/05/2006
Description .. : le vk_f10 est repercuté
Suite ........ : - FB 18166 - le F5 qui ouvre le guide ne doit fct que sur la
Suite ........ : grille
Suite ........ : - FB 18168 - branchement du du ctrl+suppr
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULGUIDE.FormKeyDown( Sender : TObject; var Key: Word; Shift: TShiftState );
begin
 inherited;
 if ( csDestroying in Ecran.ComponentState ) then Exit ;

case Key of
  VK_F10,
  VK_F5 : if FListe.Focused then
           begin
            Key:=0 ;
            FListeDblClick(nil);
           end;
  VK_DELETE : if FListe.Focused and ( Shift=[ssCtrl] ) then begin Key:=0 ; BDeleteClick(nil) ; end ;
 end; // case

 if ( Ecran <> nil ) and ( Ecran is  TFMul ) then
    TFMul(Ecran).FormKeyDown(Sender,Key,Shift);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 10/07/2007
Modifié le ... :   /  /
Description .. : - LG - FB 19719 - 10/07/2007 - acces a la fct en mode
Suite ........ : classic
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULGUIDE.BGuideClick(Sender: TObject);
begin
  // Bouton accessible uniquement en mode 'Sélection'
  //if stMode = 'CLASSIC' then Exit ;
  EditionGuide ;
end ;

procedure TOF_CPMULGUIDE.FListeDblClick(Sender: TObject);
begin
  if stMode = 'CLASSIC' then
    // Mul utilisé pour consulter ou editer un guide
    EditionGuide
  else
    // Mul utilisé pour sélectionner un guide en vue d'une saisie
    OuvreSaisie ;
end;
                      
{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 10/07/2007
Modifié le ... :   /  /
Description .. : - LG - FB 19719 - 10/07/2007 - acces a la fct en mode 
Suite ........ : classic
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULGUIDE.BInsertClick(Sender: TObject);
begin
  // Bouton accessible uniquement en mode 'Classique'
  //if stMode <> 'CLASSIC' then Exit ;
  ParamGuide( '', 'NOR' , taCreat ) ;
  TFMUL(ECRAN).BChercheClick(nil) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 28/03/2006
Modifié le ... : 10/07/2007
Description .. : - FB 17714 - 28/03/2006 - un seul message qd on supprime 
Suite ........ : toute la grille
Suite ........ : - LG - FB 17714 - 20/09/2006 - un message de confirmation 
Suite ........ :  en entete et plsu a la ligne
Suite ........ : - LG - FB 17714 - 10/10/2006 - refonte du deplacement, ne 
Suite ........ : fct pas en eAGL
Suite ........ : - LG - FB 19719 - 10/07/2007 - acces a la fct en mode 
Suite ........ : classic
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULGUIDE.BDeleteClick(Sender: TObject);
var
 i       : integer ;
begin
  // Bouton accessible uniquement en mode 'Classique'
  //if stMode <> 'CLASSIC' then Exit ;

  if PGIAskCancel( 'Confirmez-vous la suppression des guides sélectionnés ?', Ecran.Caption ) <> mrYes then Exit ;

 {$IFNDEF EAGLCLIENT}
  TFMul(Ecran).Q.First;
  if FListe.AllSelected then
    while not TFMul(Ecran).Q.EOF do
     begin
      if transactions( SuppressionGuide ,1) <> oeOk then exit ;
      TFMul(Ecran).Q.Next;
     end
  else
  {$ENDIF}
   for i := 0 to FListe.nbSelected - 1 do
    begin
     FListe.GotoLeBookmark(i);
     {$IFDEF EAGLCLIENT}
     TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
     {$ENDIF}
     if transactions( SuppressionGuide ,1) <> oeOk then exit ;
  end;

 TFMUL(ECRAN).BChercheClick(nil) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 21/02/2003
Modifié le ... :   /  /    
Description .. : Supprime le guide courant en base
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULGUIDE.SuppressionGuide ;
var lStCode : string ;
    lStType : string ;
begin
  lStCode := GetField('GU_GUIDE') ;
  lStType := 'NOR' ; // GetField('GU_TYPE') ;
  if (lStType = '') or (lStCode = '') then Exit ;
  ExecuteSQL( 'DELETE FROM GUIDE WHERE GU_TYPE="'  + lStType + '" AND GU_GUIDE="' + lStCode + '"' ) ;
  ExecuteSQL( 'DELETE FROM ECRGUI WHERE EG_TYPE="' + lStType + '" AND EG_GUIDE="' + lStCode + '"' ) ;
  ExecuteSQL( 'DELETE FROM ANAGUI WHERE AG_TYPE="' + lStType + '" AND AG_GUIDE="' + lStCode + '"' ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 21/02/2003
Modifié le ... :   /  /    
Description .. : Ouvre la fiche des guides en mode édition
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULGUIDE.EditionGuide;
begin
 {$IFDEF EAGLCLIENT}
 TheMulQ := TFMul(Ecran).Q.TQ;
 {$ELSE}
  TheMulQ := TFMul(Ecran).Q ;
 {$ENDIF}
 if not (GetDataSet.Eof And GetDataSet.Bof) then
  ParamGuide( GetField('GU_GUIDE') , 'NOR' , taModif ) ;
 TFMUL(ECRAN).BChercheClick(nil) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 21/02/2003
Modifié le ... :   /  /    
Description .. : Ouvre la saisie pièce avec application du guide courant
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULGUIDE.OuvreSaisie;
{$IFDEF CCSTD}
{$ELSE}
var R : RMVT ;
{$ENDIF}
begin
{$IFDEF CCSTD}
  Pgiinfo('Fonction non implémentée.', Ecran.Caption);
{$ELSE}
  if not (GetDataSet.Eof And GetDataSet.Bof) then
    if isInside(FEcran) then
    begin
      FillChar(R,Sizeof(R),#0) ;
      R.LeGuide       := GetField('GU_GUIDE') ;
      R.DateC         := V_PGI.DateEntree ;
      R.Exo           := QuelExoDT( R.DateC ) ;
      R.Simul         := 'N' ;
      R.TypeGuide     := 'NOR' ;
      R.ANouveau      := False ;
      R.SaisieGuidee  := True ;
      LanceSaisie( Nil , taCreat , R ) ;
    end
    else
      FEcran.ModalResult := mrOk ;
{$ENDIF}
end;

procedure TOF_CPMULGUIDE.OnUpdate;
begin
 inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/09/2007
Modifié le ... :   /  /    
Description .. : - LG - 17/09/2007 - FB 21427 - gestion du <<tous>> pour els 
Suite ........ : etablissement et la nature
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULGUIDE.OnLoad;
var
 lStSQLP  : string ;
 lStSQLE  : string ;
 lStTemp  : string ;
begin

 inherited;

 lStSQLP := '' ;
 lStSQLE := '' ;

 TraductionTHMultiValComboBox( FcbNatPiece , lStSQLP , lStTemp ,'GU_NATUREPIECE' ) ;

 if lStSQLP <> '' then
  lStSQLP := '('  + lStSQLP + ' OR ( GU_NATUREPIECE="") ) ' ;

 TraductionTHMultiValComboBox( FcbEtab , lStSQLE , lStTemp ,'GU_ETABLISSEMENT' ) ;

 if lStSQLE <> '' then
  lStSQLE := '('  + lStSQLE + ' OR ( GU_ETABLISSEMENT="") ) ' ;

 if lStSQLE <> '' then
  lStSQLP := lStSQLP + ' AND ' + lStSQLE ;

 SetControlText('XX_WHERE', lStSQLP ) ;
 
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 06/04/2006
Modifié le ... : 11/04/2006
Description .. : - LG - FB 17655 - on ne renseigne pas le libelle qd on
Suite ........ : renseigne le code
Suite ........ : - LG - FB 17665 - on ne bourre plus les comptes
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULGUIDE.OnExitE_General(Sender: TObject);
var lQuery : TQuery;
begin

  if (csDestroying in E_General.ComponentState) then Exit;

  if E_General.Text = '' then
    Exit;

  try
    lQuery := OpenSQL('SELECT G_GENERAL,G_LIBELLE,G_CONFIDENTIEL FROM GENERAUX WHERE ' +
                      'G_GENERAL = "' + E_GENERAL.Text + '"', True);

    if not lQuery.Eof then
     begin
      // Test de la confidentialité des comptes
      if EstConfidentiel(lQuery.findField('G_CONFIDENTIEL').asString) then
        E_General.Text := '' ;
     end ;

  finally
    Ferme( lQuery );
  end;
end;

procedure TOF_CPMULGUIDE.OnElipsisClickE_General(Sender: TObject);
begin
 if ( Sender is THEdit ) and ( THEdit(Sender).Canfocus )  then
  LookUpList( THEdit(Sender),
              'Comptes généraux',
              'GENERAUX',
              'G_GENERAL' ,
              'G_LIBELLE',
              CGenereSQLConfidentiel('G'),
              'G_GENERAL' ,
              True,
              1 );
end;


procedure TOF_CPMULGUIDE.OnKeyDownEcran( Sender : TObject ; var Key : Word; Shift : TShiftState);
begin

 inherited ;

 case Key of
   VK_F9  : TToolBarButton97(GetControl('BCherche', True)).Click

  end ; // case

end ;


Initialization
  registerclasses ( [ TOF_CPMULGUIDE ] ) ;
end.




