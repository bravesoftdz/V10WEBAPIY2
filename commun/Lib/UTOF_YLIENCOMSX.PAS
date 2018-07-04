{***********UNITE*************************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YLIENCOMSX ()
                : inspiré de MBOINTERCOMPTA
Mots clefs ... : TOF;YLIENCOMSX
*****************************************************************}
Unit UTOF_YLIENCOMSX ;

Interface

Uses
     StdCtrls,
     Controls,
     Classes,
     Vierge,
     Variants,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ifndef GIGI}
     MAJTable,
{$ENDIF}
     fe_main,
{$else}
     maineagl,
{$ENDIF}
     UTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UtilGc,
     HTB97,
{$ifdef GIGI}
     DpTofDossierSel ,
{$endif}
     M3FP,
     UtilLienCOmsx
 ;

Type
  TOF_YLIENCOMSX = Class (TOF)
  Private
    NotClose, RapportFinTrt, GenerationFileOk, BlocageFiche : Boolean;
    FichierEcr, DateFile, BaseCompta, CodeCpta : string;
    BProg : TToolbarButton97;
    BImp : TToolbarButton97;

    procedure OnClickTotal;
    procedure OnClickReInit;
    procedure OnclickTacheSuiv( sTacheSuivante : String );
    procedure OnClickImprimer (Sender : TObject) ;
    procedure PrepareTOB( TOBCptRdu: TOB );
    procedure OnClickRazBaseCompta;
    //Lance la fiche de programmation des tâches
    procedure OnClickProg( Sender: TObject ) ;
    // Suppression de Job
    procedure AnnuleJobCree ;
    procedure OnExitCodeCpta;

    procedure ChargeTobEnteteLienCOmsx ;
    procedure LancementTransfertCompta ( TacheSuivante : TTypeTacheInterCpta = icNull );
    function ValidationSelection : String ;
    procedure MajFicheLienCOmsx ;
    procedure ForceEnregistrement ( FinTraitement : Boolean = False );
    procedure CreateMemo ;
    //Gestion d'affichage de la fiche
    procedure AffichageMemo ;
    procedure AffichageInfoTache ;
    procedure AffichageBoutonTache ;
    procedure AffichageImport ;
    procedure BloquerFiche ;
    //Controle de modification des champs de la fiche
    function IsChampsModifie : Boolean ;
    // Verifie si la tâche planifiée existe toujours
    procedure GestionNumJob ;
    // Vérifie que le paramétrage est complet
    function ParametrageOk : boolean ;
    // Gestion dates relatives
    procedure GestionDateRelative ( SAV : boolean = False ) ;
{$IFDEF GIGI}
  procedure DossierElipsisClick (Sender:TObject);
{$endif}

  Public
    Pages : TPageControl;
    ActionFiche : TActionFiche;
    ComptaExterne : Boolean;

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (St : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

procedure LanceYLienComsx(CodeCpta, Action : String) ;

Implementation

uses
  {$IFDEF EAGLCLIENT}
  UtileAGL,
  {$ELSE}
  EdtREtat,
  {$ENDIF}
  ParamSoc,
  uFicheJob,
  ed_tools,
  UtilPGI;
const  TexteMessage: array[1..6] of string = (
    {1} 'Veuillez confirmer la comptabilisation des écritures déjà extraites.'
    {2}, 'Attention: veuillez confirmer l''intégration des écritures dans la base courante.'
    {3}, 'Confirmez-vous la sélection ci-dessous ? #13#10'
    {4}, 'Tâche programmée en process serveur.'
    {5}, 'Vous devez renseigner la racine du nom de fichier.'
    {6}, 'Vous devez renseigner la base comptable de destination.'
     );

procedure TOF_YLIENCOMSX.OnNew ;
begin
  Inherited ;
  ActionFiche := taCreat ;
  BlocageFiche := False ;
  // Initialisation de la TOBEntete
  if Assigned(TobEnteteLienCOmsx) then
    FreeAndNil( TobEnteteLienCOmsx) ;
  TobEnteteLienCOmsx := TOB.Create( 'YLIENCOMSX',nil ,0 ) ;
  // Initialisation de la fiche
  SetControlText( 'YLO_CODECPTA', '' ) ;
  SetControlText( 'YLO_LIBELLECPTA', '' );
  SetControlText( 'YLO_ETABLISS', '' ) ;
  THEdit( GetControl( 'YLO_DATEDEBUT') ).Text := DateToStr( Date ) ;
  THEdit( GetControl( 'YLO_DATEFIN') ).Text := DateToStr( Date ) ;
  SetControlText( 'YLO_NOMFICHIER', '' ) ;
  SetControlText( 'YLO_BASECOMPTA', '' );
  SetControlProperty( 'YLO_DEJAEXTRAIT', 'state', cbUnChecked );
  SetControlText( 'YLO_NUMJOB', '0' );
  THValComboBox( GetControl( 'YLO_PLANIFIE' ) ).Value := 'ATT' ;
  THValComboBox( GetControl( 'YLO_EXPORT' ) ).Value := 'ATT' ;
  THValComboBox( GetControl( 'YLO_IMPORT' ) ).Value := 'ATT' ;
  SetControlText( 'YLO_DATEPLANIF', DateToStr( NowH ) ) ;
  SetControlText( 'YLO_USERPLANIF', V_PGI.User ) ;
  SetControlText( 'YLO_DATEEXPORT', DateToStr( NowH ) ) ;
  SetControlText( 'YLO_USEREXPORT', V_PGI.User ) ;
  SetControlText( 'YLO_DATEIMPORT', DateToStr( NowH ) ) ;
  SetControlText( 'YLO_USERIMPORT', V_PGI.User ) ;
  SetControlText( 'YLO_DATECREATION', DateToStr( NowH ) ) ;
  SetControlText( 'YLO_CREATEUR', V_PGI.User ) ;
  SetControlText( 'YLO_DATEMODIF', DateToStr( NowH ) ) ;
  SetControlText( 'YLO_UTILISATEUR', V_PGI.User ) ;
  SetControlText( 'YLO_BLOCNOTE', '' ) ;
  // Creation du mémo
  CreateMemo ;
  // Affichage
  AffichageInfoTache ;
  AffichageBoutonTache ;
  BloquerFiche ;
  AffichageImport ;
end ;

procedure TOF_YLIENCOMSX.OnDelete ;
begin
  Inherited ;
  if PGIAsk( TraduireMemoire( 'Confirmer-vous la suppression?' ), Ecran.Caption ) = mrNo then
  begin
    LastError := -1;
    exit;
  end else
  begin
    { Suppression du job }
    AnnuleJobCree ;
    if Assigned(TobEnteteLienCOmsx) then
      TobEnteteLienCOmsx.DeleteDB ;
    NotClose := False ;
    TToolBarButton97( GetControl('BFerme' )).click;
  end ;
end ;

procedure TOF_YLIENCOMSX.OnUpdate ;
begin
  Inherited ;
  NotClose := True;
  If GetControlText ('YLO_NOMFICHIER') ='' then
  begin
    SetFocusControl('YLO_NOMFICHIER');
    LastError := 5;
    LastErrorMsg := traduireMemoire(TexteMessage[LastError]);
    exit;
    end;
  // Mise à jour des champs d'information
  if ActionFiche = taCreat then
  begin
    TobEnteteLienCOmsx.SetString( 'YLO_CREATEUR', V_PGI.User ) ;
    TobEnteteLienCOmsx.SetDateTime( 'YLO_DATECREATION', NowH );
  end
  else begin
    TobEnteteLienCOmsx.SetString( 'YLO_UTILISATEUR', V_PGI.User ) ;
    TobEnteteLienCOmsx.SetDateTime( 'YLO_DATEMODIF', NowH );
  end ;
  // Récupération des modifications de l'écran
  TobEnteteLienCOmsx.GetEcran (Ecran );
{$ifdef GIGI}
  If GetControlText('BASECOMPTAPCL') <>'' then TobEntetelienCOmsx.SetString('YLO_BASECOMPTA',GetControlText('BASECOMPTAPCL'));
{$else}
  If GetControlText('BASECOMPTAPME') <>'' then TobEntetelienCOmsx.SetString('YLO_BASECOMPTA',GetControlText('BASECOMPTAPME'))
  else   begin
    SetFocusControl('BASECOMPTAPME');
    LastError := 6;
    LastErrorMsg := traduireMemoire(TexteMessage[LastError]);
    exit;
    end;
{$endif}
  GestionDateRelative( true ) ;  //mcd à voir
  if ParametrageOk then
  begin
    // Mise à jour de la table
    if ActionFiche = taCreat then
    begin
      TobEnteteLienCOmsx.InsertDB( nil ) ;
      ActionFiche := taModif ;
    end
    else
      TobEnteteLienCOmsx.UpdateDB ;
  end ;
  AffichageBoutonTache ;
end ;

procedure TOF_YLIENCOMSX.OnLoad ;
begin
  Inherited ;
  // Chargement de la TOB d'entete du transfert comptable
  if not ( ActionFiche = taCreat ) then
  begin
    ChargeTobEnteteLienCOmsx ;
    CreateMemo ;
    // Gestion d'affichage de la fiche
    GestionNumJob ;
    AffichageInfoTache ;
    AffichageBoutonTache ;
    BloquerFiche ; 
    AffichageImport ;
    GestionDateRelative ;
  end ;
end ;

procedure TOF_YLIENCOMSX.OnArgument (St : String ) ;
Var
  Critere, ChampMul, ValMul : String ;
  x : integer;
{$ifndef GIGI}
  SocCompta : THValComboBox;
  iNbItems, i : integer;
{$endif}
begin
  Inherited ;
  ActionFiche := TFVierge( Ecran ).TypeAction ;
  Pages := TPageControl( GetControl( 'PAGES', true ) );
  repeat
    Critere := uppercase( Trim( ReadTokenSt( St ) ) );
    if Critere <> '' then
    begin
      x := pos( '=', Critere );
      if x <> 0 then
      begin
        ChampMul := copy( Critere, 1, x - 1 );
        ValMul := copy( Critere, x + 1, length( Critere ) );
        if ChampMul = 'ACTION' then
        begin
          if ValMul = 'CREATION' then
            ActionFiche := taCreat
          else if ValMul = 'MODIFICATION' then
            ActionFiche := taModif ;
        end else
        if ChampMul = 'CODECPTA' then
          CodeCpta := ValMul ;
      end;
    end;
  until Critere = '';

  // Valeur pas défaut
  //obligation d'avoir 2 champs pour saisir la base dossier,car fait différement en PCL et en PME
  //en PCL fait sur table dossier
  // en PME fait sur cegidpgi.ini
{$ifdef GIGI}
  ComptaExterne:=true; // on va utiliser les bases cabinet
  SetControlVisible( 'BASECOMPTAPCL', ComptaExterne );
  SetControlVisible( 'BASECOMPTAPME', False );
{$else}
  ComptaExterne := ( GetParamSocSecur( 'SO_COMPTAEXTERNE',False ) );
  SetControlVisible( 'BASECOMPTAPME', ComptaExterne );
  SetControlVisible( 'BASECOMPTAPCL', False );
{$endif GIGI}
  SetControlVisible( 'YLO_ECRITURES', ComptaExterne or V_PGI.Sav );
  SetControlVisible( 'TYLO_BASECOMPTA', ComptaExterne );
  SetControlVisible( '_RAZBASECOMPTA', ComptaExterne );
  SetControlText('LREPERTOIRE',GetParamSocSecur('SO_MBOCHEMINCOMPTA','') ) ;

  // Gestion de la base compta
{$ifdef GIGI}  //on utilise en GI la base cabinet (comme dans établissement)
  ThEdit(GetControl('BASECOMPTAPCL')).OnElipsisClick := DossierElipsisClick;
{$else}
  if ComptaExterne then
  begin
    SocCompta := THValComboBox( GetControl( 'BASECOMPTAPME' ) );
    ChargeDossier( SocCompta.Items, True );
  end;
  // Mise à jour des values du combo
  iNbItems := THValComboBox(GetControl( 'BASECOMPTAPME' )).Items.Count;
  for i :=0 to iNbItems do
  begin
    THValComboBox(GetControl( 'BASECOMPTAPME' )).Values.Add(THValComboBox(GetControl( 'BASECOMPTAPME' )).Items[ i ] ) ;
  end ;
{$endif GIGI}

  //Force l'affichage du rapport
  RapportFinTrt := True;
  SetControlVisible( 'BImprimer', RapportFinTrt );
  SetControlText( 'YLO_NATURETRANS', 'JRL' );

  BImp := TToolBarButton97( GetControl( 'BImprimer') );
  if( Assigned(BImp) ) then
    BImp.OnClick := OnClickImprimer;
  BProg := TToolbarButton97( GetControl( 'BProg') );
  if( Assigned(BProg) ) then
    BProg.OnClick := OnClickProg;
  if ActionFiche = taCreat then
    New ;
    //zones de planification cachée si pas process serveur et/ou pas droits acces
  if (not(ExJaiLeDroitConcept(ccProgTask, False))) or (v_pgi.FTaskServer = '') then
  begin
    SetcontrolVisible ('YLO_NUMJOB',false);
    SetcontrolVisible ('TYLO_NUMJOB',false);
    SetcontrolVisible ('TYLO_PLANIFIE',false);
    SetcontrolVisible ('YLO_PLANIFIE',false);
    SetcontrolVisible ('TINFOTACHEPLANIF',false);
    SetcontrolVisible ('GBPLANIFICATION',false);
  end;
end ;

procedure TOF_YLIENCOMSX.OnClose ;
begin
  if Lasterror <>0 then exit; // si erreur dans Onupdate il ne faut pas que cela s'affiche 2 fois
  Inherited ;
  if IsChampsModifie or NotClose then
  begin
    LastError := -1;
    LastErrorMsg := '';
  end
  else
  begin
    LastError := 0;
    FreeAndNil( TobEnteteLienCOmsx);
    if Assigned(ZBlocNotes) then
      FreeAndNil(ZBlocNotes);
  end;
  NotClose := False;
end ;

{$IFDEF GIGI}
procedure TOF_YLIENCOMSX.DossierElipsisClick (Sender:TObject);
var
   sRetour  :string;
begin    //on affiche les bases déclarée cabinet
     // retourne NoDossier;CodePer;Nom1
 sRetour := DP_SelectUnDossier (GetControlText ('BASECOMPTAPCL'), true);
 if sRetour <> '' then
 begin
  SetControlText('BASECOMPTAPCL', ReadTokenSt (sRetour));
  if ActionFiche <> taCreat then ActionFiche := taModif ;
  SetControlText( 'YLO_BASECOMPTA' ,GetControlText('BASECOMPTAPCL'));//fait tout de suite pour ok ligne suivante
  AffichageImport ;
 end;
end;
{$endif GIGI}
{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Charge la TOBEntête du transfert comptable
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.ChargeTobEnteteLienCOmsx ;
Var
  SQLEntete : String ;
  QEntete : TQuery ;
begin
  if not Assigned(TobEnteteLienCOmsx) then
    TobEnteteLienCOmsx := TOB.Create( 'YLIENCOMSX',nil ,0 ) ;
      //normal de tout prendre.. on vaut saisir les info de la table
  SQLEntete := 'SELECT * FROM YLIENCOMSX WHERE YLO_CODECPTA = "'+ CodeCpta +'"' ;
  QEntete := OpenSQL(SQLEntete,False) ;
  TobEnteteLienCOmsx.SelectDB( '', QEntete ) ;
{$ifdef GIGI}   //pour OK affectation de la valeur dans la combo
  SetCOntrolText ('BASECOMPTAPCL',TobEnteteLienComsx.Getvalue('YLO_BASECOMPTA'));
{$else}
  SetCOntrolText ('BASECOMPTAPME',TobEnteteLienComsx.Getvalue('YLO_BASECOMPTA'));
{$endif}
  TobEnteteLienCOmsx.PutEcran( Ecran );
  if GetControlText('YLO_ETABLISS')='' then
    SetControlText('YLO_ETABLISS',TraduireMemoire('<<Tous>>') );
  SetControlText( 'YLO_DATEPLANIF', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATEPLANIF')) ) ;
  SetControlText( 'YLO_DATEEXPORT', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATEEXPORT')) ) ;
  SetControlText( 'YLO_DATEIMPORT', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATEIMPORT')) ) ;
  SetControlText( 'YLO_DATECREATION', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATECREATION') )) ;
  SetControlText( 'YLO_DATEMODIF', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATEMODIF') )) ;
  Ferme(QEntete) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... : 22/05/2006
Description .. : Mise à jour de la fiche transfert comptable
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.MajFicheLienCOmsx ;
begin
  if Assigned( TobEnteteLienCOmsx ) then
  begin
    SetControlText( 'YLO_DATEPLANIF', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATEPLANIF')) ) ;
    SetControlText( 'YLO_USERPLANIF', TobEnteteLienCOmsx.GetString('YLO_USERPLANIF') ) ;
    SetControlText( 'YLO_DATEEXPORT', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATEEXPORT')) ) ;
    SetControlText( 'YLO_USEREXPORT', TobEnteteLienCOmsx.GetString('YLO_USEREXPORT') ) ;
    SetControlText( 'YLO_DATEIMPORT', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATEIMPORT')) ) ;
    SetControlText( 'YLO_USERIMPORT', TobEnteteLienCOmsx.GetString('YLO_USERIMPORT') ) ;
    SetControlText( 'YLO_DATECREATION', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATECREATION')) ) ;
    SetControlText( 'YLO_DATEMODIF', DateToStr(TobEnteteLienCOmsx.GetDateTime('YLO_DATEMODIF')) ) ;
    SetControlProperty( 'YLO_PLANIFIE','value',TobEnteteLienCOmsx.GetString('YLO_PLANIFIE') ) ;
    SetControlProperty( 'YLO_EXPORT','value',TobEnteteLienCOmsx.GetString('YLO_EXPORT') ) ;
    SetControlProperty( 'YLO_IMPORT','value',TobEnteteLienCOmsx.GetString('YLO_IMPORT') ) ;
  end  ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Lance la fiche de programmation des tâches
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.OnClickProg( Sender: TObject );
Var
  aJobAction: TActionFiche;
  iNoJob, Res: Integer;
  TOBParam: TOB;
  Q: TQuery;
  SEQDateDeb, SEQDateFin, TestQuestion ,NomExe: String ;
  ModifTacheProcess : Boolean ;
  DateAvantOuv, DateApresOuv : TDateTime ;
begin
//fct non testée... à voir avec process serveur
  if ParametrageOk then
  begin
    BlocageFiche := True ;
    ModifTacheProcess := False ;
    DateAvantOuv := iDate1900 ;
    DateApresOuv := iDate1900 ;
    // Date relative
    SEQDateDeb := NullToVide(THEdit( GetControl( 'YLO_DATEDEBUT' ) ).SequenceKey) ;
    SEQDateFin := NullToVide(THEdit( GetControl( 'YLO_DATEFIN' ) ).SequenceKey) ;
    TestQuestion := ValidationSelection ;

    iNoJob := TobEnteteLienCOmsx.GetInteger( 'YLO_NUMJOB' ) ;
    TOBParam :=   TOB.Create( 'Paramètres', nil, -1 );  //pour process serveur

    TOBParam.AddChampSupValeur( ChCodeCpta, TobEnteteLienCOmsx.GetString( 'YLO_CODECPTA' ) );
    TOBParam.AddChampSupValeur( ChSeqDateDeb, SEQDateDeb );
    TOBParam.AddChampSupValeur( ChSeqDateFin, SEQDateFin );
    if (V_PGI.RunFromLanceur) then
      TOBParam.AddChampSupValeur( ChSociete, V_PGI.defaultSectionName )
    else TOBParam.AddChampSupValeur( ChSociete, V_PGI.CurrentAlias );
    TOBParam.AddChampSupValeur( ChCaptionEcr, TCheckBox(GetControl('YLO_DEJAEXTRAIT')).Caption );
    TOBParam.AddChampSupValeur( ChJournalEvent, TestQuestion );
    if (iNoJob > 0) and (ExisteSQL('SELECT 1 FROM STKJOBS WHERE SKJ_JOBID = ' + IntToStr(iNoJob))) then
    begin
      aJobAction := taModif ;
      Q := OpenSQL( 'SELECT SKJ_DATEMODIF FROM STKJOBS WHERE SKJ_JOBID = "' + IntToStr(iNoJob) +  '"', True );
      if( Assigned(Q) ) then
      try
        if( not(Q.EOF) ) then
          DateAvantOuv := Q.FindField('SKJ_DATEMODIF').AsDateTime ;
      finally
        Ferme( Q );
      end;
    end
    else
    begin
      iNoJob := -1;
      aJobAction := taCreat;
    end;
       //nom exe pour process serveur
    Nomexe:='';
    if ctxscot in V_PGI.PgiCOntexte then Nomexe:= 'CGIGESTIONINTERNE';
    if ctxtempo in V_PGI.PgiCOntexte then Nomexe:= 'CGIAFFAIRE';
    If NomExe='' then
      begin
      PgiInfo ('Vous devez renseigner le nom du process serveur');
      exit;
      end;

    Res := AglFicheJobExt( iNoJob, aJobAction,NomExe, 'TRANSFERT_COMPTA', TOBParam, '1PS',
                    Ecran.Caption, '', False, Ecran.Caption, true );
    FreeAndNil( TOBParam );
    // Sauvegarde du numéro de la tâche dans la table de paramétrage
    if( (res < 0) and ExisteSQL('SELECT 1 FROM STKJOBS WHERE SKJ_JOBID = ' + IntToStr(iNoJob)) ) then
      res := 0;
    if( (res <> 0) and (iNoJob <> res) ) then
    begin
      ModifTacheProcess := True ;
      SetControltext( 'YLO_NUMJOB', IntToStr( res ) );
      //TobEnteteLienCOmsx.SetInteger( 'YLO_NUMJOB', res ) ;
      TobEnteteLienCOmsx.SetString( 'YLO_USERPLANIF', V_PGI.User ) ;
      Q := OpenSQL( 'SELECT SKJ_DATECREATION FROM STKJOBS WHERE SKJ_JOBID = "' + IntToStr(res) +  '"', True );
      if( Assigned(Q) ) then
      try
        if( not(Q.EOF) ) then
          TobEnteteLienCOmsx.SetDateTime( 'YLO_DATEPLANIF', Q.FindField('SKJ_DATECREATION').AsDateTime );
      finally
        Ferme( Q );
      end;
    end else
    begin
      // Test si Tache Process modifiée
      Q := OpenSQL( 'SELECT SKJ_DATEMODIF FROM STKJOBS WHERE SKJ_JOBID = "' + IntToStr(iNoJob) +  '"', True );
      if( Assigned(Q) ) then
      try
        if( not(Q.EOF) ) then
          DateApresOuv := Q.FindField('SKJ_DATEMODIF').AsDateTime ;
      finally
        Ferme( Q );
      end;
      if ( DateAvantOuv <> iDate1900 ) and ( DateAvantOuv <> DateApresOuv ) then
        ModifTacheProcess := True ;
    end ;
    if ModifTacheProcess then
    begin
      SetControlEnabled('BEXECUTEIMPORT', False) ;
      SetControlEnabled('BEXECUTEEXPORT', False) ;
      SetControlEnabled('BTOTAL',False) ;
      BlocageFiche := True ;
      THValComboBox( GetControl( 'YLO_PLANIFIE' ) ).Value := 'ENC' ;
    end ;
    BloquerFiche ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 28/06/2006
Modifié le ... :   /  /    
Description .. : Suppression du job 
Mots clefs ... : 
*****************************************************************}
procedure TOF_YLIENCOMSX.AnnuleJobCree;
var
  NoJob: Variant;
  sNoJob: string;
begin
//fct non testée , à voir avec process serveur
  NoJob := StrToInt ( GetControlText( 'YLO_NUMJOB' ) ) ;
  if( (VarType(NoJob) = varInteger) and (NoJob > 0) ) then
  begin
    sNoJob := IntToStr( NoJob );
    ExecuteSQL( 'DELETE FROM STKDETAILS WHERE SKD_JOBID = ' + sNoJob );
    ExecuteSQL( 'DELETE FROM STKJOBS WHERE SKJ_JOBID = ' + sNoJob );
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Imprime le bloc note
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.OnClickImprimer( Sender : TObject );
var
  EdtTOB, TOBL: TOB;
begin
  EdtTOB := TOB.Create( '', nil, -1 );
  try
    TOBL := TOB.Create( '', EdtTOB, -1 );
    PrepareTOB( TOBL );
    LanceEtatTob( 'E', 'YLC', 'YLC', EdtTOB, True, False, False, nil, '', Ecran.Caption, False );
  finally
    EdtTOB.Free;
  end;
end;

procedure TOF_YLIENCOMSX.PrepareTOB( TOBCptRdu: TOB );
begin

if Assigned( TOBCptRdu ) then
  begin
    TOBCptRdu.AddChampSupValeur( 'CAPTION', 'Interface comptable'  );
    TOBCptRdu.AddChampSupValeur( 'COMPTERENDU',  ZBlocNotes.Text );
    TOBCptRdu.AddChampSupValeur( 'TITREENTETE', CodeCpta + ' / ' + GetControlText('YLO_LIBELLECPTA') );
  end;
end;

(*****************************************************************
FONCTION DE TRAITEMENT
*****************************************************************)
{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Lance le transfert comptable en entier
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.LancementTransfertCompta ( TacheSuivante : TTypeTacheInterCpta = icNull );
var
  DateDeb, DateFin, EtatCompta,
  TestQuestion,   ClauseEtab,Etat : string;
  TrtOk: Boolean;
  // Date relative pour le transfert comptable en process server.
  SEQDateDeb, SEQDateFin : String ;
begin
  inherited;
  if ParametrageOk then
  begin
    BlocageFiche := True ;
    LastError := 0;
    ZBlocNotes.Clear;
    // Récupération des paramètres pour remplir la table INTERCOMPTA
    DateDeb := TobEnteteLienCOmsx.GetString( 'YLO_DATEDEBUT' );
    DateFin := TobEnteteLienCOmsx.GetString( 'YLO_DATEFIN' );

    SEQDateDeb := THEdit( GetControl( 'YLO_DATEDEBUT' ) ).SequenceKey ;
    SEQDateFin := THEdit( GetControl( 'YLO_DATEFIN' ) ).SequenceKey ;

    BaseCompta := TobEnteteLienCOmsx.GetString( 'YLO_BASECOMPTA' );
{$ifdef GIGI} // il faut donner le bon nom de base à partir de la table dossier
    if not (V_PGI.RunFromLanceur) then
        BaseCompta := TobEnteteLienCOmsx.GetString( 'YLO_BASECOMPTA' )
    else if basecompta<>'' then Basecompta :='DB'+Basecompta + '@'+V_PGI.DEfaultSectionName;
{$endif}
    ClauseEtab := GetControlText('YLO_ETABLISS') ;//MultiComboTousInSQL (THMultiValComboBox( GetControl( 'YLO_ETABLISS' ) ) ) ;
    //ClauseEtab := StringReplace( ClauseEtab, ',', ';', [ rfReplaceAll ] );
    if( not TobEnteteLienCOmsx.GetBoolean( 'YLO_DEJAEXTRAIT' ) ) then
      EtatCompta := '"ATT"'
    else
    begin
      //Demande msg de conf
      if PGIAsk( TraduireMemoire( TexteMessage[ 1 ] ), Ecran.Caption ) = mrNo then
      begin
        LastError := -1;
        exit;
      end;
      EtatCompta := '"ATT","EXP"';
    end;
    // Confirmation si intégration dans le même base
{$ifdef GIGI}
    if BaseCompta = V_PGI.DefaultSectionName then
{$else}
    if BaseCompta = V_PGI.CurrentAlias then
{$endif}
    begin
      if PGIAsk( TraduireMemoire( TexteMessage[ 2 ] ), Ecran.Caption ) = mrNo then
      begin
        LastError := -1;
        Exit;
      end;
    end;
    TestQuestion := ValidationSelection ;
    if PGIAsk( TexteMessage[3] + TestQuestion , Ecran.Caption ) <> mrYes then
    begin
      LastError := -1;
      Exit;
    end ;
    //Calcul du nom du fichier .TRA ou .TRT
    if ( (TOBEnteteLienCOmsx.GetString('YLO_EXPORT')='ATT') ) then
      FichierEcr := NomDuFichierExport( GetControlText( 'YLO_NOMFICHIER' ),  DateFile )
    else   // déjà en cours de traitement
      FichierEcr := GetControlText( 'YLO_NOMFICHIER' );
    TOBEnteteLienCOmsx.setString( 'YLO_NOMFICHIER' , FichierEcr );
    Try
      TransfertCompta( LastError,
                         ClauseEtab,
                         GetControlText('YLO_DATEDEBUT'), GetControlText('YLO_DATEFIN'),
                        DateFile, EtatCompta, BaseCompta, FichierEcr,
                        TCheckBox(GetControl('YLO_DEJAEXTRAIT')).Caption,
                        TrtOk, GenerationFileOk,
                        GetControlText('YLO_DEJAEXTRAIT') <> '-', ComptaExterne, TacheSuivante, ZBlocNotes );
    finally
      SetControlVisible( 'BImprimer', RapportFinTrt );
      SetControlProperty( 'PRAPPORT', 'TabVisible', RapportFinTrt );
      // Mise à jour bloc note
      AffichageMemo ;
      ComsxSetLastError( LastError, Ecran.Caption );
      if LastError <>0 then Etat :='Err' else Etat:='Ok';
      MajJnalEvent ( 'COM', Etat,TraduireMemoire('COMSX : ')+GetCOntrolText('YLO_CODECPTA')+
        ' '+ GetCOntrolText('YLO_LIBELLECPTA'),ZBlocNotes.Text ) ;
      MajFicheLienCOmsx ;
      // Gestion d'affichage de la fiche
      AffichageInfoTache ;
      AffichageBoutonTache ;
      BloquerFiche ;
      NotClose := False ;
      ForceEnregistrement( True ) ;
    end ;
  end ;
end;

(***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Validation de la sélection et création d'un evenement dans
Suite ........ : le journal des evenements
Mots clefs ... :
*****************************************************************)
function TOF_YLIENCOMSX.ValidationSelection : String ;
Var
  TestQuestion,  Etablissement, BaseCompta,
  sEcritures, DejaExtraite, EnProcess: String ;
  SProcess : Integer ;
begin
  Result := '' ;
  Etablissement := GetControlText('YLO_ETABLISS') ;
  BaseCompta := THValComboBox( GetControl( 'YLO_BASECOMPTA' ) ).Text;
  sEcritures := GetControlText( 'YLO_DEJAEXTRAIT' );
  if( sEcritures = '-' ) then
    DejaExtraite := TraduireMemoire('Non')
  else
    DejaExtraite := TraduireMemoire('Oui')  ;
  SProcess := StrToInt( GetControlText( 'YLO_NUMJOB') ) ;
  if( SProcess = 0 ) then
    EnProcess := TraduireMemoire('Non')
  else
    EnProcess := TraduireMemoire('Oui')  ;

  TestQuestion := Format( TraduireMemoire('Paramétrage du transfert : %s'),
      [GetControlText('YLO_CODECPTA')]) + #13#10 ;
  TestQuestion := TestQuestion + '  ' + TraduireMemoire('Etablissements : ') + Etablissement + #13#10 ;
  TestQuestion := TestQuestion + '  ' + Format( TraduireMemoire( 'Période du %s au %s ' ),
      [ GetControlText( 'YLO_DATEDEBUT' ), GetControlText( 'YLO_DATEFIN' ) ] ) + #13#10 ;
  if BaseCompta <> '' then
    TestQuestion := TestQuestion + '  ' + TraduireMemoire('Base compta : ') + BaseCompta + #13#10 ;
    TestQuestion := TestQuestion + '  ' + TraduireMemoire('Comptabilisation des écritures déjà extraites : ') + DejaExtraite + #13#10 ;
    TestQuestion :=  TestQuestion + '  ' + TraduireMemoire('Planifier cette tâche pour exécution sur le serveur : ') + EnProcess ;
  Result := TestQuestion ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 21/04/2006
Modifié le ... :   /  /
Description .. : Force l'enregistrement de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOF_YLIENCOMSX.ForceEnregistrement ( FinTraitement : Boolean = False );
begin
  TobEnteteLienCOmsx.UpdateDB ;
  TobEnteteLienCOmsx.PutEcran( Ecran ) ;
  if FinTraitement then
    if GetControlText('YLO_ETABLISS')='' then
      SetControlText('YLO_ETABLISS',TraduireMemoire('<<Tous>>') );
  MajFicheLienComsx ;
end;

procedure TOF_YLIENCOMSX.CreateMemo ;
Var
  NomChampMemo : String ;
begin
  if Assigned ( ZBlocNotes ) then
    FreeAndNil(ZBlocNotes);
  ZBlocNotes := TStringList.Create;
  NomChampMemo := 'YLO_BLOCNOTE';
  ZBlocNotes.Text := TobEnteteLienCOmsx.GetString( NomChampMemo );
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 21/04/2006
Modifié le ... :   /  /    
Description .. : Affichage et enregistrement du bloc note
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.AffichageMemo( );
var
  NomChampMemo : string;
begin
  NomChampMemo := 'YLO_BLOCNOTE';
  TobEnteteLienCOmsx.setString( NomChampMemo, ZBlocNotes.Text  );
  SetControlText( NomChampMemo, ZBlocNotes.Text );
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 16/05/2006
Modifié le ... :   /  /    
Description .. : Affichage des informations concernant la tâche : Date / 
Suite ........ : utilisateur
Mots clefs ... : 
*****************************************************************}
procedure TOF_YLIENCOMSX.AffichageInfoTache ;
Var
  sInfo, sUtilisateur : String ;
begin
  // Info des tâches
  sInfo := TraduireMemoire('le ') + GetControlText( 'YLO_DATEPLANIF'  ) ;
  sUtilisateur := RechDom( 'TTUTILISATEUR', GetControlText( 'YLO_USERPLANIF'), False ) ;
  sInfo := sInfo + TraduireMemoire(' par ') + sUtilisateur ;
  SetControlText('TINFOTACHEPLANIF' , sInfo) ;
  sInfo := TraduireMemoire('le ') + GetControlText( 'YLO_DATEEXPORT'  ) ;
  sUtilisateur := RechDom( 'TTUTILISATEUR', GetControlText( 'YLO_USEREXPORT'), False ) ;
  sInfo := sInfo + TraduireMemoire(' par ') + sUtilisateur ;
  SetControlText('TINFOTACHEEXPORT' , sInfo) ;
  sInfo := TraduireMemoire('le ') + GetControlText( 'YLO_DATEIMPORT'  ) ;
  sUtilisateur := RechDom( 'TTUTILISATEUR', GetControlText( 'YLO_USERIMPORT'), False ) ;
  sInfo := sInfo + TraduireMemoire(' par ') + sUtilisateur ;
  SetControlText('TINFOTACHEIMPORT' , sInfo) ;
  // Info de création
  sInfo := TraduireMemoire('Créé le ') + GetControlText( 'YLO_DATECREATION' ) ;
  sUtilisateur := RechDom( 'TTUTILISATEUR',GetControlText( 'YLO_CREATEUR' ),False ) ;
  sInfo := sInfo + TraduireMemoire(' par ') + sUtilisateur ;
  SetControlText('TINFOCREATION' , sInfo) ;
  // Info de modification
  sInfo := TraduireMemoire('Modifié le ') + GetControlText( 'YLO_DATEMODIF' ) ;
  sUtilisateur := RechDom( 'TTUTILISATEUR',GetControlText( 'YLO_UTILISATEUR' ), False ) ;
  sInfo := sInfo + TraduireMemoire(' par ') + sUtilisateur ;
  SetControlText('TINFOMODIF' , sInfo) ;
end ;

procedure TOF_YLIENCOMSX.AffichageBoutonTache ;
Var
  TacheSuivante : TTypeTacheInterCpta ;
begin
  TacheSuivante := QuelleTacheSuivante  ;
  case TacheSuivante of
    icPlanifie :
      begin
        SetControlEnabled('BEXECUTEEXPORT', False) ;
        SetControlEnabled('BEXECUTEIMPORT', False) ;
        SetControlEnabled('BTOTAL',False) ;
        BlocageFiche := True ;
      end ;
    icExport :
      begin
        SetControlEnabled('BEXECUTEEXPORT', True) ;
        SetControlVisible('BEXECUTEIMPORT', False) ;
        SetControlEnabled('BTOTAL',True) ;
        BlocageFiche := True ;
      end ;
    icImport :
      begin
        SetControlEnabled('BEXECUTEEXPORT', False) ;
        SetControlVisible('BEXECUTEIMPORT', True) ;
        SetControlEnabled('BTOTAL',False) ;
        if ComptaExterne and ( GetControlText( 'YLO_BASECOMPTA')='' ) then
          SetControlEnabled( 'BREINIT' , True ) ;
        BlocageFiche := True ;
      end ;
    icNull :
      begin
        SetControlEnabled('BEXECUTEIMPORT', False) ;
        SetControlEnabled('BEXECUTEEXPORT', False) ;
        SetControlEnabled('BTOTAL',True) ;
        BlocageFiche := True ;
      end ;
    end ;
    // Execution total possible si toutes les taches sont en ATT
    if  ( THValComboBox( GetControl( 'YLO_PLANIFIE' ) ).Value = 'ATT' )
    and ( THValComboBox( GetControl( 'YLO_EXPORT' )).Value = 'ATT' ) then
    begin
      if not ( ActionFiche = taCreat ) then
        SetControlEnabled( 'BTOTAL' , True ) ;
      BlocageFiche := False ;
    end ;
   If( THValComboBox( GetControl( 'YLO_EXPORT' )).Value = 'OK' ) then
    begin
      SetControlEnabled( 'BTOTAL' , False ) ;
    end ;
end ;

procedure TOF_YLIENCOMSX.AffichageImport ;
Var
  IsVisible : Boolean ;
begin
  if ( ComptaExterne ) or ( THValComboBox( GetControl( 'YLO_BASECOMPTA' ) ).Text <> '' ) then
    IsVisible := True
  else
    isVisible := False ;
  SetControlVisible( 'TYLO_IMPORT', IsVisible ) ;
  SetControlVisible( 'YLO_IMPORT', IsVisible ) ;
  SetControlVisible( 'TINFOTACHEIMPORT',IsVisible ) ;
  SetControlVisible( 'BEXECUTEIMPORT',IsVisible ) ;
end ;

procedure TOF_YLIENCOMSX.BloquerFiche ;
begin
  SetControlEnabled( 'YLO_CODECPTA' , not BlocageFiche ) ;
  SetControlEnabled( 'YLO_LIBELLECPTA' , not BlocageFiche ) ;
  SetControlEnabled( 'YLO_ETABLISS' , not BlocageFiche ) ;
  SetControlEnabled( 'YLO_DATEDEBUT' , not BlocageFiche ) ;
  SetControlEnabled( 'YLO_DATEFIN' , not BlocageFiche ) ;
  SetControlEnabled( 'YLO_BASECOMPTA' , not BlocageFiche ) ;
  SetControlEnabled( '_RAZBASECOMPTA' , not BlocageFiche ) ;
  SetControlEnabled( 'YLO_DEJAEXTRAIT' , not BlocageFiche ) ;
  SetControlEnabled( 'YLO_NOMFICHIER', not BlocageFiche );
  SetControlEnabled( 'BProg' , not BlocageFiche ) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 01/06/2006
Modifié le ... : 01/06/2006
Description .. : Vérifie si les champs de la fiche ont été modifié
Mots clefs ... : 
*****************************************************************}
function TOF_YLIENCOMSX.IsChampsModifie : Boolean ;
Var
  Compare : Integer ;
  TOBTempEntete : TOB ;
  ListeChamps : String ;
begin
  Result := False ;
  if not Assigned(TobEnteteLienCOmsx) then
    exit ;
  TOBTempEntete := TOB.Create( 'YLIENCOMSX',nil ,0 ) ;
  TOBTempEntete.GetEcran( Ecran ) ;
{$ifdef GIGI}
  If GetControlText('BASECOMPTAPCL') <>'' then TobTempEntete.SetString('YLO_BASECOMPTA',GetControlText('BASECOMPTAPCL'));
{$else}
  If GetControlText('BASECOMPTAPME') <>'' then TobTempEntete.SetString('YLO_BASECOMPTA',GetControlText('BASECOMPTAPME'));
{$endif}
  ListeChamps := 'YLO_LIBELLECPTA;YLO_ETABLISS;YLO_DATEDEBUT;YLO_DATEFIN;'+
                'YLO_NOMFICHIER;YLO_BASECOMPTA;YLO_DEJAEXTRAIT;YLO_NUMJOB;'+
                'YLO_BLOCNOTE' ;
  Compare := CompareTOB( TobEnteteLienCOmsx, TOBTempEntete, ListeChamps );
  if Compare <> 0 then
  begin
    if PGIAsk( 'Confirmez-vous l''abandon des modifications?', Ecran.Caption ) = mrNo then
    begin
      NotClose := True ;
      Result := True ;
    end  else
    (* Annule le job que l'on vient de créer *)
    AnnuleJobCree;
  end ;
  FreeAndNil( TOBTempEntete );
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 01/06/2006
Modifié le ... :   /  /    
Description .. : Vérifie si la tâche en process serveur existe toujours.
Suite ........ : Sinon, réinitialise le numéro de la tâche. 
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.GestionNumJob ;
Var
  iNoJob : integer ;
begin
//fct non testée, à voir avec process serveur
  iNoJob := TobEnteteLienCOmsx.GetInteger( 'YLO_NUMJOB' ) ;
  if not ExisteSQL('SELECT 1 FROM STKJOBS WHERE SKJ_JOBID = ' + IntToStr(iNoJob)) then
  begin
    SetControltext( 'YLO_NUMJOB', '0' ) ;
    MAJStatutTache (TobEnteteLienCOmsx, 'ATT' , icPlanifie) ;
    MajFicheLienComsx ;
  end else
    BlocageFiche := True ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 05/06/2006
Modifié le ... :   /  /    
Description .. : Vérifie que le paramétrage de la liaison est ok
Mots clefs ... : 
*****************************************************************}
function TOF_YLIENCOMSX.ParametrageOk : boolean ;
begin
  Result := True ;
  if GetControlText('YLO_CODECPTA') = '' then
  begin
    Result := False ;
    PGIError( TraduireMemoire( 'Vous devez renseigner un code ' ), Ecran.Caption ) ;
    SetFocusControl('YLO_CODECPTA') ;
  end else
  if GetControlText('YLO_LIBELLECPTA') = '' then
  begin
    Result := False ;
    PGIError( TraduireMemoire( 'Vous devez renseigner un libellé ' ), Ecran.Caption ) ;
    SetFocusControl('YLO_LIBELLE') ;
  end else
  if GetControlText('YLO_ETABLISS') = '' then
  begin
    Result := False ;
    PGIError( TraduireMemoire( 'Vous devez renseigner un établissement ' ), Ecran.Caption );
    SetFocusControl('YLO_ETABLISS') ;
  end else
  if GetControlText('YLO_DATEDEBUT') = '' then
  begin
    Result := False ;
    PGIError( TraduireMemoire( 'Vous devez renseigner la date de début ' ), Ecran.Caption ) ;
    SetFocusControl('YLO_DATEDEBUT') ;
  end else
  if GetControlText('YLO_DATEFIN') = '' then
  begin
    Result := False ;
    PGIError( TraduireMemoire( 'Vous devez renseigner la date de fin ' ), Ecran.Caption ) ;
    SetFocusControl('YLO_DATEFIN') ;
  end else
  if ( ActionFiche = taCreat ) and ExisteSQL('SELECT YLO_CODECPTA FROM YLIENCOMSX WHERE YLO_CODECPTA="'+GetControlText('YLO_CODECPTA')+'"') then
  begin
    Result := False ;
    PGIError( TraduireMemoire( 'Ce code existe déjà, vous devez le modifier ' ), Ecran.Caption ) ;
    SetFocusControl('YLO_CODECPTA') ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 06/06/2006
Modifié le ... :   /  /    
Description .. : Gestion des dates relatives 
Mots clefs ... : 
*****************************************************************}
procedure TOF_YLIENCOMSX.GestionDateRelative ( SAV : boolean = False ) ;
//Var
 // SequenceDebut, SequenceFin : String ;
begin
(* fct non testée
utiliseée à priori pour le process serveur
à voir quand mis en place)
  if not SAV then
  begin
    // Affichage séquence des dates
    SequenceDebut := TobEnteteLienCOmsx.GetString( 'YLO_SEQDATEDEBUT' ) ;
    if SequenceDebut <> ''  then
    begin
      THEdit( GetControl( 'YLO_DATEDEBUT' ) ).SequenceKey := SequenceDebut ;
      THEdit( GetControl( 'YLO_DATEDEBUT' ) ).PlaySequenceKey ;
    end ;
    SequenceFin := TobEnteteLienCOmsx.GetString( 'YLO_SEQDATEFIN' ) ;
    if SequenceFin <> '' then
    begin
      THEdit( GetControl( 'YLO_DATEFIN' ) ).SequenceKey := SequenceFin ;
      THEdit( GetControl( 'YLO_DATEFIN' ) ).PlaySequenceKey ;
    end ;
  end else
  begin
    // Enregistrement séquence
    TobEnteteLienCOmsx.SetString( 'YLO_SEQDATEDEBUT' , NullToVide( THEdit( GetControl( 'YLO_DATEDEBUT' ) ).SequenceKey ) ) ;
    TobEnteteLienCOmsx.SetString( 'YLO_SEQDATEFIN' , NullToVide( THEdit( GetControl( 'YLO_DATEFIN' ) ).SequenceKey ) ) ;
  end ; *)
end ;

(*****************************************************************
ACTIONS FICHES
*****************************************************************)



{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Execute le transfert comptable totalement
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.OnClickTotal( );
begin
 ExecutionTotal := True ;
 if TobEnteteLienCOmsx.GetInteger( 'YLO_NUMJOB' ) = 0 then
    LancementTransfertCompta ( )
  else
    PGIInfo( TraduireMemoire( TexteMessage[ 4 ] ), Ecran.Caption );
end ;

procedure AGLOnClickTotal( parms : array of variant; nb : integer );
var
  F : TForm;
  TOTOF : TOF;
begin
  F := TForm( Longint( Parms[ 0 ] ) );
  if ( F is TFVierge ) then
    TOTOF := TFVierge( F ).LaTOF
  else
    exit ;
  if ( TOTOF is TOF_YLIENCOMSX ) then
    TOF_YLIENCOMSX( TOTOF ).OnClickTotal;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Réinitialise le transfert comptable
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.OnClickReInit( );
var  ii : integer;
begin
  BlocageFiche := False ;
  // Initialisation des tâches
  TobEnteteLienCOmsx.SetString( 'YLO_PLANIFIE' , 'ATT' ) ;
  TobEnteteLienCOmsx.SetString( 'YLO_EXPORT' , 'ATT' ) ;
  TobEnteteLienCOmsx.SetString( 'YLO_IMPORT' , 'ATT' ) ;
    //on efface le nom du fichier, tout en gardant la racine.
  ii:= Pos ( '-', TobEnteteLienCOmsx.GetString( 'YLO_NOMFICHIER'));
  if II > 0 then TobEnteteLienCOmsx.SetString( 'YLO_NOMFICHIER' ,Copy (TobEnteteLienCOmsx.GetString( 'YLO_NOMFICHIER'),1,ii-1 )) 
  else TobEnteteLienCOmsx.SetString( 'YLO_BLOCNOTE' , '' );
  (* Annule le job que l'on vient de créer *)
  AnnuleJobCree;
  TobEnteteLienCOmsx.SetString( 'YLO_NUMJOB' , '0' );
  MajFicheLienCOmsx ;
  AffichageBoutonTache ;
  BloquerFiche ;
  AffichageImport ;
  ForceEnregistrement ;
end ;

procedure AGLOnClickReInit( parms : array of variant; nb : integer );
var
  F : TForm;
  TOTOF : TOF;
begin
  F := TForm( Longint( Parms[ 0 ] ) );
  if ( F is TFVierge ) then
    TOTOF := TFVierge( F ).LaTOF
  else
    exit ;
  if ( TOTOF is TOF_YLIENCOMSX ) then
    TOF_YLIENCOMSX( TOTOF ).OnClickReInit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Execute la tache suivante transfert comptable totalement
Mots clefs ... :
*****************************************************************}
procedure TOF_YLIENCOMSX.OnclickTacheSuiv( sTacheSuivante : String );
Var
  TacheSuivante : TTypeTacheInterCpta ;
begin
  TFVierge (Ecran).Bvaliderclick(self) ; // sauvegarde obligatoire
  ExecutionTotal := False ;
  if sTacheSuivante = 'icExport' then
    TacheSuivante := icExport
  else if sTacheSuivante = 'icImport' then
    TacheSuivante := icImport
  else
    TacheSuivante := icNull ;
  if TobEnteteLienCOmsx.GetInteger( 'YLO_NUMJOB' ) = 0 then
    LancementTransfertCompta ( TacheSuivante )
  else
    PGIInfo( TraduireMemoire( TexteMessage[ 4 ] ), Ecran.Caption );   
end ;

procedure AGLOnclickTacheSuiv( parms : array of variant; nb : integer );
var
  F : TForm;
  TOTOF : TOF;
begin
  F := TForm( Longint( Parms[ 0 ] ) );
  if ( F is TFVierge ) then
    TOTOF := TFVierge( F ).LaTOF
  else
    exit;
  if ( TOTOF is TOF_YLIENCOMSX ) then
    TOF_YLIENCOMSX( TOTOF ).OnclickTacheSuiv ( Parms[ 1 ] );
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 16/05/2006
Modifié le ... :   /  /    
Description .. : Remet à blanc la base comptable
Mots clefs ... : 
*****************************************************************}
procedure TOF_YLIENCOMSX.OnClickRazBaseCompta( );
begin
  TobEnteteLienCOmsx.SetString( 'YLO_BASECOMPTA','' ) ;
  SetControlText( 'YLO_BASECOMPTA' , '' );
  SetControltext( 'BASECOMPTAPCL', '' );
  SetControlText( 'BASECOMPTAPME', '' );
  AffichageImport ;
end ;

procedure AGLOnClickRazBaseCompta( parms : array of variant; nb : integer );
var
  F : TForm;
  TOTOF : TOF;
begin
  F := TForm( Longint( Parms[ 0 ] ) );
  if ( F is TFVierge ) then
    TOTOF := TFVierge( F ).LaTOF
  else
    exit;
  if ( TOTOF is TOF_YLIENCOMSX ) then
    TOF_YLIENCOMSX( TOTOF ).OnClickRazBaseCompta;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 16/05/2006
Modifié le ... : 30/06/2006
Description .. : En sortie de champ, controle que le code n'existe pas
Mots clefs ... : 
*****************************************************************}
procedure TOF_YLIENCOMSX.OnExitCodeCpta;
begin
if ( ActionFiche = taCreat )  and ExisteSQL('SELECT YLO_CODECPTA FROM YLIENCOMSX WHERE YLO_CODECPTA="'+GetControlText('YLO_CODECPTA')+'"') then
  begin
    SetFocusControl('YLO_CODECPTA') ;
    PGIError( TraduireMemoire( 'Ce code existe déjà, vous devez le modifier ' ), Ecran.Caption )
  end ;
end ;

procedure AGLOnExitCodeCpta( parms : array of variant; nb : integer );
var
  F : TForm;
  TOTOF : TOF;
begin
  F := TForm( Longint( Parms[ 0 ] ) );
  if ( F is TFVierge ) then
    TOTOF := TFVierge( F ).LaTOF
  else
    exit;
  if ( TOTOF is TOF_YLIENCOMSX ) then
    TOF_YLIENCOMSX( TOTOF ).OnExitCodeCpta( );
end ;

{***********A.G.L.***********************************************
Auteur  ...... : MC Desseignet
Créé le ...... : 18/01/2007
Modifié le ... :   /  /    
Description .. : Ouvre la fiche du paramétrage automatisme COMSX
Mots clefs ... : 
*****************************************************************}
procedure LanceYLienCOmsx(CodeCpta, Action : String) ;
begin
  AGLLanceFiche( 'YY', 'YLIENCOMSX', '', '', 'CODECPTA=' + CodeCpta + ';ACTION='+Action );
end ;

Initialization
  RegisterAglProc( 'OnClickTotal', TRUE, 0, AGLOnClickTotal );
  RegisterAglProc( 'OnClickReInit', TRUE, 0, AGLOnClickReInit );
  RegisterAglProc( 'OnClickTacheSuiv', TRUE, 1, AGLOnclickTacheSuiv );
  RegisterAglProc( 'OnClickRazBaseCompta', TRUE, 0, AGLOnClickRazBaseCompta );
  RegisterAglProc( 'OnExitCodeCpta', TRUE, 1, AGLOnExitCodeCpta );
  registerclasses ( [ TOF_YLIENCOMSX ] ) ;
end.
