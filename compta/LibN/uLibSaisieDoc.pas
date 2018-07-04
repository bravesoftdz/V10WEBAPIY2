{***********UNITE*************************************************
Auteur  ...... : SBO
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. :
Suite ........ : Gestionnaire du visualisateur de documents dans la saisie
Suite ........ : paramétrable.
Suite ........ :
Suite ........ : Affichage en fonction des paramètres sociétés de la page
Suite ........ : "SAISIE"
Suite ........ :
Suite ........ : Possibilité d'intégration des documents en visu dans la GED
Mots clefs ... :
*****************************************************************}
unit uLibSaisieDoc;

interface

uses
  Hctrls,            // THGrid
  SysUtils,          // TSearchRec
  Forms,             // TForm
  Menus,             // TPopupMenu
  Controls,          // pour le TwinControl
  SHDocVw_TLB,       // TWebBrowser_V1
  HTB97,             // TToolBarButton97 , TDock97
  UGedFileViewer,    // TGedInternalViewer
  hpanel,            // THpanel
  uLibPieceCompta,    // TPieceCompta
  HEnt1,             // TActionFiche, sourissablier
  UTob
  ;

Type

TViewerPos = ( vpBas, vpGauche, vpDroite, vpHaut ) ;


TSaisieDoc = Class
  private
    // Variables de gestion des fichiers
    FStFichier        : string;
    FStChemin         : string;
    FStExt            : string;
    FSearchRec        : TSearchRec;
    FBoModeWeb        : Boolean ;
    FEtatBoutons      : TGedButtonStates ;
    FBoFichierOk      : Boolean ;
    gFindFirst        : Boolean ;

    // Composants présents sur la fiche
    FEcran           : TForm ;

    FPanelRight      : THPanel ;
    FPanelBottom     : THPanel ;
    FPanelLeft       : THPanel ;
    FPanelTop        : THPanel ;

    FSplitterRight   : THSplitter ;
    FSplitterBottom  : THSplitter ;
    FSplitterLeft    : THSplitter ;
    FSplitterTop     : THSplitter ;

    FPopUpViewer     : TPopupMenu ;
    FBoutonMenu      : TToolBarButton97 ;
    FBoutonMenuMove  : TToolBarButton97 ;
    FBExt            : TToolBarButton97 ;
    FBInt            : TToolBarButton97 ;

    // Composants créés pour la GED
    FPanelActif      : THPanel ;
    FSplitterActif   : THSplitter ;
    FWebBrowser      : TWebBrowser_V1 ;
    FInternalViewer  : TGedInternalViewer;

    // Rechargement des info paramSoc
    FBoAvecViewer    : Boolean ;
    FBoAvecImportDoc : Boolean ;
    FPosition        : TViewerPos ;   // 1BA = vpBas , 2DR = vpDroite, 3GA = vpGauche, 4HO = vpHaut

    // Gestion de la pièce
    FPieceCompta     : TPieceCompta ; // référence à la pièce
    FInGroupeActif   : integer ;      // Numéro de groupe auquel est lié le document en cours

    // init
    procedure initialisation ;
    procedure SetPanelVisible ( vBoVisible : Boolean = true );
    procedure ReloadParam ( vTobParam : Tob ) ;

    procedure SetFichier  ( vStFic : string ) ;
    function  SetDossier  (Rep : String) : String ;

  public

    constructor Create           ( vEcran : TForm ; vPiece : TPieceCompta ) ;
    destructor  Destroy          ; override ;

    // init
    procedure InitViewer( vPanelB, vPanelR, vPanelL, vPanelT : THPanel ;
                          vSplitB, vSplitR, vSplitL, vSplitT : THSplitter ) ;
    procedure DesInitViewer ;

    // gestion de fichier
    function  FindFirstFile (var vSearchRec : TSearchRec ) : String ;        // se positionne sur le 1er fichier à traiter
    function  FindNextFile  (var vSearchRec : TSearchRec ) : String ;        // se positionne sur le fichier suivant à traiter
    function  FindPrevFile          : integer ;
    procedure SetSearchRec ;
    function  GetCheminStockage     : string ;      // retourne l'emplacement utilisé poru le stockage des fichiers
    function  StockeFichier         ( vStFichier : string ) : boolean ;   // Déplace le fichier dans le rép de stockage + stocke réf dans la Tob
    procedure UpdateChemin ;
{$IFDEF SCANGED}
    function  AjouterFichierDansGed ( vStFichier : string ; vTobEcr : Tob ) : string ;
    procedure GetFichierGED ;
{$ENDIF SCANGED}

    // gestion affichage
    procedure ActiverViewer ( vTobParam : Tob ; vBoVisible : Boolean = true ) ;    // Gestion de l'affichage du viewer dans la fenêtre de saisie en fonction du paramétrage
    procedure DesactiveViewer ;                      // vide le visualiseur
    procedure Refresh ;                              // Affiche le document sélectionné dans le visaualiseur
    procedure MontreViewer      ( vBoVal : Boolean ) ;

    procedure SetVisibleBDoc(vBoInt : Boolean) ;

    procedure WebRefresh ;
    procedure InternalRefresh ;
    procedure GereEtatBoutons (Sender: TObject; ButtonStates: TGedButtonStates) ;

    procedure ChargeDocument ( vBoSuivant : boolean = true ; vBoRefresh : boolean = true ; vBoStocke : boolean = true ) ;

    // autres
    function  GetInfoTob    ( vTobEcr : Tob ) : string ;
    function  IsActif       : boolean ;
    function  IsVisible     : boolean ;
    function  EstFichierOk  ( vStFic : string ) : boolean ;
    function  IsFichier     ( vStFic : string ) : boolean ;
    function  FitToPageOk   : boolean ;
    function  FitToLargeOk  : boolean ;
    function  FitTo100Ok    : boolean ;
    function  PanOk         : boolean ;
    function  ZoomOk        : boolean ;
    function  PrintOk       : boolean ;
    function  DeplacementOk : boolean ;
    function  RotateImageOk : boolean ;
    function  ZoomImageOk   : boolean ;

    // Actions images
    procedure EnBasADroite ;
    procedure EnHautAGauche ;
    procedure EnBas         ( vBoLarge : Boolean ) ;
    procedure EnHaut        ( vBoLarge : Boolean ) ;
    procedure ADroite       ( vBoLarge : Boolean ) ;
    procedure AGauche       ( vBoLarge : Boolean ) ;
    procedure ZoomImage     ( vBoPlus  : boolean ) ;
    procedure MoveDoc       ( vRepDst : String ; vBoSuivant : Boolean = True) ;
    procedure RotateImage   ;
    procedure Rotate90      ;

    // ---> Popup Action Viewer
    procedure SetMenuPop                  ( vPopup : TPopupMenu ; vBouton : TToolBarButton97 ; vBoutonMove : TToolBarButton97 ) ;
    procedure SetBDoc                     ( vBExt  : TToolBarButton97 ; vBInt : TToolBarButton97 ) ;
    procedure OnPopUpViewer               ( Sender : TObject ) ;
    procedure InitPopupViewer             ( vActivation : Boolean = True  ) ;
    procedure ViewerPrintClick            ( Sender : TObject ) ;
    procedure ViewerPanClick              ( Sender : TObject ) ;
    procedure ViewerZoomAvClick           ( Sender : TObject ) ;
    procedure ViewerZoomArClick           ( Sender : TObject ) ;
    procedure ViewerFitPageClick          ( Sender : TObject ) ;
    procedure ViewerFitLargeClick         ( Sender : TObject ) ;
    procedure ViewerFit100Click           ( Sender : TObject ) ;
    procedure ViewerNextDocClick          ( Sender : TObject ) ;
    procedure ViewerPrevDocClick          ( Sender : TObject ) ;
    procedure ViewerRotateClick           ( Sender : TObject ) ;
    procedure ViewerMoveDocACHClick       ( Sender : TObject ) ;
    procedure ViewerMoveDocBQEClick       ( Sender : TObject ) ;
    procedure ViewerMoveDocCAIClick       ( Sender : TObject ) ;
    procedure ViewerMoveDocODClick        ( Sender : TObject ) ;
    procedure ViewerMoveDocVTEClick       ( Sender : TObject ) ;
    procedure ViewerMoveDocCorbeilleClick ( Sender : TObject ) ;

    // ================
    // === PROPERTY ===
    // ================
    property  Fichier        : string              read FStFichier          write SetFichier ;
    property  Chemin         : string              read FStChemin           write FStChemin ;
    property  FichierOk      : boolean             read FBoFichierOk        ;
    property  Position       : TViewerPos          read FPosition           ;
    property  PanelActif     : THPanel             read FPanelActif         ;
    property  SplitterActif  : THSPLitter          read FSplitterActif      ;
    property  WebBrowser     : TWebBrowser_V1      read FWebBrowser         ;
    property  InternalViewer : TGedInternalViewer  read FInternalViewer     ;
    property  Piece          : TPieceCompta        read FPieceCompta        write FPieceCompta ;
    property  GroupeActif    : integer             read FInGroupeActif      write FInGroupeActif ;
    property  SearchRec      : TSearchRec          read FSearchRec          write FSearchRec ;
  end ;

implementation

uses
  {$IFDEF EAGLCLIENT}
  {$ELSE}
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF EAGLCLIENT}
  {$IFDEF VER150}
   Variants,
  {$ENDIF}
  hmsgbox,           // pour MessageAlerte, pgiInfo
  CBPPath,           // TCBPPath
  uGedFiles,         // V_GEDFILES
{$IFDEF SCANGED}
  utilGed,           // InsertDocumentGed
{$ENDIF SCANGED}
  windows,           // moveFiles
  messages,          // WM_KEYDOWN
  uSmoothImage,      // TSmoothImage
  stdCtrls,          // TScrollBar
  graphics,          // TBitmap
  Classes,           // Rect
  ParamSoc,
  hdebug,            // Debug
  galoutil,          // GetParamSocDP
  wCommuns           // wGetMonth + wYear
  ;

{ TSaisieDoc }

procedure TSaisieDoc.ActiverViewer ( vTobParam : Tob ; vBoVisible : boolean );
begin

  // rechargement des paramètres
  ReloadParam( vTobParam ) ;

  // viewer désactiver
  if not FBoAvecViewer then
  begin
    FPanelActif    := nil ;
    FSplitterActif := nil ;
    if not FBoModeWeb and assigned(FBoutonMenu) then
    begin
      FBoutonMenu.Visible := False ;
      if assigned(FBoutonMenuMove) then
         FBoutonMenuMove.Visible := False ;
    end;
  end
  // viewer en fonction de la position //  TViewerPos = ( vpBas, vpGauche, vpDroite, vpHaut ) ;
  else
    begin
    Case FPosition Of
      vpBas :
        begin
        FPanelActif    := FPanelBottom ;
        FSplitterActif := FSplitterBottom ;
        end ;
      vpDroite :
        begin
        FPanelActif    := FPanelRight ;
        FSplitterActif := FSplitterRight ;
        end ;
      vpGauche :
        begin
        FPanelActif    := FPanelLeft ;
        FSplitterActif := FSplitterLeft ;
        end ;
      vpHaut :
        begin
        FPanelActif    := FPanelTop ;
        FSplitterActif := FSplitterTop ;
        end ;
      end; // case
    end; // else

  SetPanelVisible (vBoVisible)   ;

  // === MODE WEB BROWSER ===
  if FBoModeWeb then
    begin
    if Assigned(FPanelActif) then
      begin
      // Si le panel parent à changer, on repositionne le browser
      if not Assigned(FWebBrowser.parent) or
         ( TControl(FWebBrowser.parent).Name <> FPanelActif.Name ) then
        begin
        FWebBrowser.Visible := False ;
        TWinControl(FWebBrowser).parent  := TWinControl(FPanelActif) ;
        FWebBrowser.Top     := FPanelActif.Top ;
        FWebBrowser.Left    := FPanelActif.Left ;
        FWebBrowser.Height  := FPanelActif.Height ;
        FWebBrowser.Width   := FPanelActif.Width ;
        FWebBrowser.Align   := alClient ;
        end ;
      // Affichage du browser
      FWebBrowser.Visible   := true ;
      FWebBrowser.Invalidate ;
      end
    else
      FWebBrowser.Visible := false ;
    end
  else
  // === MODE INTERNAL VIEWER ===
    begin
    // si changement de panel actif, alors on libère l'ancien viewer pour recréation dans nouveau panel
    if Assigned(FPanelActif) then
      begin
      // Création du viewer interne
      if Assigned( FInternalViewer ) and (FInternalViewer.owner <> FPanelActif) then
        FreeAndNil( FInternalViewer ) ;
      FBoutonMenu.Visible := vBoVisible ;
      if assigned(FBoutonMenuMove) then
         FBoutonMenuMove.Visible := vBoVisible ; 
      end ;
    end ;

end;

function TSaisieDoc.FindFirstFile (var vSearchRec : TSearchRec) : String ;
var i :integer ;
begin
  if (FPieceCompta.Action = taConsult) then
  begin
     result := '';
     Exit;
  end;
  gFindFirst := true;
  i := FindFirst( FStChemin + '*.' + FStExt, faAnyFile, vSearchRec ) ;
  if i = 0 then
  begin
    result := vSearchRec.Name;
    if not(EstFichierOk(result)) then
       result := FindNextFile(vSearchRec);
  end
  else
    result := '' ;
  gFindFirst := false;
end;

function TSaisieDoc.FindNextFile (var vSearchRec : TSearchRec) : String ;
var i : integer;
begin
  if (FPieceCompta.Action = taConsult) then
  begin
     Result := '';
     Exit;
  end;
  i := FindNext(vSearchRec);
  if i = 0 then
  begin
    result := vSearchRec.Name;
    if not(EstFichierOk(result)) then
       result := FindNextFile(vSearchRec);
  end
  else if not gFindFirst then
    result := FindFirstFile(vSearchRec)
  else
    result := '';
end;

function TSaisieDoc.FindPrevFile : integer ;
var vSearchRec : TSearchRec;
begin
  result := -1;                                 
  if FindFirstFile( FSearchRec ) = '' then Exit;
  FindFirstFile( vSearchRec ) ;
  while(FindNextFile(vSearchRec) <> '') do
  begin
    if (vSearchRec.Name = FStFichier) then
    begin
       result  := 0;
       Fichier := FSearchRec.Name;
       break;
    end;
    FindNextFile(FSearchRec);
  end;
  if result = -1 then
  begin
    FindFirstFile( vSearchRec ) ;
    if (vSearchRec.Name = FStFichier) then
    begin
       result  := 0;
       Fichier := FSearchRec.Name;
    end;
  end; 
  SysUtils.FindClose( vSearchRec );
end;

procedure TSaisieDoc.InitViewer( vPanelB, vPanelR, vPanelL, vPanelT : THPanel ;
                                 vSplitB, vSplitR, vSplitL, vSplitT : THSplitter ) ;
begin

  initialisation ;

  // Récup des composants de la fiche
  FPanelBottom     := vPanelB ;
  FSplitterBottom  := vSplitB ;

  FPanelRight      := vPanelR ;
  FSplitterRight   := vSplitR ;

  FPanelLeft       := vPanelL ;
  FSplitterLeft    := vSplitL ;

  FPanelTop        := vPanelT ;
  FSplitterTop     := vSplitT ;

  // Création des composants non géré dans Décla
//  FWebBrowser         := TWebBrowser_V1.Create( FEcran ) ;
//  FWebBrowser.Name    := 'FWebBrowser' ;

end;

procedure TSaisieDoc.DesInitViewer ;
begin
  FPanelBottom     := nil ;
  FSplitterBottom  := nil ;

  FPanelRight      := nil ;
  FSplitterRight   := nil ;

  FPanelLeft       := nil ;
  FSplitterLeft    := nil ;

  FPanelTop        := nil ;
  FSplitterTop     := nil ;

  FPanelActif      := nil ;
  FSplitterActif   := nil ;
                             
  FBoutonMenu      := nil ;
  FBoutonMenuMove  := nil ;
  FBExt            := nil ;
  FBInt            := nil ;
  FPopUpViewer     := nil ;
end;

procedure TSaisieDoc.initialisation;
begin
  FStFichier      := '' ;
  FBoFichierOk    := False ;  
  FPanelActif     := nil ;
  FSplitterActif  := nil ;
  FBoModeWeb      := False ;  
end;

procedure TSaisieDoc.ReloadParam( vTobParam : Tob );
var lStPos : string ;   
    lStChemin : string ;
begin
  if (ctxPCL in V_PGI.PGIContexte) then
  begin
    FBoAvecViewer     := GetParamSocSecur('SO_CPVIEWERACTIF',false) ;
    FBoAvecImportDoc  := GetParamSocSecur('SO_CPVIEWERIMPORTGED',false) ;
    FStExt            := '*' ; // Temporaire !!
    lStPos            := GetParamSocSecur('SO_CPVIEWERPOSITION','1BA') ;
    UpdateChemin;
  end
  else
  begin
    FBoAvecViewer     := vTobParam.GetString( 'VIEWERACTIF' ) = 'X' ;
    FBoAvecImportDoc  := vTobParam.GetString( 'VIEWERIMPORTGED' ) = 'X' ;

    // gestion du chemin des fichiers
    lStChemin := SetDossier(vTobParam.GetString( 'VIEWERDIRECTORY' ));
    if lStChemin <> FStChemin then
    begin
       Fichier := '' ;
       FStChemin := lStChemin ;
    end;
    // gestion de l'extension
    FStExt            := vTobParam.GetString( 'VIEWEREXTENSION' ) ;
    // gestion position
    lStPos            := vTobParam.GetString( 'VIEWERPOSITION' )  ;
  end;

  if (lStPos = '1BA') then
    FPosition := vpBas
  else if (lStPos = '2GA') then
    FPosition := vpGauche
  else if (lStPos = '3DR') then
    FPosition := vpDroite
  else if (lStPos = '4HO') then
    FPosition := vpHaut
  else // par défaut
    FPosition := vpBas  ;

  if Trim(FStExt) = '' then
    FStExt := '*' ;
end;

function TSaisieDoc.SetDossier (Rep : String) : String ;
var lInLg : integer ;
begin
  result := Rep ;
  if Trim(result) = '' then
     result := TCbpPath.GetPersonal ;
  lInLg := Length(result);
  if result[lInLg] <> '\' then
    result := result + '\' ;
end;

procedure TSaisieDoc.UpdateChemin ;  
var lStChemin : string ;
begin
  lStChemin := SetDossier(GetParamSocDP('SO_CPSCANPATH')) +  V_PGI.NoDossier + '\' + FPieceCompta.Info.GetString('J_NATUREJAL') + '\';
  if lStChemin <> FStChemin then
  begin
    Fichier := '';
    FStChemin := lStChemin;
  end;
end;

{$IFDEF SCANGED}
procedure TSaisieDoc.GetFichierGED ;
var vStTitre  : String;
    vStDoc    : String;
    vStGuid   : String;
    vStChemin : String;
begin
   if FPieceCompta.Action <> taCreat then Exit ;
   vStChemin := '';
   Fichier   := '';
   vStGuid   := Trim(FPieceCompta.RechGuidId(FPieceCompta.CurIdx));
   if vStGuid <> '' then
      vStDoc     := ExtraitDocument(vStGuid,vStTitre);
   if (Trim(vStDoc) <> '') and FileExists(vStDoc) then
   begin
     FStFichier   := ExtractFileName(vStDoc);
     if not FileExists (FStChemin + FStFichier) then
     begin
        vStChemin := FStChemin;
        FStChemin := ExtractFileDir (vStDoc) + '\';
        MoveDoc(vStChemin,false);
        FStChemin := vStChemin;
     end
     else
        V_GeDFiles.DeleteFile ( vStDoc );
     Fichier   := FStFichier;
   end;
end;
{$ENDIF SCANGED}

procedure TSaisieDoc.DesactiveViewer;
var    Flags, TargetFrameName, PostData, Headers: OleVariant;
begin

  if FBoModeWeb then
    begin
    if Assigned(FWebBrowser) then
      FWebBrowser.Navigate ('', Flags, TargetFrameName, PostData, Headers);
    end
  else
    begin
    if Assigned(FInternalViewer) then
      begin
      FInternalViewer.Clear;
      FreeAndNil(FInternalViewer);
      end ;
    end ;

  MontreViewer(False);

end;


procedure TSaisieDoc.SetPanelVisible ( vBoVisible : Boolean );
begin
  if assigned(FSplitterRight) then
     FSplitterRight.visible    := False ;
  if assigned(FPanelRight) then
     FPanelRight.visible       := False ;

  if assigned(FSplitterBottom) then
     FSplitterBottom.visible   := False ;
  if assigned(FPanelBottom) then
     FPanelBottom.visible      := False ;

  if assigned(FSplitterLeft) then
     FSplitterLeft.visible     := False ;
  if assigned(FPanelLeft) then
     FPanelLeft.visible        := False ;

  if assigned(FSplitterTop) then
     FSplitterTop.visible      := False ;
  if assigned(FPanelTop) then
     FPanelTop.visible         := False ;

  MontreViewer( vBoVisible ) ;
  
end;


destructor TSaisieDoc.Destroy;
begin
//  SysUtils.FindClose(FSearchRec);

  if Assigned( FInternalViewer ) then
    FreeAndNil( FInternalViewer ) ;


  if Assigned( FWebBrowser ) then
    FreeAndnil( FWebBrowser ) ;

  inherited;
end;

function TSaisieDoc.StockeFichier( vStFichier : string ) : boolean ;
var lStFic     : string ;
    lStCible   : string ;
    lPFic      : PChar;
    lPCible    : PChar;
{$IFDEF SCANGED}
    lStGuid    : string ;
    lInLig     : integer ;
    lInA       : integer ;
    lTobEcr    : Tob ;
{$ENDIF SCANGED}
begin
  // fichier en cours
  result := False ;
  if vStFichier='' then Exit ;
  if not Assigned(Piece) then Exit ;

  // répertoire
  lStCible := GetCheminStockage ;
  if not V_GedFiles.DirExists( lStCible ) then
    V_GedFiles.CreateDir( lStCible ) ;

  // Recopie du fichier dans le répertoire de stockage
  lStFic   := FStChemin + vStFichier ;
  lStCible := lStCible + '\' + vStFichier ;
  lPFic   := PChar( lStFic ) ;
  lPCible := PChar( lStCible ) ;

  // Insert en fichier GED
{$IFDEF SCANGED}
 if FBoAvecImportDoc then
    begin
    if FInGroupeActif = 0
      then lInLig := 1
      else Piece.GetBornesGroupe( FInGroupeActif, lInLig, lInA ) ;
    if ( Piece.RechGuidId(lInLig) = '' ) or
        ( PGIAsk('Il existe déjà un document associé à la pièce , voulez-vous le supprimer ?') = mrYes) then
      begin
      lTobEcr := Piece.CGetTob( lInLig ) ;
      lStGuid := AjouterFichierDansGed( lStFic, lTobEcr ) ;
      if lStGuid <> '' then
        Piece.AjouteGuidId( lInLig, lStGuid ) ;
      end ;
    end ;
{$ENDIF SCANGED}

  if FileExists( lPCible ) then
     V_GeDFiles.DeleteFile ( lPCible );
  result := V_GeDFiles.MoveFile( lPFic, lPCible );
  if not result then
    PgiInfo( V_Gedfiles.LastErrorMsg ) ;
end;

function TSaisieDoc.GetCheminStockage: string;
begin
//  result := TCbpPath.GetCegidUserLocalAppData + '\SAISIE' ;
  result := FStChemin + 'BAK' ;
end;


procedure TSaisieDoc.Refresh;
begin

  if not isActif then Exit ;

  if FichierOk then
    begin
    // Si on a cacher le viewer, on le réaffiche
    if not FPanelActif.Visible then
      MontreViewer( True ) ;
    // maj du composant
    if FBoModeWeb
      then WebRefresh
      else InternalRefresh ;
    end
  else
    // si pas de fichier a afficher, on desactive le viewer
    DesactiveViewer ;

end;

{$IFDEF SCANGED}
function TSaisieDoc.AjouterFichierDansGed( vStFichier: string ; vTobEcr : Tob ): string ;
var lStFileID  : string;
    lStDocID    : string;
    lTobDoc      : Tob;
    lTobDocGed   : Tob;
    lStErreur    : string;
begin
  result := '' ;
  if not Assigned(Piece) then Exit ;

  // Insertion du document
  lStFileID := V_GedFiles.Import( vStFichier ) ;
  if lStFileID <> '' then
    begin
    lStDocID  := '';
    lTobDoc    := Tob.Create('YDOCUMENTS', nil, -1);
    lTobDocGed := Tob.Create('DPDOCUMENT', nil, -1);
    try
        lTobDoc.LoadDb;
        lTobDoc.PutValue('YDO_LIBELLEDOC'    , Copy( GetInfoTob( vTobEcr ), 1, 70 ) ) ;
        lTobDoc.PutValue('YDO_NATDOC'        , 'Documents scanné') ;
        lTobDoc.PutValue('YDO_MOTSCLES'      , 'ECRITURE') ;
        lTobDoc.PutValue('YDO_ANNEE'         , FormatDateTime('yyyy', Date) ) ;

        lTobDocGed.LoadDb;
        lTobDocGed.PutValue('DPD_NODOSSIER', V_PGI.NoDossier) ;
//        lTobDocGed.PutValue('DPD_CODEGED', CodeGed);

        lStDocID := InsertDocumentGed( lTobDoc, lTobDocGed, lStFileID, lStErreur);

      finally
        FreeAndNil(lTobDoc);
        FreeAndNil(lTobDocGed);
      end;

    // Si fichier non référencé, on l'enlève de la Ged
    if lStDocID = ''
      then V_GedFiles.Erase ( lStFileID )
      else Result := lStDocID ;

    end;

end;
{$ENDIF SCANGED}

function TSaisieDoc.GetInfoTob(vTobEcr: Tob): string;
var lDtDateC : TDateTime ;
begin
  lDtDateC := vTobEcr.GetDateTime('E_DATECOMPTABLE') ;

  if vTobEcr.GetString('E_MODESAISIE') <> '-' then
    result := 'Document scanné Période : ' + wGetMonth( wMonth(lDtDateC) ) + ' ' + IntToStr( wYear( lDtDateC ) )
                           + ' Journal : ' + vTobEcr.GetString('E_JOURNAL')
                           + ' Folio : '   + vTobEcr.GetString('E_NUMEROPIECE')
  else
    result := 'Document scanné Date : '     + DateToStr( lDtDateC )
                           + ' Journal : '  + vTobEcr.GetString('E_JOURNAL')
                           + ' N° Pièce : ' + vTobEcr.GetString('E_NUMEROPIECE') ;
end;

function TSaisieDoc.IsActif: boolean;
begin
  result := FBoAvecViewer and Assigned( FPanelActif ) ;
end;

constructor TSaisieDoc.Create( vEcran : TForm ; vPiece : TPieceCompta ) ;
begin
  inherited Create ;

  FEcran        := vEcran ;
  FPieceCompta  := vPiece ;

end;



procedure TSaisieDoc.InternalRefresh;
var lStUrl : String ;
begin

  lStUrl := FStChemin + Fichier ;

  // Affichage
  try
    if not Assigned( FInternalViewer ) then
      begin
      FInternalViewer := TGedInternalViewer.Create(FPanelActif) ;
      FInternalViewer.OnUpdateButtons := GereEtatBoutons ;
      end ;

     FInternalViewer.ShowFile(lStUrl) ;
     FInternalViewer.FitToLarge ;
     if Assigned( PanelActif ) then
       PanelActif.setFocus ;
     // attente utilisateur
     Application.ProcessMessages;

  except
    On E:Exception do
      PGIInfo(E.Message, FEcran.Caption);
  end;

end;

procedure TSaisieDoc.WebRefresh;
var lStUrl        : String ;
    Flags, TargetFrameName, PostData, Headers: OleVariant;
begin
{*
  if Fichier=FStOldFic then
    begin
    FWebBrowser.Refresh ;
    Exit ;
    end ;
 *}
  lStUrl := 'file:///' + FStChemin + Fichier ;
  SourisSablier ;

  // Affichage
  try
    FWebBrowser.Navigate (lStUrl, Flags, TargetFrameName, PostData, Headers);
    // attente utilisateur
    Application.ProcessMessages;
  except
    On E:Exception do
      PGIInfo(E.Message, FEcran.Caption);
  end;
  SourisNormale ;

end;

procedure TSaisieDoc.ChargeDocument ( vBoSuivant : boolean ; vBoRefresh : boolean ; vBoStocke : boolean ) ;
begin

  if not IsActif then Exit ;

  // recherche du document à afficher
  if FichierOk
    // Fichier actuel ok ==> on passe au suivant
    then begin
         if vBoSuivant then
           begin              
           if vBoStocke then
              StockeFichier( Fichier ) ;
           Fichier := FindNextFile(FSearchRec) ;
           end ;
         end
    // pas de fichier actuel ==> on essaye de recharger un 1er fichier
    else begin
         Fichier := FindFirstFile(FSearchRec);
         end ;

  if vBoRefresh then refresh ;


end;

procedure TSaisieDoc.MontreViewer(vBoVal: Boolean);
begin
  if Assigned( FPanelActif ) then
    FPanelActif.visible := vBoVal ;
  if Assigned( FSplitterActif ) then
    FSplitterActif.visible := vBoVal ;
  if not FBoModeWeb then
  begin
    if not assigned(FBoutonMenu) then Exit;
    if FBoAvecViewer then
    begin
       FBoutonMenu.Visible := vBoVal;
       if assigned(FBoutonMenuMove) then
          FBoutonMenuMove.Visible := vBoVal;
    end
    else
    begin
       FBoutonMenu.Visible := false;
       if assigned(FBoutonMenuMove) then
          FBoutonMenuMove.Visible := false;
    end;
  end;

end;

procedure TSaisieDoc.SetVisibleBDoc(vBoInt : Boolean) ;
begin
  FBExt.Visible := not vBoInt;
  FBInt.Visible := vBoInt;
end;

procedure TSaisieDoc.InitPopupViewer(vActivation: Boolean);
var i,j : Integer ;
    vPopUpViewer : TMenuItem ;
begin

  if not Assigned( FPopUpViewer ) then Exit ;

  for i := 0 to FPopUpViewer.Items.Count - 1 do
    begin
    // Impression
    if FPopUpViewer.Items[i].Name = 'VIEWERPRINT' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := FichierOk and PrintOk
        else FPopUpViewer.Items[i].OnClick := ViewerPrintClick;
      end
    // Changement de régime fiscal
    else if FPopUpViewer.Items[i].Name = 'VIEWERPAN' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := FichierOk and PanOk
        else FPopUpViewer.Items[i].OnClick := ViewerPanClick;
      end
    // Modification du RIB
    else if FPopUpViewer.Items[i].Name = 'VIEWERZOOMAV' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := FichierOk and ZoomOk
        else FPopUpViewer.Items[i].OnClick := ViewerZoomAVClick ;
      end
    // Modification du RIB
    else if FPopUpViewer.Items[i].Name = 'VIEWERZOOMAR' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := FichierOk and ZoomOk
        else FPopUpViewer.Items[i].OnClick := ViewerZoomARClick ;
      end
    // Modification du RIB
    else if FPopUpViewer.Items[i].Name = 'VIEWERFITPAGE' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := FichierOk and FitToPageOk
        else FPopUpViewer.Items[i].OnClick := ViewerFitPageClick ;
      end
    // Modification du RIB
    else if FPopUpViewer.Items[i].Name = 'VIEWERFITLARGE' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := FichierOk and FitToLargeOk
        else FPopUpViewer.Items[i].OnClick := ViewerFitLargeClick ;
      end
    // Modification du RIB
    else if FPopUpViewer.Items[i].Name = 'VIEWERFIT100' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := FichierOk and FitTo100Ok
        else FPopUpViewer.Items[i].OnClick := ViewerFit100Click ;
      end
    // Document suivant
    else if FPopUpViewer.Items[i].Name = 'VIEWERNEXT' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := (FPieceCompta.Action <> taConsult)
        else FPopUpViewer.Items[i].OnClick := ViewerNextDocClick ;
      end      
    // Document précédent
    else if FPopUpViewer.Items[i].Name = 'VIEWERPREC' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := (FPieceCompta.Action <> taConsult)
        else FPopUpViewer.Items[i].OnClick := ViewerPrevDocClick ;
      end
    // retourner le document
    else if FPopUpViewer.Items[i].Name = 'VIEWERROTATE' then
      begin
      if vActivation
        then FPopUpViewer.Items[i].Enabled := FichierOk and RotateImageOk
        else FPopUpViewer.Items[i].OnClick := ViewerRotateClick ;
      end 
    // retourner le document
    else if FPopUpViewer.Items[i].Name = 'VIEWERMOVE' then
    begin
      vPopUpViewer := FPopUpViewer.Items[i];
      for j := 0 to vPopUpViewer.Count - 1 do
        if vPopUpViewer.Items[j].Name = 'VIEWERMOVEACH' then
          begin
          if vActivation
            then
            begin
               vPopUpViewer.Items[j].Visible := (ctxPCL in V_PGI.PGIContexte);
               vPopUpViewer.Items[j].Enabled := (FPieceCompta.Action <> taConsult) and (FPieceCompta.Info.GetString('J_NATUREJAL') <> 'ACH');
            end
            else vPopUpViewer.Items[j].OnClick := ViewerMoveDocACHClick ;
          end
        // retourner le document
        else if vPopUpViewer.Items[j].Name = 'VIEWERMOVEBQE' then
          begin
          if vActivation
            then
            begin
               vPopUpViewer.Items[j].Visible := (ctxPCL in V_PGI.PGIContexte);
               vPopUpViewer.Items[j].Enabled := (FPieceCompta.Action <> taConsult) and (FPieceCompta.Info.GetString('J_NATUREJAL') <> 'BQE');
            end
            else vPopUpViewer.Items[j].OnClick := ViewerMoveDocBQEClick ;
          end
        // retourner le document
        else if vPopUpViewer.Items[j].Name = 'VIEWERMOVECAI' then
          begin
          if vActivation
            then
            begin
               vPopUpViewer.Items[j].Visible := (ctxPCL in V_PGI.PGIContexte);
               vPopUpViewer.Items[j].Enabled := (FPieceCompta.Action <> taConsult) and (FPieceCompta.Info.GetString('J_NATUREJAL') <> 'CAI');
            end
            else vPopUpViewer.Items[j].OnClick := ViewerMoveDocCAIClick ;
          end
        // retourner le document
        else if vPopUpViewer.Items[j].Name = 'VIEWERMOVEOD' then
          begin
          if vActivation
            then
            begin
               vPopUpViewer.Items[j].Visible := (ctxPCL in V_PGI.PGIContexte);
               vPopUpViewer.Items[j].Enabled := (FPieceCompta.Action <> taConsult) and (FPieceCompta.Info.GetString('J_NATUREJAL') <> 'OD');
            end
            else vPopUpViewer.Items[j].OnClick := ViewerMoveDocODClick ;
          end
        // retourner le document
        else if vPopUpViewer.Items[j].Name = 'VIEWERMOVEVTE' then
          begin
          if vActivation
            then
            begin
               vPopUpViewer.Items[j].Visible := (ctxPCL in V_PGI.PGIContexte);
               vPopUpViewer.Items[j].Enabled := (FPieceCompta.Action <> taConsult) and (FPieceCompta.Info.GetString('J_NATUREJAL') <> 'VTE');
            end
            else vPopUpViewer.Items[j].OnClick := ViewerMoveDocVTEClick ;
          end
        // retourner le document
        else if vPopUpViewer.Items[j].Name = 'VIEWERMOVECORBEILLE' then
          begin
          if vActivation
            then vPopUpViewer.Items[j].Enabled := (FPieceCompta.Action <> taConsult)
            else vPopUpViewer.Items[j].OnClick := ViewerMoveDocCorbeilleClick  ;
          end ; 
        end ;
    end ;
end;

procedure TSaisieDoc.OnPopUpViewer(Sender: TObject);
begin
  InitPopupViewer ;
end;

procedure TSaisieDoc.ViewerFit100Click(Sender: TObject);
begin
  FInternalViewer.FitTo100;
end;

procedure TSaisieDoc.ViewerFitLargeClick(Sender: TObject);
begin
  FInternalViewer.FitToLarge;
end;

procedure TSaisieDoc.ViewerFitPageClick(Sender: TObject);
begin
  FInternalViewer.FitToPage;
end;

procedure TSaisieDoc.ViewerPanClick(Sender: TObject);
begin
  FInternalViewer.Pan;
end;

procedure TSaisieDoc.ViewerPrintClick(Sender: TObject);
begin
  FInternalViewer.Print;
end;

procedure TSaisieDoc.ViewerZoomArClick(Sender: TObject);
begin
  FInternalViewer.ZoomBackward;
end;

procedure TSaisieDoc.ViewerZoomAvClick(Sender: TObject);
begin
    FInternalViewer.ZoomForward
end;

procedure TSaisieDoc.SetMenuPop(vPopup: TPopupMenu ; vBouton : TToolBarButton97 ; vBoutonMove : TToolBarButton97 ) ;
begin
  FPopupViewer := vPopup ;
  FBoutonMenu  := vBouton ;
  FBoutonMenuMove := vBoutonMove ;
end;

procedure TSaisieDoc.SetBDoc( vBExt  : TToolBarButton97 ; vBInt : TToolBarButton97 ) ;
begin
  FBExt := vBExt;
  FBInt := vBInt;
end;

function TSaisieDoc.IsVisible: boolean;
begin
  result := IsActif and FPanelActif.Visible ;
end;

procedure TSaisieDoc.GereEtatBoutons(Sender: TObject; ButtonStates: TGedButtonStates);
begin
  FEtatBoutons := ButtonStates ;
end;

procedure TSaisieDoc.SetFichier(vStFic: string);
begin
  FStFichier   := vStFic ;
  FBoFichierOk := EstFichierOk( FStFichier ) ;
  if isFichier(FStFichier) and not FBoFichierOk then
     Fichier := FindNextFile(FSearchRec);
  if FBoFichierOK and (FSearchRec.Name <> FStFichier) then
     SetSearchRec;
end;

procedure TSaisieDoc.SetSearchRec;
var vStFichier : String;
begin
   if Fichier <> FindFirstFile(FSearchRec) then
   begin
      vStFichier := FindNextFile(FSearchRec);
      while (vStFichier <> FStFichier) and (vStFichier <> '') do
            vStFichier := FindNextFile(FSearchRec);
   end;
end;

function TSaisieDoc.EstFichierOk(vStFic: string): boolean;
begin
  result := false ;
  if isFichier(vStFic) then
     if ( FSearchRec.Attr and faDirectory = 0 ) and
        ( FSearchRec.Attr and faSysFile   = 0 ) and
        ( FSearchRec.Attr and faHidden    = 0 ) and
        ( FSearchRec.Attr and faVolumeID  = 0 ) then
        result := V_GEDFiles.FileExists( FStChemin + vStFic ) ;
end;

function TSaisieDoc.IsFichier(vStFic : string) : boolean;
begin
  if ( Trim( vStFic ) = '' ) or ( Trim( vStFic ) = '.' ) or ( Trim( vStFic ) = '..' ) then
    result := false
  else
    result := true;
end;

function TSaisieDoc.PanOk: boolean;
begin
  result := FEtatBoutons.PanVisible and FEtatBoutons.PanEnabled ;
end;

function TSaisieDoc.PrintOk: boolean;
begin
  result := FEtatBoutons.PrintVisible and FEtatBoutons.PrintEnabled ;
end;

function TSaisieDoc.ZoomOk: boolean;
begin
  result := FEtatBoutons.ZoomVisible and FEtatBoutons.ZoomEnabled ;
end;

function TSaisieDoc.FitTo100Ok: boolean;
begin
  result := FEtatBoutons.FitTo100Visible and FEtatBoutons.FitTo100Enabled ;
end;

function TSaisieDoc.FitToLargeOk: boolean;
begin
  result := FEtatBoutons.FitToLargeVisible and FEtatBoutons.FitToLargeEnabled ;
end;

function TSaisieDoc.FitToPageOk: boolean;
begin
  result := FEtatBoutons.FitToPageVisible and FEtatBoutons.FitToPageEnabled ;
end;

procedure TSaisieDoc.EnBasADroite;
var lScrollBarHor  : TScrollBar ;
    lScrollBarVert : TScrollBar ;
    lScrollPos     : integer ;
begin
  if FPanelActif.ControlCount > 0 then
  begin
    // IMAGES
    if FPanelActif.Controls[0] is TSmoothImage then
    begin
      // scrollbar verticale
      lScrollBarVert := TSmoothImage( FPanelActif.Controls[0] ).VScrollBar ;
      if Assigned( lScrollBarVert ) then
        lScrollBarVert.position := lScrollBarVert.max ;
      TSmoothImage( FPanelActif.Controls[0] ).ScrollBarScroll(lScrollBarVert, scEndScroll, lScrollPos);

      // scrollbar horizontale
      lScrollBarHor  := TSmoothImage( FPanelActif.Controls[0] ).HScrollBar ;
      if Assigned( lScrollBarHor ) then
        lScrollBarHor.position := lScrollBarHor.Max ;
      TSmoothImage( FPanelActif.Controls[0] ).ScrollBarScroll(lScrollBarHor, scEndScroll, lScrollPos);
    end ;
  end ;
end;

procedure TSaisieDoc.ViewerNextDocClick(Sender: TObject);
var lBoOldImportDoc : Boolean ;
begin
  lBoOldImportDoc  := FBoAvecImportDoc ;
  FBoAvecImportDoc := False ;
  ChargeDocument( True , True , False) ;
  FBoAvecImportDoc := lBoOldImportDoc ;
end;

procedure TSaisieDoc.ViewerPrevDocClick  ( Sender : TObject ) ;
var
  vSearchRecSAV : TSearchRec ;
begin
  vSearchRecSAV := FSearchRec;
  if FindPrevFile <> 0 then
  begin
     FSearchRec := vSearchRecSAV ;
     Fichier := FSearchRec.Name;
  end;
  ChargeDocument(false,true,false);
end;

procedure TSaisieDoc.ViewerMoveDocACHClick  ( Sender : TObject ) ;
begin
  MoveDoc(SetDossier(GetParamSocDP('SO_CPSCANPATH')) +  V_PGI.NoDossier + '\ACH\');
end;

procedure TSaisieDoc.ViewerMoveDocBQEClick  ( Sender : TObject ) ;
begin
  MoveDoc(SetDossier(GetParamSocDP('SO_CPSCANPATH')) +  V_PGI.NoDossier + '\BQE\');
end;

procedure TSaisieDoc.ViewerMoveDocCAIClick  ( Sender : TObject ) ;
begin
  MoveDoc(SetDossier(GetParamSocDP('SO_CPSCANPATH')) +  V_PGI.NoDossier + '\CAI\');
end;

procedure TSaisieDoc.ViewerMoveDocODClick  ( Sender : TObject ) ;
begin
  MoveDoc(SetDossier(GetParamSocDP('SO_CPSCANPATH')) +  V_PGI.NoDossier + '\OD\');
end;

procedure TSaisieDoc.ViewerMoveDocVTEClick  ( Sender : TObject ) ;
begin
  MoveDoc(SetDossier(GetParamSocDP('SO_CPSCANPATH')) +  V_PGI.NoDossier + '\VTE\');
end;

procedure TSaisieDoc.ViewerMoveDocCorbeilleClick  ( Sender : TObject ) ;
begin
  MoveDoc(GetCheminStockage + '\');
end;

procedure TSaisieDoc.MoveDoc (vRepDst : String ; vBoSuivant : Boolean ) ;
begin
  if not V_GedFiles.DirExists( vRepDst ) then
    V_GedFiles.CreateDir( vRepDst ) ;
  if FileExists( vRepDst + FStFichier ) then
     V_GeDFiles.DeleteFile ( vRepDst + FStFichier );
  if not V_GeDFiles.MoveFile( FStChemin + FStFichier, vRepDst + FStFichier ) then
    PgiInfo( V_Gedfiles.LastErrorMsg ) ;
  if vBoSuivant then
     ChargeDocument(true,true,false);
end;

procedure TSaisieDoc.EnHautAGauche;
var lScrollBarHor  : TScrollBar ;
    lScrollBarVert : TScrollBar ;
    lScrollPos     : integer ;
begin
  if FPanelActif.ControlCount > 0 then
   begin
   // IMAGES
   if FPanelActif.Controls[0] is TSmoothImage then
     begin
     // scrollbar verticale
     lScrollBarVert := TSmoothImage( FPanelActif.Controls[0] ).VScrollBar ;
     if Assigned( lScrollBarVert ) then
       lScrollBarVert.position := 0 ;
     TSmoothImage( FPanelActif.Controls[0] ).ScrollBarScroll(lScrollBarVert, scEndScroll, lScrollPos);

     // scrollbar horizontale
     lScrollBarHor  := TSmoothImage( FPanelActif.Controls[0] ).HScrollBar ;
     if Assigned( lScrollBarHor ) then
       lScrollBarHor.position := 0 ;
     TSmoothImage( FPanelActif.Controls[0] ).ScrollBarScroll(lScrollBarHor, scEndScroll, lScrollPos);
     end ;
   end ;
end;

procedure TSaisieDoc.EnBas( vBoLarge : Boolean ) ;
var lScrollBarVert : TScrollBar ;
    lScrollPos     : integer ;
begin
  if FPanelActif.ControlCount > 0 then
   begin
   // IMAGES
   if FPanelActif.Controls[0] is TSmoothImage then
     begin
     // scrollbar verticale
     lScrollBarVert := TSmoothImage( FPanelActif.Controls[0] ).VScrollBar ;
     if Assigned( lScrollBarVert ) then
       begin
       if  vBoLarge
         then lScrollBarVert.position := lScrollBarVert.position + lScrollBarVert.LargeChange
         else lScrollBarVert.position := lScrollBarVert.position + lScrollBarVert.SmallChange ;
       TSmoothImage( FPanelActif.Controls[0] ).ScrollBarScroll(lScrollBarVert, scEndScroll, lScrollPos);
       end ;
     end ;
   end ;
end;

procedure TSaisieDoc.EnHaut ( vBoLarge : Boolean ) ;
var lScrollBarVert : TScrollBar ;
    lScrollPos     : integer ;
begin
  if FPanelActif.ControlCount > 0 then
   begin
   // IMAGES
   if FPanelActif.Controls[0] is TSmoothImage then
     begin
     // scrollbar verticale
     lScrollBarVert := TSmoothImage( FPanelActif.Controls[0] ).VScrollBar ;
     if Assigned( lScrollBarVert ) then
       begin
       if  vBoLarge
         then lScrollBarVert.position := lScrollBarVert.position - lScrollBarVert.LargeChange
         else lScrollBarVert.position := lScrollBarVert.position - lScrollBarVert.SmallChange ;
       TSmoothImage( FPanelActif.Controls[0] ).ScrollBarScroll(lScrollBarVert, scEndScroll, lScrollPos);
       end ;
     end ;
   end ;
end;

function TSaisieDoc.DeplacementOk: boolean;
begin
  result := IsActif and IsVisible and (FPanelActif.Controls[0] is TSmoothImage) ;
end;

procedure TSaisieDoc.ADroite(vBoLarge: Boolean);
var lScrollBar : TScrollBar ;
    lScrollPos : integer ;
begin
  if FPanelActif.ControlCount > 0 then
   begin
   // IMAGES
   if FPanelActif.Controls[0] is TSmoothImage then
     begin
     // scrollbar verticale
     lScrollBar := TSmoothImage( FPanelActif.Controls[0] ).HScrollBar ;
     if Assigned( lScrollBar ) then
       begin
       if  vBoLarge
         then lScrollBar.position := lScrollBar.position + lScrollBar.LargeChange
         else lScrollBar.position := lScrollBar.position + lScrollBar.SmallChange ;
       TSmoothImage( FPanelActif.Controls[0] ).ScrollBarScroll(lScrollBar, scEndScroll, lScrollPos);
       end ;
     end ;
   end ;
end;

procedure TSaisieDoc.AGauche(vBoLarge: Boolean);
var lScrollBar : TScrollBar ;
    lScrollPos : integer ;
begin
  if FPanelActif.ControlCount > 0 then
   begin
   // IMAGES
   if FPanelActif.Controls[0] is TSmoothImage then
     begin
     // scrollbar verticale
     lScrollBar := TSmoothImage( FPanelActif.Controls[0] ).HScrollBar ;
     if Assigned( lScrollBar ) then
       begin
       if  vBoLarge
         then lScrollBar.position := lScrollBar.position - lScrollBar.LargeChange
         else lScrollBar.position := lScrollBar.position - lScrollBar.SmallChange ;
       TSmoothImage( FPanelActif.Controls[0] ).ScrollBarScroll(lScrollBar, scEndScroll, lScrollPos);
       end ;
     end ;
   end ;
end;

procedure TSaisieDoc.ZoomImage(vBoPlus: boolean);
var X,Y    : integer ;
    lScale : Real ;
begin
  if not ZoomImageOk then Exit ;
  X := FPanelActif.width div 2 ;
  Y := FPanelActif.height div 2 ;
  lScale := TSmoothImage( FPanelActif.Controls[0] ).Scale ;
  if vBoPlus
    then TSmoothImage( FPanelActif.Controls[0] ).Zoom(X, Y, lScale * 1.25)
    else TSmoothImage( FPanelActif.Controls[0] ).Zoom(X, Y, lScale * 0.75) ;
end;

function TSaisieDoc.ZoomImageOk: boolean;
begin
  result := IsActif and IsVisible and (FPanelActif.Controls[0] is TSmoothImage) ;
end;

procedure TSaisieDoc.RotateImage;
var  bmp1 : TBitMap;
     bmp2 : TBitMap;
begin

  if not ZoomImageOk then Exit ;
Rotate90 ;
Exit ;
  // bmp2 est le bitmap résultat
  bmp2 := TBitmap.Create;
  bmp1 := TSmoothImage( FPanelActif.Controls[0] ).Picture ;

  bmp2.width  := bmp1.width;
  bmp2.height := bmp1.height;
  bmp2.Canvas.StretchDraw( Rect( bmp1.width-1 , Bmp1.height-1 , -1 , -1 ) , bmp1 ) ;

  TSmoothImage( FPanelActif.Controls[0] ).picture := bmp2 ;

  freeAndNil(bmp2);

end;

function TSaisieDoc.RotateImageOk: boolean;
begin
  result := IsActif and IsVisible and (FPanelActif.Controls[0] is TSmoothImage) ;
end;

procedure TSaisieDoc.ViewerRotateClick(Sender: TObject);
begin
  if RotateImageOk then
    RotateImage ;
end;

procedure TSaisieDoc.Rotate90;
type
  TManoRGB  = packed record
                 rgb    : TRGBTriple;
                 dummy  : byte;
              end;
  TRGBArray = ARRAY[0..0] OF TRGBTriple;
  pRGBArray = ^TRGBArray;
var
  aStream : TMemorystream;          // zone mémoire
  header  : TBITMAPINFO;            // header bitmap
  dc      : hDC;                    // ressource pour GetDIBits
  P       : ^TManoRGB;              // pointeur vers 4 octets
  RowOut  :  pRGBArray;               // pointeur vers 3 octets
  x,y,h,w : integer;
  Bitmap  : TBitmap;
  bmp1    : TBitmap;
begin
  aStream := nil;
  Bitmap := TBitmap.Create;
  try
    try
      // Bitmap.canvas.Draw(0,0,bmp1);
      {JP : problème avec le Draw !!}
      bmp1 := TSmoothImage( FPanelActif.Controls[0] ).Picture ;

      Bitmap.width  := bmp1.width;
      Bitmap.height := bmp1.height;
      Bitmap.Canvas.StretchDraw(Rect(0, 0, bmp1.width-1, Bmp1.height-1), bmp1);
      {JP : problème avec le Draw !!}

      Bitmap.pixelformat := pf24bit;
      aStream := TMemoryStream.Create;                 // réservation mémoire
      aStream.SetSize(Bitmap.Height*Bitmap.Width * 4); // chaque pixel = 4 octets
      with header.bmiHeader do begin                   // bitmap mémoire
        biSize := SizeOf(TBITMAPINFOHEADER);
        biWidth := Bitmap.Width;
        biHeight := Bitmap.Height;
        biPlanes := 1;
        biBitCount := 32;              // 32 bits par pixel = 4 octets
        biCompression := 0;
        biSizeimage := aStream.Size;
        biXPelsPerMeter :=1;
        biYPelsPerMeter :=1;
        biClrUsed :=0;
        biClrImportant :=0;
      end;
      dc := GetDC(0);                  // folklore des ressources GDI windows
      P  := aStream.Memory;
      // copie du bitmap dans le flux (stream). Passe de 3 à 4 octets
      GetDIBits(dc,Bitmap.Handle,0,Bitmap.Height,P,header,dib_RGB_Colors);
      ReleaseDC(0,dc);                 // folklore des ressources GDI windows
      w := bitmap.Height;
      h := bitmap.Width;
      bitmap.Width  := w;              // permute largeur / hauteur
      bitmap.height := h;
      for y := 0 to (h-1) do     // boucle pilotée par y et x coodonnées sortie
      begin
        rowOut := Bitmap.ScanLine[y]; // sortie 3 octets = 24 bits par pixels
        P  := aStream.Memory;
        inc(p,y);                    // p = adresse ligne du stream
        for x := 0 to (w-1) do
        begin
          rowout[x] := p^.rgb;        // copie 3 octets sur les 4 du stream
          inc(p,h);                   // parcours de la ligne du stream
        end;
      end;

      TSmoothImage( FPanelActif.Controls[0] ).picture := bitmap ;

      //  TImage(GetControl('IMAGE')).picture.bitmap.assign(bitmap);
      //  TImage(GetControl('IMAGE')).Refresh;
    finally
      if assigned(bitmap) then freeandnil(bitmap);
      if assigned(aStream) then aStream.Free;
    end;
  except
     on E : Exception do
            PGIError('Impossible de retourner le document : '#13#13 + E.Message);
  end;
end;

end.


