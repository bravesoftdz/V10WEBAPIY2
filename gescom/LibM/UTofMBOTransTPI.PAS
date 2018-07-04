{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Source TOF de la FICHE : MBOTRANSTPI ()
Mots clefs ... : TOF;MFORETRANSJOUR
*****************************************************************}
Unit UTOFMBOTRANSTPI ;

Interface

Uses
  Controls, Classes, ComCtrls, graphics, sysutils, forms, stdctrls, FileCtrl,
{$IFNDEF EAGLCLIENT}
  dbTables,
{$ENDIF}
  HTB97, HStatus, HMsgBox, HCtrls, HEnt1, UTOF, UTOB, Paramsoc,
  TPI_Base;

Type
  TEtatTransTpi = ( ttpiNONE, ttpiSTART, ttpiSTOP );

Type
  TOF_MBOTRANSTPI = Class (TOF)
  private
    AnimateTPI        : TAnimate ;
    CompteRendu       : TRichEdit ;
    TypeTPI           : string ;
    Tpi               : TTPI ;
    EtatTransTpi      : TEtatTransTpi ;
    NbConnexion       : integer ;
    NbAbort           : integer ;
    NbLigne           : integer ;
    Tob_InvTrans      : TOB ;        // TOB stockant l'inventaire transmis
    Tob_Lig           : TOB ;        // TOB Fille stockant une ligne d'inventaire
    Tob_PieceTPI      : TOB ;        // TOB stockant les articles du TPI
    TOBArticles       : TOB ;
    CodeDepot         : string ;
    CodeEmpl          : string ;
    SelectFourniss    : string ;
    NaturePiece       : string ;
    Domaine           : string ;
    NumLigInv         : integer ;
    NbArtInv          : integer ;
    NbCabInconnu      : integer ;
    NbCabFerme        : integer ;
    NbCabMauvaisFourn : integer ;
    AfficheTrace      : boolean ;
    BPieceTPI         : boolean ;
    procedure AfficheMemo(Texte : String ; Style : TFontStyles ; Saut, Traduire : Boolean ) ;
    procedure ActiveSablier( Active : Boolean ) ;
    procedure OnStartCom( var Cancel : Boolean ) ;
    procedure OnReceiveCom( Data : String ; var Cancel : Boolean ) ;
    procedure TraiteTrameSymbol ( Data : string ) ;
    procedure TraiteTrameOpticon ( Data : string ) ;
    procedure OnStopCom( Code : Integer ; Libelle : String ) ;
    procedure LanceCommunication ;
    procedure LanceMagestel ;
    function  TraiteFichierMagestel : boolean ;
    procedure BATTENTEClick(Sender: TObject) ;
    procedure BARRETClick(Sender: TObject) ;
    procedure GCTYPETPIChange(Sender: TObject) ;
    procedure EnregistrementInvTrans(CodDepot_Defaut,LibTrans:String; Tob_Trans:TOB) ;
    procedure InsertNewArt(CodeBarre : string; Q : TQuery; TOBArt : TOB; Qte : Integer);
    procedure EnregistrementArtCB(Tob_Trans:TOB) ;
    function  RechCodeDepotValide (CodDepot_Tpi,CodDepot_Defaut : String) : String ;
    function  RechCodeTransDepot (CodDepot : String) : String ;
    procedure CopieFichierInvTrans ( CodDepot_Defaut, LibTrans : string ; Tob_Trans : TOB) ;
    procedure BQUITClick(Sender: TObject) ;
  public
    procedure OnLoad                     ; override ;
    procedure OnArgument( Arg : String ) ; override ;
    procedure OnClose                    ; override ;
  end ;

Implementation

Uses
  EntGC, FactUtil , FactArticle, FactTOB;

Const
  ACK = $06 ;
  HT  = $09 ;
  VT  = $0B ;
  Chr_BTQ : String = chr(ACK);
  Chr_QTE : String = chr(HT) ;
  Chr_EMP : String = chr(VT) ;

  // libellés des messages
  TexteMessage: array[1..2] of string = (
    {1}   'Vous devez renseigner l''établissement',
    {2}   'Vous devez installer le programme pour OPTICON de MAGESTEL'
  );

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Affiche un texte dans le mémo.
Mots clefs ... : 
*****************************************************************}
procedure TOF_MBOTRANSTPI.AfficheMemo ( Texte : string ; Style : TFontStyles ; Saut, Traduire : boolean ) ;
var Stg : string ;
begin
  if CompteRendu <> Nil then
  begin
    if CompteRendu.CanFocus then CompteRendu.SetFocus ;
    if Saut then CompteRendu.lines.Append ('') ;
    CompteRendu.SelAttributes.Style := Style ;
    if Traduire then
      Stg := TraduireMemoire (Texte)
    else
      Stg := Texte ;
    CompteRendu.lines.Append (Stg) ;
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Active ou le non le sablier
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.ActiveSablier ( Active : boolean ) ;
Var CurCursor, NewCursor : TCursor ;
    Cmp                  : TComponent ;
    Ind                  : Integer ;
begin
  if Active then
  begin
    SourisSablier ;
    CurCursor := crDefault ;
    NewCursor := crAppStart ;
  end else
  begin
    SourisNormale ;
    CurCursor := crAppStart ;
    NewCursor := crDefault ;
  end ;
  // Changement du curseur des boutons
  for Ind := 0 to Ecran.ComponentCount -1 do
  begin
    Cmp := Ecran.Components[Ind] ;
    if (Cmp is TToolBarButton97) and (TToolBarButton97(Cmp).Cursor = CurCursor) then
      TToolBarButton97(Cmp).Cursor := NewCursor ;
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Démarrage de la communication
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.OnStartCom ( var Cancel : boolean ) ;
begin
  AfficheMemo('Début de la communication', [fsBold], false, True) ;
  if AnimateTPI <> Nil then
  begin
    AnimateTPI.Visible := False ;
    AnimateTPI.Active := False ;
  end ;
  Inc (NbConnexion) ;
  NbLigne := 0 ;
  InitMove (100, Ecran.Caption) ;
  // Création de la Tob de réception de l'inventaire transmis
  Tob_InvTrans := TOB.Create ('Lignes Inv Trans', nil, -1);
  Tob_Lig := Nil ;
  if not BPieceTPI then CodeDepot := GetControlText ('DEPOT') ;
  CodeEmpl := '' ;
  NumLigInv := 0 ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Réception d'une trame
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.OnReceiveCom ( Data : string ; var Cancel : boolean ) ;
begin
  if (AfficheTrace) then AfficheMemo ('>'+Data+'<', [fsBold], false, False) ;
  Inc (NbLigne) ;
  MoveCur (False) ;
  if Length (Data) < 1 then Exit ;
  // Chargement en TOB de l'information reçue
  if TypeTPI = '004' then
    TraiteTrameOpticon (Data)
  else
    TraiteTrameSymbol (Data) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Traitement d'une trame d'un SYMBOL/TELXON
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.TraiteTrameSymbol ( Data : string ) ;
var Qte, CodeBarre : string ;
    NewCB : Boolean;
    OldQte : Double;
begin
  if Data[1] = Chr_BTQ then
  begin
    // Réception du code Boutique (Dépot)
    CodeDepot := Trim (Copy (Data, 2, Length (Data) -1)) ;
  end else
  if Data[1] = Chr_QTE then
  begin
    // Réception d'une quantité
    Qte := Trim (Copy (Data, 2, Length (Data) -1)) ;
    if Tob_Lig <> Nil then
    begin
      OldQte := Tob_Lig.GetValue ('GIN_QTEINV');
      if (TypeTPI <> '001') and (OldQte > 1) then
        Tob_Lig.PutValue ('GIN_QTEINV', (OldQte - 1) + StrToFloat (Qte))
      else Tob_Lig.PutValue ('GIN_QTEINV', StrToFloat (Qte)) ;
    end;
  end else
  if Data[1] = Chr_EMP then
  begin
    // Réception d'un code Emplacement
    CodeEmpl := Trim (Copy (Data, 2, Length (Data) -1)) ;
  end else
  begin
    // Réception d'un Code à barres
    CodeBarre := Trim (Data);
    NewCB := True;
    if TypeTPI <> '001' then
    begin
      Tob_Lig := Tob_InvTrans.FindFirst(['GIN_CODEBARRE','GIN_EMPLACEMENT'], [CodeBarre,CodeEmpl], False);
      if Tob_Lig <> Nil then NewCB := False;
    end;
    if NewCB then
    begin
      Tob_Lig := TOB.CREATE ('TRANSINVLIG', Tob_InvTrans, -1) ;
      Tob_Lig.initValeurs ;
      Tob_Lig.PutValue ('GIN_CODEBARRE', CodeBarre) ;
      Tob_Lig.PutValue ('GIN_QTEINV', 1) ;
      if Not BPieceTPI then
      begin
        Tob_Lig.PutValue ('GIN_DEPOT', CodeDepot) ;
        Tob_Lig.PutValue ('GIN_EMPLACEMENT', CodeEmpl) ;
        Inc (NumLigInv) ;
        Tob_Lig.PutValue ('GIN_NUMLIGNE', NumLigInv) ;
      end;
    end
    else Tob_Lig.PutValue ('GIN_QTEINV',Tob_Lig.GetValue ('GIN_QTEINV') + 1) ;

  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 08/01/2003
Description .. : Traitement d'une trame d'un OPTICON/MAGESTEL
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.TraiteTrameOpticon ( Data : string ) ;
var CodeBarre : string;
    NewCB : boolean;
begin
  if Data[1] = 'E' then
  begin
    // Réception d'un code Emplacement
    CodeEmpl := Trim (Copy (Data, 2, 4)) ;
    // Réception du code Boutique/Dépot (utilisation du champ opérateur)
    CodeDepot := Trim (Copy (Data, 10, 3)) ;
  end else
  if Data[1] = 'D' then
  begin
    // Réception d'un Code à barres
    CodeBarre := Trim (Copy (Data, 2, MaxInt));
    NewCB := True;
    if TypeTPI <> '001' then
    begin
      Tob_Lig := Tob_InvTrans.FindFirst(['GIN_CODEBARRE','GIN_EMPLACEMENT'], [CodeBarre,CodeEmpl], False);
      if Tob_Lig <> Nil then NewCB := False;
    end;
    if NewCB then
    begin
      // Réception d'un Code à barres
      Tob_Lig := TOB.CREATE ('TRANSINVLIG', Tob_InvTrans, -1) ;
      Tob_Lig.initValeurs ;
      Tob_Lig.PutValue ('GIN_CODEBARRE', CodeBarre) ;
      Tob_Lig.PutValue ('GIN_QTEINV', 1) ;
      if not BPieceTPI then
      begin
        Tob_Lig.PutValue ('GIN_DEPOT', CodeDepot) ;
        Tob_Lig.PutValue ('GIN_EMPLACEMENT', CodeEmpl) ;
        Inc (NumLigInv) ;
        Tob_Lig.PutValue ('GIN_NUMLIGNE', NumLigInv) ;
      end ;
    end
    else Tob_Lig.PutValue ('GIN_QTEINV',Tob_Lig.GetValue ('GIN_QTEINV') + 1) ; 
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Arrêt de la communication
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.OnStopCom ( Code : integer ; Libelle : string ) ;
begin
  if Code=0 then
    AfficheMemo ('Fin de la communication.', [fsBold], False, True)
  else
    AfficheMemo ('Fin de la communication - erreur n°'+IntToStr(Code)+' '+Libelle, [fsBold], False, False) ;
  AfficheMemo (IntToStr(NbLigne)+' Lignes reçues', [], False, False) ;
  FiniMove ;
  if (AnimateTPI <> Nil) and (GetControlText ('GCREMISEATTENTE') = 'X') then
  begin
    AnimateTPI.Visible := True ;
    AnimateTPI.Active := True ;
  end ;

  if Code <> 0 then
  begin
    Inc (NbAbort) ;
    Tob_InvTrans.Free;
    exit ;
  end ;

  // Enregistrement de l'inventaire transmis
  try
    AfficheMemo ('Enregistrement des données en cours ...', [fsBold], True, True) ;
    NbArtInv := 0 ;
    NbCabInconnu := 0 ;
    NbCabFerme := 0 ;
    NbCabMauvaisFourn := 0 ;
    if BPieceTPI then
      EnregistrementArtCB (Tob_InvTrans)
    else if GetCheckBoxState ('GCENREGFICHIER') = cbChecked then
      CopieFichierInvTrans (GetControlText ('DEPOT'), GetControlText ('LIBTRANS'), Tob_InvTrans)
    else
      EnregistrementInvTrans(GetControlText ('DEPOT'), GetControlText ('LIBTRANS'), Tob_InvTrans) ;
    AfficheMemo ('Enregistrement terminé.', [fsBold], False, True) ;
  finally
    Tob_InvTrans.Free;
    if BPieceTPI then
      AfficheMemo (IntToStr (NbArtInv)+' Articles référencés', [], False, False)
    else
      AfficheMemo (IntToStr (NbArtInv)+' Articles inventoriés', [], False, False) ;
    if NbCabInconnu > 0 then
    begin
      if NbCabInconnu=1 then
        AfficheMemo ('Attention, 1 Code à barres est inconnu', [], False, False)
      else
        AfficheMemo ('Attention, '+IntToStr(NbCabInconnu)+' Codes à barres sont inconnus', [], False, False) ;
    end;
    if NbCabFerme > 0 then
      AfficheMemo ('Attention, '+IntToStr(NbCabFerme)+' Codes à barres sont inconnu', [], False, False);
    if NbCabMauvaisFourn > 0 then
      AfficheMemo ('Attention, '+IntToStr(NbCabMauvaisFourn)+' Codes à barres sont inconnu', [], False, False);
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Lancement de la communication avec le protocole SYMBOL
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.LanceCommunication ;
var CodeErr            : integer ;
    RemiseAttente      : boolean ;
    PortTpi, ParamsTpi : string ;
begin
  // Récupération des paramètres de communication
  PortTpi := GetControlText ('GCPORTCOM') ;
  ParamsTpi := GetControlText ('GCPORTBAUDS')     + ';'
             + GetControlText ('GCPORTPARITE')    + ';'
             + GetControlText ('GCPORTDATABITS')  + ';'
             + GetControlText ('GCPORTSTOPBITS')  + ';'
             + GetControlText ('GCPORTHNDSHK') ;
  RemiseAttente := (GetControlText ('GCREMISEATTENTE') = 'X') ;
  AfficheTrace  := (GetControlText ('GCTRACE') = 'X') ;
  // Lancement de la communication
  Tpi := TTPI.Create (Application, TypeTpi) ;
  try
    CodeErr := Tpi.ChargePortetParams (PortTpi, ParamsTpi) ;
    if CodeErr = 0 then
    begin
      Tpi.OnStartCom   := OnStartCom ;
      Tpi.OnStopCom    := OnStopCom ;
      Tpi.OnReceiveCom := OnReceiveCom ;
      CodeErr := Tpi.Transaction (RemiseAttente) ;
    end ;
  finally
    Tpi.Free ;
  end ;
  if CodeErr <> 0 then
    AfficheMemo (TPI_MsgErrDefaut(CodeErr), [fsBold], False, True) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 08/01/2003
Modifié le ... : 08/01/2003
Description .. : Donne le nom du dossier qui contient le  programme de
Suite ........ : communication  avec le protocole MAGESTEL
Mots clefs ... :
*****************************************************************}
function RepertoireProgMagestel : string ;
begin
  Result := IncludeTrailingBackslash (GetCurrentDir) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/12/2002
Modifié le ... : 08/01/2003
Description .. : Donne le nom du programme de communication avec le
Suite ........ : protocole MAGESTEL
Mots clefs ... :
*****************************************************************}
function NomProgMagestel ( AvecRepertoire : boolean ) : string ;
begin
  if AvecRepertoire then
    Result := RepertoireProgMagestel
  else
    Result := '' ;
  Result := Result + 'TPI_OPTICON.EXE' ;  // 'PHL1700.EXE' ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 08/01/2003
Modifié le ... : 08/01/2003
Description .. : Donne le nom du dossier qui contient le fichier reçu du
Suite ........ : TPI/OPTICON
Suite ........ :
Suite ........ : Par défaut sous NT :
Suite ........ : C:\Documents and Settings\NomUtilisateur\Application
Suite ........ : Data\Cegid\NomApplication
Mots clefs ... :
*****************************************************************}
function RepertoireFichierMagestel : string ;
var NomDir : string ;
    Ok     : boolean ;
begin
  Ok := False ;
  NomDir := Trim (GetEnvVar ('APPDATA')) ;
  if (NomDir <> '') and (DirectoryExists (NomDir)) then
  begin
    NomDir := NomDir + '\' + Apalatys + '\' + NomHalley ;
    if not DirectoryExists (NomDir) then ForceDirectories (NomDir) ;
    Ok := DirectoryExists (NomDir) ;
  end ;
  if not Ok then NomDir := RepertoireProgMagestel ;
  Result := IncludeTrailingBackslash (NomDir) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/12/2002
Modifié le ... : 08/01/2003
Description .. : Donne le nom du fichier qui contient les éléments reçu du
Suite ........ : TPI/OPTICON
Mots clefs ... :
*****************************************************************}
function NomFichierMagestel ( AvecRepertoire : boolean ) : string ;
begin
  if AvecRepertoire then
    Result := RepertoireFichierMagestel
  else
    Result := '' ;
  Result := Result + 'SYMBOL.CR' ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/12/2002
Modifié le ... : 08/01/2003
Description .. : Lancement de la communication avec le protocole MAGESTEL
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.LanceMagestel ;
var sCmd, Stg, AppDir, DataDir : string ;
begin
  if not FileExists (NomProgMagestel (True)) then
  begin
    PgiBox (TexteMessage[2] + ' : ' + NomProgMagestel (False)) ;
    Exit ;
  end ;
  AppDir := RepertoireProgMagestel ;
  DataDir := RepertoireFichierMagestel ;
  repeat
    sCmd := NomProgMagestel (True) ;
    // Récupération des paramètres de communication
    Stg := GetControlText('GCPORTCOM') ;     // port de communication
    if Stg = 'CO1' then
      sCmd := sCmd + ' /P1'
    else if Stg = 'CO2' then
      sCmd := sCmd + ' /P2'
    else if Stg = 'CO3' then
      sCmd := sCmd + ' /P3'
    else if Stg = 'CO4' then
      sCmd := sCmd + ' /P4' ;
    Stg := GetControlText('GCPORTBAUDS') ;   // vitesse de communication
    if Stg = '004' then
      sCmd := sCmd + ' /B1200'
    else if Stg = '005' then
      sCmd := sCmd + ' /B2400'
    else if Stg = '006' then
      sCmd := sCmd + ' /B4800'
    else if Stg = '007' then
      sCmd := sCmd + ' /B9600'
    else if Stg = '009' then
      sCmd := sCmd + ' /B19200'
    else if Stg = '010' then
      sCmd := sCmd + ' /B38400' ;
    //sCmd := sCmd + ' /T15' ;  // TimeOut de 15 secondes
    // Lancement de la communication
    if DataDir <> AppDir then ChDir (DataDir) ;
    FileExecAndWait (sCmd) ;
    if DataDir <> AppDir then ChDir (AppDir) ;
  until not (TraiteFichierMagestel) or (GetControlText ('GCREMISEATTENTE') <> 'X') ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/12/2002
Modifié le ... : 08/01/2003
Description .. : Traitement du fichier reçu avec le protocole MAGESTEL
Mots clefs ... :
*****************************************************************}
function TOF_MBOTRANSTPI.TraiteFichierMagestel : boolean ;
var NomFichier, NouveauNom : TFileName ;
    EnregInv               : string ;
    FichierInv             : textfile ;
    Cancel                 : boolean ;
begin
  Cancel := False ;
  OnStartCom (Cancel) ;
  NomFichier := NomFichierMagestel (True) ;
  if FileExists (NomFichier) then
  begin
{$I-}      // directive de compilation pour inactiver les exceptions et utiliser IOResult car pb en CWAS !!
    AssignFile (FichierInv, NomFichier) ;
    FileMode := 0;   // Accès en lecture seule au fichier
    try
      Reset (FichierInv);
      while not EOF (FichierInv) and (IOResult = 0) do
      begin
        Readln (FichierInv, EnregInv) ;
        OnReceiveCom (EnregInv, Cancel) ;
      end ;
    finally
      // Fermeture du fichier
      CloseFile (FichierInv) ;
      OnStopCom (0, '') ;
    end ;
{$I+}
    // renommage du fichier en SYMBOL.OLD
    NouveauNom := ChangeFileExt (NomFichier, '.OLD') ;
    if FileExists (NouveauNom) then DeleteFile (NouveauNom) ;
    RenameFile (NomFichier, NouveauNom) ;
    Result := True ;
  end else
  begin
    OnStopCom (1, NomFichier + ' : ' + TraduireMemoire ('fichier inexistant')) ;
    Result := False ;
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Bouton mise en attente
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.BATTENTEClick(Sender: TObject) ;
var Ctl : TControl ;
begin
  if not (EtatTransTpi = ttpiNONE) then Exit ;
  if not BPieceTPI and (GetControlText('DEPOT') = '') then
  begin
    PgiInfo (TexteMessage[1], Ecran.Caption);
    SetFocusControl('DEPOT');
    Exit ;
  end;
  EtatTransTpi := ttpiSTART ;
  NbConnexion := 0 ;
  NbAbort := 0;
  NbLigne := 0 ;
  // Démarrage de la communication
  AfficheMemo('Démarrage de la communication', [fsUnderline, fsItalic], True, True) ;
  ActiveSablier(True) ;
  // Activation de l'animation
  if AnimateTPI = nil then
  begin
    if ctxFO in V_PGI.PGIContexte then
      Ctl := GetControl('PBOUTONS')
    else
      Ctl := GetControl('PBOUTON');
    if Ctl <> nil then
    begin
      AnimateTPI := TAnimate.Create(Ctl) ;
      AnimateTPI.Parent := TWinControl(Ctl) ;
      AnimateTPI.Left := 4 ;
      AnimateTPI.Top := 1 ;
      AnimateTPI.Width := 80 ;
      AnimateTPI.Height := 50 ;
      AnimateTPI.CommonAVI := aviFindComputer ;
      AnimateTPI.StopFrame := 29 ;
      AnimateTPI.Center := True ;
      AnimateTPI.Visible := True ;
      AnimateTPI.Active := True ;
    end ;
  end else
  begin
    AnimateTPI.Visible := True ;
    AnimateTPI.Active := True ;
  end ;
  // Modification de la configuration interdite
  SetControlEnabled('PCONFIG', False) ;
  // Lancement de la communication
  TypeTpi := GetControlText('GCTYPETPI') ;
  if TypeTPI = '004' then
    LanceMagestel
  else
    LanceCommunication ;
  // Fin de la communication
  ActiveSablier(False) ;
  AfficheMemo('Fin de la communication', [fsUnderline, fsItalic], True, True) ;
  AfficheMemo(IntToStr(NbConnexion)+' connexions', [], False, False) ;
  if NbAbort > 0 then AfficheMemo(IntToStr(NbAbort)+' connexions interrompues', [], False, False) ;
  if AnimateTPI <> nil then
  begin
    AnimateTPI.Visible := False ;
    AnimateTPI.Active := False ;
  end ;
  SetControlEnabled('PCONFIG', True) ;
  EtatTransTpi := ttpiNONE ;
  if ctxFO in V_PGI.PGIContexte then
  begin
    Ctl := GetControl('BARRET') ;
    if (Ctl <> Nil) and (Ctl is TToolBarButton97) then
      TToolBarButton97(Ctl).Down := True ;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Bouton Arrêt
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.BARRETClick(Sender: TObject) ;
begin
  if not (EtatTransTpi = ttpiSTART) then Exit ;
  EtatTransTpi := ttpiSTOP ;
  // Arrêt de la communication
  if Tpi <> Nil then Tpi.Arret := True ;
  AfficheMemo('Arrêt de la communication demandé', [fsBold, fsItalic], True, True) ;
  if AnimateTPI <> Nil then
  begin
    AnimateTPI.Visible := False ;
    AnimateTPI.Active := False ;
  end ;
  SetControlEnabled('PCONFIG', True) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 08/01/2003
Description .. : Procédure d'enregistrement de l'inventaire transmis, en Table
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.EnregistrementInvTrans(CodDepot_Defaut,LibTrans:String; Tob_Trans:TOB) ;
var i    : Integer ;
    TobEntete : TOB ;
    TobE,TobL : TOB ;
    CodDepot_Tpi    : String ;  // Code Dépot fournit par le TPI
    CodDepot_Valide : String ;  // Code Dépot validé : Code Dépot du TPI si existant, sinon le Dépôt par défaut saisi en entête d'écran
    CodTrans : String ;
    NumLigne : Integer ;
    Q : TQuery ;
begin
  if Tob_Trans.Detail.Count = 0 then Exit;
  CodDepot_Tpi := '' ;
  CodDepot_Valide := '' ;
  CodTrans := '' ;
  NumLigne := 0 ;
  TobEntete := TOB.Create('Entete Inv Transmis', nil, -1);
  TobE := Nil ;

  InitMove(Tob_Trans.Detail.Count,'');
  for i:=0 to Tob_Trans.Detail.Count-1 do
  begin
    TobL := Tob_Trans.Detail[i] ;
    if (i = 0) or (TobL.GetValue('GIN_DEPOT')<>CodDepot_Tpi) then
    begin
      CodDepot_Tpi := TobL.GetValue('GIN_DEPOT') ;

      // Recherche du code Dépôt sur lequel sera affecté l'inventaire
      CodDepot_Valide := RechCodeDepotValide (CodDepot_Tpi, CodDepot_Defaut) ;

      // Recherche si un N° de trnsmission a déjà été affecté à ce dépot
      TobE := TobEntete.FindFirst(['GIT_DEPOT'], [CodDepot_Valide], False) ;
      if TobE<>nil then
      begin
        CodTrans := TobE.GetValue('GIT_CODETRANS') ;
        NumLigne := TobE.GetValue('NUMLIG_COURS') ;
      end else
      begin
        // Recherche du premier N° de transmission libre pour le dépôt
        CodTrans := RechCodeTransDepot (CodDepot_Valide) ;
        AfficheMemo ('Code Transmission : '+CodTrans, [], False, True) ;
        NumLigne := 0 ;

        // Création d'une Fille dans la TOB Entête des inventaires transmis
        TobE := TOB.Create('TRANSINVENT', TobEntete, -1) ;
        TobE.PutValue('GIT_DEPOT', CodDepot_Valide);
        TobE.PutValue('GIT_CODETRANS', CodTrans);
        TobE.PutValue('GIT_LIBELLE', LibTrans);
        TobE.PutValue('GIT_ETABLISSEMENT', CodDepot_Valide);
        //TobE.PutValue('GIT_DATECREATION', NowH);
        //TobE.PutValue('GIT_DATEMODIF', NowH);
        TobE.PutValue('GIT_INTEGRATION', '-');
        TobE.PutValue('GIT_CREATEUR', V_PGI.USER);
        TobE.PutValue('GIT_UTILISATEUR', V_PGI.USER);

        TobE.AddChampSup('NUMLIG_COURS', True) ;
        TobE.PutValue('NUMLIG_COURS', NumLigne) ;
      end ;
    end ;
    Inc(NumLigne) ;
    TobL.PutValue('GIN_DEPOT', CodDepot_Valide) ;
    TobL.PutValue('GIN_CODETRANS', CodTrans) ;
    TobL.PutValue('GIN_NUMLIGNE', NumLigne) ;
    if TobE <> Nil then TobE.PutValue('NUMLIG_COURS', NumLigne) ;
    NbArtInv := NbArtInv + TobL.GetValue('GIN_QTEINV') ;

    // Recherche de l'article à partir du Code à barres
    Q := OpenSQL('Select GA_ARTICLE,GA_CODEARTICLE from ARTICLE where GA_CODEBARRE="'+TobL.GetValue('GIN_CODEBARRE')+'"',True) ;
    if Not Q.EOF then
    begin
      TobL.PutValue('GIN_ARTICLE', Q.FindField('GA_ARTICLE').AsString) ;
      TobL.PutValue('GIN_CODEARTICLE', Q.FindField('GA_CODEARTICLE').AsString) ;
    end else
     Inc(NbCabInconnu) ;
    Ferme(Q) ;
    MoveCur(False) ;
  end ;
  FiniMove;

  // Mise à jour de la Table TRANSINVENT : Entête des inventaires transmis
  TobEntete.InsertDB(nil, False);
  TobEntete.Free ;

  // Mise à jour de la Table TRANSINVLIG
  Tob_Trans.InsertDB(nil, True);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. :
Mots clefs ... :
*****************************************************************}
function RetourneLibelleAvecDimensions(RefArtDim:String;TOBArt:TOB): string;
var k : integer;
    Grille,CodeDim,LibDim: String ;
begin
LibDim:='';
if TOBArt<>Nil then
  begin
  for k:=1 to 5 do
    begin
    Grille:=TOBArt.GetValue('GA_GRILLEDIM'+IntToStr(k)) ;
    CodeDim:=TOBArt.GetValue('GA_CODEDIM'+IntToStr(k)) ;
    if ((Grille<>'') and (CodeDim<>'')) then
      LibDim:=LibDim+'  '+RechDom('GCGRILLEDIM'+IntToStr(k),Grille,True)+' '+GCGetCodeDim(Grille,CodeDim,k);
    end;
  end;
Result:=LibDim;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.InsertNewArt(CodeBarre : string; Q : TQuery; TOBArt : TOB; Qte : Integer);
var TOB_Fille : TOB;
    RefArt,Designation : string;
begin
if (Q<>nil) then
  begin TOBArt:=CreerTOBArt(TOBArticles); TOBArt.SelectDB('',Q); end;
if (TOBArt<>nil) then
  begin
  RefArt := TOBArt.getvalue('GA_ARTICLE');
  if TOBArt.getvalue('GA_STATUTART')='DIM' then
    Designation := Copy(TOBArt.getvalue('GA_LIBELLE')+RetourneLibelleAvecDimensions(RefArt,TOBArt),1,70)
  else Designation := TOBArt.getvalue('GA_LIBELLE');
  end
else Exit ;

TOB_Fille := TOB.Create('',Tob_PieceTPI,-1);
With TOB_Fille do
  begin
  AddChampSupValeur('RefArticle',RefArt); AddChampSupValeur('CB',CodeBarre);
  AddChampSupValeur('LIBELLE',Designation); AddChampSupValeur('QTE',Qte);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : 
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Procédure de stockage des articles pour être utilisé dans la 
Suite ........ : saisie par code barre
Mots clefs ... : 
*****************************************************************}
procedure TOF_MBOTRANSTPI.EnregistrementArtCB(Tob_Trans:TOB) ;
var i    : Integer ;
    CodeBarre, WhereNat : string ;
    TobL,TOB_Fille,TOBNat,TOBG,TOBArt : TOB ;
    Q : TQuery ;
    isFerme : boolean;
    Qte : Integer;
begin
  if Tob_Trans.Detail.Count = 0 then Exit;

  InitMove(Tob_Trans.Detail.Count,'');
  for i:=0 to Tob_Trans.Detail.Count-1 do
  begin
    TobL := Tob_Trans.Detail[i] ;
    Qte := TobL.GetValue('GIN_QTEINV') ;

    //Recherche dans TOBGSA
    CodeBarre := TobL.GetValue('GIN_CODEBARRE') ;
    TOB_Fille := Tob_PieceTPI.FindFirst(['CB'],[CodeBarre],False);
    if TOB_Fille <> nil then
    //Si l'on trouve, on ajoute la nouvelle quantité à l'ancienne trouvée
    begin
      TOB_Fille.PutValue('QTE',TOB_Fille.GetValue('QTE')+Qte); NbArtInv := NbArtInv + Qte ;
    end else
    //Sinon on recherche l'article dans la base
    begin
      //Recherche Article dans TOBArticle par son Code Barre
      TOBArt:=TOBArticles.FindFirst(['GA_CODEBARRE'],[CodeBarre],False);
      if TOBArt<>Nil then
      begin
        InsertNewArt(CodeBarre, Nil, TOBArt, Qte); NbArtInv := NbArtInv + Qte ;
      end else
      begin
        //Dans le cas d'une gestion Mono-fournisseur, seuls les articles du fournisseur sont dispo.
        WhereNat:=FabricWhereNatArt(NaturePiece,DOMAINE,SelectFourniss) ;
        if WhereNat<>'' then WhereNat:=' AND '+WhereNat ;
        // Recherche via code barre
        Q:=OpenSQL('Select * from ARTICLE Where GA_CODEBARRE="'+CodeBarre+'" '+WhereNat,True) ;
        try
          if Not Q.EOF then
          begin
            isFerme:=(Q.FindField('GA_FERME').AsString='X');
            if isFerme then
            begin
              TOBNat:=VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'],[NaturePiece],False) ;
              if TOBNat <> Nil then
                TOBG:=TOBNat.FindFirst(['GAP_NATUREPIECEG','GAP_ARTICLE'],[NaturePiece,Q.FindField('GA_ARTICLE').AsString],False)
              else
                TOBG := Nil;
              if TOBG<>Nil then IsFerme:=(TOBG.GetValue('GAP_SUSPENSION')='X');
            end;
            if isFerme then
              Inc(NbCabFerme)
            else
            begin
              InsertNewArt(CodeBarre, Q, Nil, Qte);
              NbArtInv := NbArtInv + Qte ;
            end;
          end else
          begin
            if SelectFourniss='' then
              Inc(NbCabInconnu)
            else
              Inc(NbCabMauvaisFourn);
          end;
        finally
          Ferme(Q);
        end;
      end;
    end;
    MoveCur(False) ;
  end ;
  FiniMove;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Fonction de recherche du code Dépôt sur lequel sera
Suite ........ : affecté l'inventaire
Mots clefs ... :
*****************************************************************}
function TOF_MBOTRANSTPI.RechCodeDepotValide (CodDepot_Tpi,CodDepot_Defaut : String) : String ;
var CodDepot : String ;
    lg : Integer ;
begin
   Result := CodDepot_Defaut ;

   if CodDepot_Tpi<>'' then
      begin
      // Recherche si le code Dépôt retourné par le TPI existe en Table
      if ExisteSQL('Select * from DEPOTS where GDE_DEPOT="'+CodDepot_Tpi+'"') then
         Result := CodDepot_Tpi
      else
         begin
         // On essait de formatter le code Dépôt en ajoutant des zéros à gauche
         lg := length(CodDepot_Tpi);
         if lg < 3 then
            begin
            CodDepot := Copy('00',1,3-lg) + CodDepot_Tpi ;
            if ExisteSQL('Select * from DEPOTS where GDE_DEPOT="'+CodDepot+'"') then
               Result := CodDepot ;
            end;
         end;
      end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Fonction de recherche du premier N° de transmission libre
Suite ........ : pour le dépôt
Mots clefs ... :
*****************************************************************}
function TOF_MBOTRANSTPI.RechCodeTransDepot ( CodDepot : string ) : string ;
var FNextCode : Integer;
    QQ : TQuery;
begin
  QQ := OpenSQL('Select MAX(GIT_CODETRANS) AS NEXTCODE from TRANSINVENT where GIT_DEPOT="'+CodDepot+'"',True);
  if QQ.FindField('NEXTCODE').AsString='' then
   FNextCode := 1 else FNextCode := QQ.FindField('NEXTCODE').AsInteger + 1;
  Ferme(QQ);
  Result := Format('%.3d',[FNextCode]) ;
{  NextCode:=1 ;
  while ExisteSQL('Select GIT_CODETRANS from TRANSINVENT '+
                   'where GIT_DEPOT="'+CodDepot+'" and '+
                   'GIT_CODETRANS="'+Format('%.3d',[NextCode])+'"') do inc(NextCode);
  Result := Format('%.3d', [NextCode]) ;}
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : Procédure d'enregistrement dans un fichier de l'inventaire
Suite ........ : transmis
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.CopieFichierInvTrans ( CodDepot_Defaut, LibTrans : string ; Tob_Trans : TOB) ;
var Ind                  : integer ;
    QteInv               : double ;
    TOBL                 : TOB ;
    CodDepot_Tpi         : string ;  // Code Dépot fournit par le TPI
    CodDepot_Valide      : string ;  // Code Dépot validé : Code Dépot du TPI si existant, sinon le Dépôt par défaut saisi en entête d'écran
    EnregInv, NomFichier : string  ;
    FichierInv           : textfile ;
begin
  if Tob_Trans.Detail.Count = 0 then Exit ;
  // Ouverture du fichier
  NomFichier := GetControlText ('GCNOMFICENREG') ;
  if NomFichier = '' then Exit ;
  AssignFile (FichierInv, NomFichier) ;
  if FileExists (NomFichier) then Append (FichierInv) else Rewrite (FichierInv) ;
  try
    // Recopie de la TOB dans le fichier au format import
    InitMove (Tob_Trans.Detail.Count, '');
    CodDepot_Tpi := CodDepot_Defaut ;
    CodDepot_Valide := CodDepot_Defaut ;
    for Ind := 0 to Tob_Trans.Detail.Count -1 do
    begin
      TOBL := Tob_Trans.Detail[Ind] ;
      QteInv := TOBL.GetValue ('GIN_QTEINV') ;
      if TOBL.GetValue ('GIN_DEPOT') <> CodDepot_Tpi then
      begin
        // Recherche du code Dépôt sur lequel sera affecté l'inventaire
        CodDepot_Tpi := TOBL.GetValue ('GIN_DEPOT') ;
        CodDepot_Valide := RechCodeDepotValide (CodDepot_Tpi, CodDepot_Defaut) ;
      end ;
      EnregInv := 'INVC1 ' + Format ('%-3.3s'  , [CodDepot_Valide]) +
                             Format ('%-18.18s', [TOBL.GetValue('GIN_CODEBARRE')]) +
                             Format ('%6.0f'   , [QteInv]) +
                             Format ('%-17.17s', [TOBL.GetValue('GIN_EMPLACEMENT')]) ;

      EnregInv := Trim (EnregInv) ;
      writeln (FichierInv, EnregInv);
      MoveCur (False) ;
    end ;
  finally
    // Fermeture du fichier
    CloseFile (FichierInv) ;
  end ;
  FiniMove ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.OnLoad ;
var Ctl : TControl ;
begin
  Inherited ;
  // Initialisation
  EtatTransTpi := ttpiNONE ;
  if ctxFO in V_PGI.PGIContexte then
  begin
    SetControlVisible ('PBOUTONS', True);
    //SetControlVisible ('PBOUTON', False);
    SetControlVisible ('BCONFIG2', False); SetControlVisible ('BGO', False);
    SetControlVisible ('BSTOP', False); SetControlVisible ('BFERME', False);
  end;
  // Libellé établissement
  if BPieceTPI then
  begin
    SetControlVisible('PENTETE',False);
    //SetControlVisible('PBOUTON',True);
    SetControlVisible('GCENREGFICHIER',False);
    SetControlVisible('TGCNOMFICENREG',False);
    SetControlVisible('GCNOMFICENREG',False);
  end else
  begin
    if not VH_GC.GCMultiDepots then
    begin
      Ctl := GetControl('TDEPOT') ;
      if (Ctl <> Nil) and (Ctl is THLabel) then
        THLabel(Ctl).Caption := 'Etablissement' ;
    end ;
    // Reprise du code Dépot ou Etablissement par défaut
    Ctl := GetControl('DEPOT') ;
    if (Ctl <> Nil) and (Ctl is THValComboBox) then
      THValComboBox(Ctl).Value := VH_GC.GCDepotDefaut ;
  end;

  // Pointeur sur le memo
  Ctl := GetControl('COMPTERENDU') ;
  if (Ctl <> Nil) and (Ctl is TRichEdit) then
  begin
    CompteRendu := TRichEdit(Ctl) ;
    CompteRendu.Text := '' ;
  end ;

  // Configuration par défaut
  SetControlText('GCTYPETPI', GetParamSoc('SO_GCTYPETPI')) ;
  SetControlText('GCPORTCOM', GetParamSoc('SO_GCPORTCOM')) ;
  SetControlText('GCPORTBAUDS', GetParamSoc('SO_GCPORTBAUDS')) ;
  SetControlText('GCPORTPARITE', GetParamSoc('SO_GCPORTPARITE')) ;
  SetControlText('GCPORTDATABITS', GetParamSoc('SO_GCPORTDATABITS')) ;
  SetControlText('GCPORTSTOPBITS', GetParamSoc('SO_GCPORTSTOPBITS')) ;
  SetControlText('GCPORTHNDSHK', GetParamSoc('SO_GCPORTHNDSHK')) ;
  if GetParamSoc('SO_GCREMISEATTENTE') = True then
    SetControlText('GCREMISEATTENTE', 'X')
  else
    SetControlText('GCREMISEATTENTE', '-') ;
  if GetParamSoc('SO_GCTRACE') = True then
    SetControlText('GCTRACE', 'X')
  else
    SetControlText('GCTRACE', '-') ;

  if GetControlText('GCTYPETPI')      = '' then SetControlText('GCTYPETPI', '001') ;
  if GetControlText('GCPORTCOM')      = '' then SetControlText('GCPORTCOM', 'CO1') ;
  if GetControlText('GCPORTBAUDS')    = '' then SetControlText('GCPORTBAUDS', '004') ;
  if GetControlText('GCPORTPARITE')   = '' then SetControlText('GCPORTPARITE', '003') ;
  if GetControlText('GCPORTDATABITS') = '' then SetControlText('GCPORTDATABITS', '003') ;
  if GetControlText('GCPORTSTOPBITS') = '' then SetControlText('GCPORTSTOPBITS', '001') ;
  if GetControlText('GCPORTHNDSHK')   = '' then SetControlText('GCPORTHNDSHK', '001') ;

  if not BPieceTPI then SetFocusControl('DEPOT');

  // Branchement des fonctions OnClick sur les boutons
  if ctxFO in V_PGI.PGIContexte then
  begin
    Ctl := GetControl('BATTENTE') ;
    if (Ctl <> Nil) and (Ctl is TToolBarButton97) then
      TToolBarButton97(Ctl).Onclick := BATTENTEClick ;
    Ctl := GetControl('BARRET') ;
    if (Ctl <> Nil) and (Ctl is TToolBarButton97) then
      TToolBarButton97(Ctl).Onclick := BARRETClick ;
    Ctl := GetControl('BQUIT') ;
    if (Ctl <> Nil) and (Ctl is TToolBarButton97) then
      TToolBarButton97(Ctl).Onclick := BQUITClick ;
  end else
  begin
    Ctl := GetControl('BGO') ;
    if (Ctl <> Nil) and (Ctl is TToolBarButton97) then
      TToolBarButton97(Ctl).Onclick := BATTENTEClick ;
    Ctl := GetControl('BSTOP') ;
    if (Ctl <> Nil) and (Ctl is TToolBarButton97) then
      TToolBarButton97(Ctl).Onclick := BARRETClick ;
  end ;

  // Branchement de la fonctions OnChange sur le type de TPI
  Ctl := GetControl('GCTYPETPI') ;
  if (Ctl <> Nil) and (Ctl is THValComboBox) then
    THValComboBox(Ctl).OnChange := GCTYPETPIChange ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : En sortie d'une nouvelle sélection du type de TPI
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.GCTYPETPIChange(Sender: TObject) ;
var Stg : string ;
begin
  Stg := GetControlText ('GCTYPETPI') ;
  if (Stg = '002') or (Stg = '003') then
  begin                                      // TELXON
    SetControlText ('GCPORTBAUDS', '004') ;     // 1200 bauds
    SetControlText ('GCPORTPARITE', '003') ;    // parité paire
    SetControlText ('GCPORTDATABITS', '003') ;  // 7 bits
    SetControlText ('GCPORTSTOPBITS', '001') ;  // 1 bit d'arrêt
    SetControlText ('GCPORTHNDSHK', '001') ;    // Aucun
  end else
  if Stg = '004' then
  begin                                      // OPTICON
    SetControlText ('GCPORTBAUDS', '009') ;     // 19200 bauds
    SetControlText ('GCPORTPARITE', '001') ;    // sans parité
    SetControlText ('GCPORTDATABITS', '004') ;  // 8 bits
    SetControlText ('GCPORTSTOPBITS', '001') ;  // 1 bit d'arrêt
    SetControlText ('GCPORTHNDSHK', '001') ;    // Aucun
    if not FileExists (NomProgMagestel (True)) then
      PgiBox (TexteMessage[2] + ' : ' + NomProgMagestel (False)) ;
  end else
  begin                                      // SYMBOL
    SetControlText ('GCPORTBAUDS', '007') ;     // 9600 bauds
    SetControlText ('GCPORTPARITE', '001') ;    // sans parité
    SetControlText ('GCPORTDATABITS', '004') ;  // 8 bits
    SetControlText ('GCPORTSTOPBITS', '001') ;  // 1 bit d'arrêt
    SetControlText ('GCPORTHNDSHK', '001') ;    // Aucun
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.OnArgument( Arg : String ) ;
var Critere,ChampCritere,ValeurCritere,VenteAchat : string;
    i_ind : integer;
begin
  inherited;
  SelectFourniss := '' ; BPieceTPI := False ;
  Repeat
    Critere:=Trim(ReadTokenSt(Arg)) ;
    if Critere<>'' then
      begin
      i_ind:=pos('=',Critere);
      if i_ind<>0 then
        begin
        ChampCritere:=copy(Critere,1,i_ind-1);
        ValeurCritere:=copy(Critere,i_ind+1,length(Critere));
        if ChampCritere='PIECETPI' then BPieceTPI:=True
        else if ChampCritere='NATUREPIECE' then NaturePiece:=ValeurCritere
        else if ChampCritere='DOMAINE' then Domaine:=ValeurCritere
        else if ChampCritere='FOURNISSEUR' then
          begin
          //Dans le cas d'une gestion Mono-fournisseur, seuls les articles du fournisseur sont dispo.
          VenteAchat:=GetInfoParPiece(NaturePiece,'GPP_VENTEACHAT')  ;
          if (VenteAchat = 'ACH') and (GetParamsoc('SO_MONOFOURNISS') = True) and
             (GetInfoParPiece(NaturePiece,'GPP_ARTFOURPRIN')='X') then SelectFourniss := ValeurCritere ;
          end;

        end;
      end;
  until Critere = '';

  if BPieceTPI then
    begin
    TOBArticles := TOB(LaTOB.Data);
    Tob_PieceTPI := LaTOB;
    LaTOB.Data := Nil; LaTOB := Nil ;
    end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 23/12/2002
Modifié le ... : 23/12/2002
Description .. : OnClose
Mots clefs ... :
*****************************************************************}
procedure TOF_MBOTRANSTPI.OnClose ;
begin
  Inherited ;
  if not (EtatTransTpi = ttpiNONE) then
  begin
    LastError := -1 ;                     // la fiche ne sera pas fermée
    Exit ;
  end ;

  // Sauvegarde du paramétrage dans les paramètres Société
  SetParamSoc('SO_GCTYPETPI', GetControlText('GCTYPETPI')) ;
  SetParamSoc('SO_GCPORTCOM', GetControlText('GCPORTCOM')) ;
  SetParamSoc('SO_GCPORTBAUDS', GetControlText('GCPORTBAUDS')) ;
  SetParamSoc('SO_GCPORTPARITE', GetControlText('GCPORTPARITE')) ;
  SetParamSoc('SO_GCPORTDATABITS', GetControlText('GCPORTDATABITS')) ;
  SetParamSoc('SO_GCPORTSTOPBITS', GetControlText('GCPORTSTOPBITS')) ;
  SetParamSoc('SO_GCPORTHNDSHK', GetControlText('GCPORTHNDSHK')) ;
  if GetControlText('GCREMISEATTENTE') = 'X' then
    SetParamSoc('SO_GCREMISEATTENTE', True)
  else
    SetParamSoc('SO_GCREMISEATTENTE', False) ;
  if GetControlText('GCTRACE') = 'X' then
    SetParamSoc('SO_GCTRACE', True)
  else
    SetParamSoc('SO_GCTRACE', False) ;
end ;

procedure TOF_MBOTRANSTPI.BQUITClick(Sender: TObject) ;
begin
  TToolBarButton97(GetControl('BFerme')).Click;
end;

Initialization
  registerclasses ( [ TOF_MBOTRANSTPI ] ) ;
end.
