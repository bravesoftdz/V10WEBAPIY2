{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPECHEMULTI ()
Mots clefs ... : TOF;CPECHEMULTI
*****************************************************************}
Unit CPECHEMULTI_TOF ;

Interface

Uses Controls,           // TKeyEvent
     Classes,            // TShiftState
     sysutils,           // FreeAndNil
     HCtrls,             // THGrid
     HEnt1,              // TActionFiche
     HMsgBox,            // THMsgBox
     Saisutil,           // NbJoursOK, RDevise
     hTB97,              // TToolBarButton97
     UTOF,               // TOF
     uLibPieceCompta;    // TMultiEche

procedure CPLanceFiche_EcheMulti( vStArg : String ) ;

Type
  TOF_CPECHEMULTI = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnAfterFormShow          ;
  private

    // Composants
    FListe        : THGrid ;           // Grille de saisie
    MsgBox        : THMsgBox;
    FModeFinal    : THValComboBox;
    FModeInit     : THValComboBox;
    FAPayer       : THNumEdit;
    FSaisi        : THNumEdit;
    FReste        : THNumEdit;

    // Boutons
    BRepart       : TToolbarButton97 ;
    BSolde        : TToolbarButton97 ;
    BZoom         : TToolbarButton97 ;

    // Objet de gestion
    FMultiEche    : TMultiEche ;       // Objet de gestion de l'échéance

    // Indicateurs
    FAction       : TActionFiche ;     // Type d'action fiche
    FBoInit       : Boolean ;          // Init interface

    // Evt AGL
    FOldKeyDown   : TKeyEvent ;

    // ======= Initialisations
    procedure InitMsgBox;
    procedure InitComposants ;
    procedure InitEvenements ;
    procedure InitVariables ;
    procedure InitInterface ;
    procedure AnalyseParametre( vStArg : String ) ;

    // ======= Evènements fiches
    procedure MODEFINALChanged  ( Sender : TObject ) ;
    procedure BSoldeClick       ( Sender : TObject ) ;
    procedure BRepartClick      ( Sender : TObject ) ;
    procedure BZoomClick        ( Sender : TObject ) ;
    procedure FListeCellExit    ( Sender : TObject ; var ACol, ARow : Longint ; var Cancel : Boolean ) ;
    procedure FormCloseQuery    ( Sender : TObject ; var CanClose : Boolean ) ;
    procedure OnKeyDown         ( Sender: TObject; var Key: Word; Shift: TShiftState) ;

    // affichage
    procedure AfficheEcheances( vMultiEche : TMultiEche );

    // Validation des données
    procedure VerifLimiteModepaie(ARow: Integer);
    function  RecupEche: boolean;
    procedure AfficheTotal;
    function  GetLimiteMP   ( vStCodeMP : String ) : Single ;
    function  GetModeRempMP ( vStCodeMP : String ) : String ;
    procedure SetModePaie   ( vStCodeMP : String ; vRow : Integer ) ;
    function  GetDev : RDevise ;

  end ;

Implementation

uses
     Grids,            // GoRowSelect
{$IFDEF EAGLCLIENT}
     MainEagl,
{$ELSE}
     Fe_Main,
{$ENDIF}

   {$IFDEF MODENT1}
     CPProcGen,
     CPTypeCons,
   {$ENDIF MODENT1}
     uTob,
     Vierge,             // TFVierge
     AglInit,            // theTob
     Ent1,               // ChangeMask
     FichComm,           // FicheRegle_AGL
     UtilPGI,            // EuroToDevise
     uLibEcriture        // CSetMontants...
     ;

// ==============================================================================================
procedure CPLanceFiche_EcheMulti( vStArg : String ) ;
begin
  if pos('ACTION=', vStArg) = 0 then
     vStArg := 'ACTION=CONSULTATION;' + vStArg ;
  AGLLanceFiche( 'CP', 'CPECHEMULTI', '', '', vStArg ) ;
end ;
// ==============================================================================================



// ============================================================================================================

procedure TOF_CPECHEMULTI.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPECHEMULTI.OnArgument (S : String ) ;
begin
  Inherited ;

  FBoInit := True ;

  // Objet gérant les échéances passé en paramètres via TheTob :
  FMultiEche := TMultiEche( TFVierge(Ecran).laTof.laTob ) ;

  AnalyseParametre( S ) ;
  InitVariables ;
  InitMsgBox;
  InitComposants ;
  InitEvenements ;
  InitInterface ;

  // Affichage des données
  AfficheEcheances( FMultiEche ) ;

end ;

procedure TOF_CPECHEMULTI.OnClose ;
begin
  Inherited ;
  if Assigned( MsgBox ) then
    FreeAndNil( MsgBox );
end ;

procedure TOF_CPECHEMULTI.MODEFINALChanged(Sender: TObject);
var lMultiEche : TMultiEche ;
    lStCode    : string ;
begin

  if not FMultiEche.EstModifiable then Exit;

  lStcode := FModeFinal.Value ;
  if not FMultiEche.Info.LoadModeRegle( lStcode ) then Exit ;

  if not FBoInit then
    begin

    lMultiEche := TMultiEche.Create( nil, FMultiEche.Piece, FMultiEche.TobEcr, FMultiEche.Action, mfAlone ) ;
    if lMultiEche.NbEche = 0 then
      lMultiEche.NewEche ;
    lMultiEche.SetMontant( 1, FMultiEche.MontantSaisi ) ;

    if CCalculEche ( lMultiEche, 0, lStCode, lMultiEche.Info ) then
      AfficheEcheances( lMultiEche ) ;

    FreeAndNil( lMultiEche ) ;

    end ;

end;

procedure TOF_CPECHEMULTI.AfficheEcheances( vMultiEche : TMultiEche );
var i        : Integer ;
    lTobEche : Tob ;
begin

  // Pb affichage ValCombo
  if Assigned( FListe.ValCombo ) and FListe.ValCombo.Visible then
    begin
    FListe.ValCombo.Visible := False ;
    FListe.ValCombo.Refresh ;    
    end ;

  // Init Affichage
  for i := 1 to 12 do
    begin
    FListe.Cells[0, i] := IntToStr(i);
    FListe.Cells[1, i] := '';
    FListe.Cells[2, i] := '';
    FListe.Cells[3, i] := '';
    end ;

  // Remplissage du Grid
  for i := 1 to ( vMultiEche.nbEche ) do
    begin
    lTobEche := vMultiEche.GetEche( i ) ;
    SetModePaie( lTobEche.GetString('E_MODEPAIE') , i);
    FListe.Cells[2, i] := StrfMontant( lTobEche.GetDouble('E_DEBITDEV') + lTobEche.GetDouble('E_CREDITDEV'), 15, GetDev.Decimale, '', TRUE);
    FListe.Cells[3, i] := lTobEche.GetString('E_DATEECHEANCE');
    end ;

  AfficheTotal;

end;

function TOF_CPECHEMULTI.GetLimiteMP(vStCodeMP: String): Single;
var lTobMP : Tob ;
begin

  result := 0.0 ;
  if vStCodeMP = '' then Exit ;

  if not FMultiEche.Info.LoadModePaie(vStCodeMP) then Exit ;
  lTobMP := FMultiEche.Info.ModePaie.Item ;
  if lTobMP.GetString('MP_CONDITION') = 'X'
      then result := lTobMP.GetValue('MP_MONTANTMAX') ;

  result := EuroToDevise(  result, GetDev.Taux, GetDev.Quotite, GetDev.Decimale ) ;

end;

function TOF_CPECHEMULTI.GetModeRempMP(vStCodeMP: String): String;
var lTobMP : Tob ;
begin

  result := '' ;
  if vStCodeMP = '' then Exit ;

  if not FMultiEche.Info.LoadModePaie(vStCodeMP) then Exit ;
  lTobMP := FMultiEche.Info.ModePaie.Item ;

  if lTobMP.GetValue('MP_CONDITION') = 'X'
      then result := lTobMP.GetString('MP_REMPLACEMAX') ;

end;

procedure TOF_CPECHEMULTI.SetModePaie(vStCodeMP: String; vRow: Integer);
begin

  if vStCodeMP = '' then
    begin
    FListe.CellValues [1, vRow] := '' ;
    FListe.Cells      [1, vRow] := '' ;
    end
  else
    begin
    // Maj code de l'item sélectionné dans la combo
    FListe.CellValues [1, vRow] := vStCodeMP;
    // Maj texte affiché
    FListe.Cells      [1, vRow] := RechDom( 'TTMODEPAIE', vStCodeMP, False ) ;
    end ;

end;

procedure TOF_CPECHEMULTI.VerifLimiteModepaie(ARow: Integer);
var lMRRemplacement : string;
    lMontantMax     : Single;
begin
  if FAction = taConsult then Exit;
  if FListe.CellValues[1, Arow] = '' then Exit ;

  lMRRemplacement := GetModeRempMP( FListe.CellValues[1, Arow] ) ;
  lMontantMax     := GetLimiteMP(   FListe.CellValues[1, Arow] ) ;

  // Dépasse-t-on la limite fixé pour le mode de paiement
  if ( lMontantMax > 0 ) and ( lMontantMax < Valeur(FListe.Cells[2, ARow]) ) and ( lMRRemplacement <> '' ) then
    begin
    case MsgBox.execute(1, '', '') of
      // On remplace le mode de paiement
      mrYes : SetModePaie( lMRRemplacement, ARow ) ;
      // On stagne au max pour le mode de paiement
      mrNo  : FListe.Cells[2, ARow] := StrfMontant(lMontantMax, 15, GetDev.Decimale, '', TRUE);
      end;
    end;

end;

procedure TOF_CPECHEMULTI.BRepartClick(Sender: TObject);
var i    : Integer ;
    nb   : Integer ;
    Last : Integer ;
begin
  if not FMultiEche.EstModifiable then Exit;
  if FAction = taConsult then Exit;
  nb := 0;
  Last := 1;
  for i := 1 to 12 do
    if FListe.Cells[1, i] <> '' then inc(nb);
  if nb <= 0 then exit;
  for i := 1 to 12 do
    if FListe.Cells[1, i] <> '' then
      begin
      FListe.Cells[2, i] := StrfMontant(FAPayer.Value / nb, 15, GetDev.Decimale, '', TRUE);
      Last := i;
      end;
  AfficheTotal;
  FListe.Row := Last;
  BSoldeClick(nil);
end;

procedure TOF_CPECHEMULTI.BSoldeClick(Sender: TObject);
begin
  if FAction = taConsult then Exit;
  if GoRowSelect in FListe.Options then Exit;
  FListe.Cells[2, FListe.Row] := StrfMontant( Valeur(FListe.Cells[2, FListe.Row]) + FReste.Value, 15, GetDev.Decimale, '', TRUE);
  AfficheTotal;
end;

procedure TOF_CPECHEMULTI.BZoomClick(Sender: TObject);
begin
  FicheRegle_AGL( FModeFinal.Value, False, taConsult);
end;

procedure TOF_CPECHEMULTI.InitVariables;
begin
//  FMultiEche.Info.ModePaie.LoadAll ;
end;

procedure TOF_CPECHEMULTI.InitInterface;
begin

  ChangeMask( FAPayer,  GetDev.Decimale, GetDev.Symbole ) ;
  ChangeMask( FReste ,  GetDev.Decimale, GetDev.Symbole ) ;
  ChangeMask( FSaisi ,  GetDev.Decimale, GetDev.Symbole ) ;

  // mise en place des mode
  FModeInit.Value := FMultiEche.ModeInit ;
  if FMultiEche.ModeFinal <> ''
    then FModeFinal.Value := FMultiEche.ModeFinal ;
//    else FModeFinal.itemIndex := -1 ;

  // Gestion permission
  FModeFinal.Enabled := ( FAction <> taConsult ) and FMultiEche.EstModifiable ;
//  FApayer.Enabled := ( FAction <> taConsult ) ;
  BSolde.Enabled := ( FAction <> taConsult ) ;
  BRepart.Enabled := ( FAction <> taConsult ) and FMultiEche.EstModifiable ;

end;

procedure TOF_CPECHEMULTI.InitComposants;
begin

  // ==============================
  // ==== Pointeurs composants ====
  // ==============================
  FListe     := THGrid ( GetControl('FLISTE', True) ) ;
  FModeFinal := THValComboBox( GetControl('MODEFINAL', True) ) ;
  FModeInit  := THValComboBox( GetControl('MODEINIT', True) ) ;
  FAPayer    := THNumEdit ( GetControl('APAYER', True) ) ;
  FSaisi     := THNumEdit ( GetControl('SAISI', True) ) ;
  FReste     := THNumEdit ( GetControl('RESTE', True) ) ;
  BRepart    := TToolbarButton97 ( GetControl('BREPART', True) ) ;
  BSolde     := TToolbarButton97 ( GetControl('BSOLDE', True) ) ;
  BZoom      := TToolbarButton97 ( GetControl('BZOOM', True) ) ;

  // ============================
  // ==== Paramétrage Grille ====
  // ============================
  if FAction = taConsult then
    FListe.Options := FListe.Options - [goEditing] + [goRowSelect];
  // Mode de paiement
//  FListe.ColTypes[1]   := '' ;         // combo
  FListe.ColFormats[1] := 'CB=TTMODEPAIE||' ;
  // Montant de l'échéance
  FListe.ColTypes[2]  := 'R' ;         // Montant
  FListe.ColFormats[2] := '#,#00.00' ;  // Montant
  // Date d'échéance
  FListe.ColTypes[3]   := 'D';
  FListe.ColFormats[3] := ShortDateFormat;
  FListe.ColAligns[3]  := taCenter;

end;

procedure TOF_CPECHEMULTI.InitEvenements;
begin

  // Evt sur la form
  TFVierge(Ecran).OnAfterFormShow := OnAfterFormShow ;
  TFVierge(Ecran).OnCloseQuery    := FormCloseQuery ;
  FOldKeyDown                     := TFVierge(Ecran).OnKeyDown ;
  TFVierge(Ecran).OnKeyDown       := OnKeyDown ;

  // Evt sur liste
  FListe.onCellExit := FListeCellExit ;

  // Evt sur boutons
  BRepart.OnClick := BRepartClick ;
  BSolde.OnClick  := BSoldeClick ;
  BZoom.OnClick   := BZoomClick ;

  // EVT sur mode final
  FModeFinal.OnChange := MODEFINALChanged ;
  
end;

procedure TOF_CPECHEMULTI.InitMsgBox;
begin
  MsgBox := THMsgBox.create( Ecran );
  MsgBox.Mess.Add('0;Répartition des échéances;Erreur de répartition : le montant n''est pas réparti complètement.;W;O;O;O;');
  MsgBox.Mess.Add('1;Répartition des échéances;Le montant est trop élevé pour ce mode de paiement.Voulez-vous passer au mode de remplacement ?;Q;YN;Y;N;');
  MsgBox.Mess.Add('2;Répartition des échéances;Vous ne pouvez pas renseigner des montants négatifs.;W;O;O;O;');
  MsgBox.Mess.Add('3;Répartition des échéances;Les dates d''échéance doivent respecter la plage de saisie autorisée;W;O;O;O;');
end;




procedure TOF_CPECHEMULTI.FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  if FAction = taConsult then Exit;
  if ((FListe.Cells[1, ARow] = '') and (FListe.Cells[2, ARow] = '') and (FListe.Cells[3, ARow] = '')) then exit;

  // MAJ mode paiement => plus nécessaire
//  if ( FListe.Cells[1, ARow] = '' ) and ( FMultiEche.Info.ModePaie.Items.Detail.count > 0 ) then
//    SetModePaie( FMultiEche.Info.ModePaie.Items.Detail[0].GetString('MP_MODEPAIE'), ARow ) ;
  if ( FListe.Cells[1, ARow] = '' ) and ( FListe.ValCombo.Items.Count > 0 ) then
    SetModePaie( FListe.ValCombo.Items[0], ARow ) ;

  // MAJ montant
  if FListe.Cells[2, ARow] = '' then
    FListe.Cells[2, ARow] := StrfMontant(FReste.Value, 15, GetDev.Decimale, '', TRUE);

  // MAJ date
  if FListe.Cells[3, ARow] = '' then
    FListe.Cells[3, ARow] := DateToStr( FMultiEche.TobEcr.GetDateTime('E_DATECOMPTABLE') );

  VerifLimiteModePaie( ARow );

  AfficheTotal;

end;

procedure TOF_CPECHEMULTI.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if FAction <> taConsult
    then CanClose := RecupEche
    else CanClose := True ;

  if CanClose then
    TFVierge(Ecran).FormCloseQuery( Sender, CanClose ) ;
    
end;

function TOF_CPECHEMULTI.GetDev: RDevise;
begin
  result := FMultiEche.Devise ;
end;

procedure TOF_CPECHEMULTI.AfficheTotal;
var  i : Integer ;
     X : Double ;
     Y : Double ;
begin
  FAPayer.Value := FMultiEche.MontantSaisi ;
  X := 0;
  for i := 1 to 12 do
    begin
    Y := Valeur(FListe.Cells[2, i]);
    if Y = 0
      then FListe.Cells[2, i] := ''
      else FListe.Cells[2, i] := StrfMontant(Y, 15, GetDev.Decimale, '', TRUE);
    X := X + Y;
    end;
  FReste.Value := FAPayer.Value - X;
  FSaisi.Value := X;
end;

function TOF_CPECHEMULTI.RecupEche: boolean;
var i          : Integer ;
    lTotal     : Double ;
    lDateEche  : TDateTime ;
    lMontant   : Double ;
    lModePaie  : String ;
begin

  // recalcul du pied de page
  AfficheTotal;

  // Test tout saisi ?
  result := Arrondi( FReste.Value, GetDev.Decimale ) = 0;
  if not Result then
  begin
    MsgBox.Execute(0, '', '');
    exit;
  end;

  // Contrôles de cohérence pour chaque ligne
  for i := 1 to MaxEche do
    if ((FListe.Cells[1, i] <> '') and (FListe.Cells[2, i] <> '') and (FListe.Cells[3, i] <> '')) then
      begin
      // Test Montant Négatifs
      if ((not VH^.MontantNegatif) and (Valeur(FListe.Cells[2, i]) < 0)) then
        begin
        Result := False;
        MsgBox.Execute(2, '', '');
        Exit;
        end;
      // Test Date OK
      lDateEche := StrToDate(FListe.Cells[3, i]);
      if not NbJoursOK(FMultiEche.TobEcr.GetDateTime('E_DATECOMPTABLE'), lDateEche) then
        begin
        Result := False;
        MsgBox.Execute(3, '', '');
        Exit;
        end;

    end;

  // Affectation des valeur à l'objet
  lTotal := 0;
  FMultiEche.ClearEche ;
  for i := 1 to MaxEche do
    if ((FListe.Cells[1, i] <> '') and (FListe.Cells[2, i] <> '') and (FListe.Cells[3, i] <> '')) then
      begin
      lDateEche := StrToDate( FListe.Cells[3, i] ) ;
      lModePaie := FListe.CellValues[1, i] ;
      lMontant  := Valeur(FListe.Cells[2, i]) ;
      lTotal    := lTotal + lMontant ;
      FMultiEche.UpdateEche( i, lModePaie, lMontant, lDateEche ) ;
      end;

  // Arrondi
  FMultiEche.GereArrondi ;

  // Pourcentage ??
  FMultiEche.GerePoucentage ;

  // Mode Final
  if FMultiEche.TobEcr.GetNumChamp('MODEFINAL') > 0
    then FMultiEche.TobEcr.putValue('MODEFINAL', FModeFinal.Value )
    else FMultiEche.TobEcr.addchampsupValeur('MODEFINAL', FModeFinal.Value ) ;

  // Validation des modif sur l'objet de gestion
  if FMultiEche.Piece.ModeEche = meDeporte then
    begin
    FMultiEche.ModeFinal := FModeFinal.Value ;
    if FMultiEche.Action = taCreat then
      FMultiEche.Action := taModif ;
    end ;

end;

procedure TOF_CPECHEMULTI.OnAfterFormShow;
begin
  inherited;
  FBoInit := False ;
end;


procedure TOF_CPECHEMULTI.AnalyseParametre(vStArg: String);
var lStChaine : string;
begin

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt(vStArg);

    // Stocke le XXX de "MODE=XXX" dans la variable
    if (Pos('ACTION=', Uppercase(lStChaine)) > 0) then
      begin
      FAction                    := StringToAction( lStChaine ) ;
      TFVierge(Ecran).TypeAction := FAction ;
      end ;

    end;

end;

procedure TOF_CPECHEMULTI.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;

  case Key of
{AC}     67 : if Shift=[ssAlt]  then
                begin
                Key:=0 ;
                FModeFinal.SetFocus ;
                end ;
    end ;

  if Assigned( FOldKeyDown ) then
    FOldKeyDown( Sender, Key, Shift ) ;

end;

Initialization
  registerclasses ( [ TOF_CPECHEMULTI ] ) ;
end.
