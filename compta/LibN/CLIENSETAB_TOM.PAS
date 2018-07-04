{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 20/06/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPLIENSETAB ()
Mots clefs ... : TOF;CPLIENSETAB
*****************************************************************}
Unit CLIENSETAB_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Forms,
     Classes,
     Windows,
     Grids,     // Pour le TGridDrawState
     Graphics,
     {$IFDEF EAGLCLIENT}
       MaineAGL,
       eFichList,
     {$ELSE}
       db,
       dbGrids,
       HDB,
       FE_MAIN,
       FichList,
     {$ENDIF EAGLCLIENT}
     sysutils,
     HCtrls,
     HEnt1,
     uTob,
     hMsgBox,
     hTB97,
     UTOM ;

Procedure CPLanceFiche_LiensEtab( vStRange, vStLequel, vStArgs : string ) ;

Type
  TOM_CLIENSETAB = Class (TOM)
    procedure OnNewRecord                 ; override ;
    procedure OnDeleteRecord              ; override ;
    procedure OnUpdateRecord              ; override ;
    procedure OnLoadRecord                ; override ;
    procedure OnArgument (S : String )    ; override ;
    procedure OnChangeField ( F: TField ) ; override ;
  private
   {$IFDEF EAGLCLIENT}
     FListe : THGrid ;
   {$ELSE}
     FListe : THDBGrid ;
   {$ENDIF EAGLCLIENT}

   {$IFDEF EAGLCLIENT}
      procedure PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
   {$ELSE}
      procedure OnDrawColumnCell (Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
   {$ENDIF EAGLCLIENT}
      procedure GenererLiensAuto ;
      procedure BGenererClick( Sender : TObject ) ;
      function  ValideComptes( var vErr : Integer ; var vMsg : String ) : Boolean ;
  end ;

Implementation

// ===========================================================================
Procedure CPLanceFiche_LiensEtab( vStRange, vStLequel, vStArgs : string ) ;
begin
    AGLLanceFiche('CP', 'CPLIENSETAB', vStRange, vStLequel, vStArgs);
end ;
// ===========================================================================

procedure TOM_CLIENSETAB.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_CLIENSETAB.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CLIENSETAB.OnUpdateRecord ;
var lErr : Integer ;
    lMsg : String ;
begin
  Inherited ;

  if ( GetControlText('CLE_ETABORIG') = '' ) then
    begin
    LASTERROR := 1 ;
    LASTERRORMSG := TraduireMemoire('La saisie de l''établissement d''origine est obligatoire.') ;
    SetFocusControl('CLE_ETABORIG') ;
    Exit ;
    end ;

  if ( GetControlText('CLE_ETABDEST') = '' ) then
    begin
    LASTERROR := 1 ;
    LASTERRORMSG := TraduireMemoire('La saisie de l''établissement de destination est obligatoire.') ;
    SetFocusControl('CLE_ETABDEST') ;
    Exit ;
    end ;

  if GetControlText('CLE_ETABORIG') = GetControlText('CLE_ETABDEST') then
    begin
    LASTERROR := 1 ;
    LASTERRORMSG := TraduireMemoire('L''établissement de destination doit être différent de l''établissement d''origine.') ;
    SetFocusControl('CLE_ETABDEST') ;
    Exit ;
    end ;

  if not ValideComptes( lErr, lMsg ) then
    begin
    LASTERROR    := lErr ;
    LASTERRORMSG := lMsg ;
    Exit ;
    end ;

end ;

procedure TOM_CLIENSETAB.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_CLIENSETAB.OnArgument (S : String ) ;
var lCombo : THValComboBox ;
begin
  Inherited ;

  {$IFDEF EAGLCLIENT}
  FListe := THGrid( GetControl('FLISTE') ) ;
  FListe.PostDrawCell := PostDrawCell ;
  {$ELSE}
  FListe := THDBGrid( GetControl('FLISTE') ) ;
  FListe.OnDrawColumnCell := OnDrawColumnCell ;
  {$ENDIF EAGLCLIENT}

  TToolBarButton97( GetControl( 'BGENERER', True ) ).OnClick := BGenererClick ;

  // PLan Dev Ligne 3230 : Génération auto des liens
  lCombo := THValComboBox( GetControl('CLE_ETABDEST', True) ) ;
  lCombo.Vide := True ;
  lCombo.VideString := TraduireMemoire('<<Tous>>') ;
  lCombo.Reload ;

  lCombo := THValComboBox( GetControl('CLE_ETABORIG', True) ) ;
  lCombo.Vide := True ;
  lCombo.VideString := TraduireMemoire('<<Tous>>') ;
  lCombo.Reload ;
end ;



{$IFDEF EAGLCLIENT}
procedure TOM_CLIENSETAB.PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var Value : String ;
    Text  : String ;
    Rc    : TRect ;
begin
  Value := FListe.Cells[ACol, ARow] ;
  if Value <> '' then
    begin
    Text := RechDom( 'TTETABLISSEMENT', Value, False ) ;
    if Text = '' then Exit ;
    Rc := FListe.CellRect(ACol, ARow) ;
    Canvas.FillRect( Rc );
    Canvas.TextRect( Rc, (Rc.Left+3), (Rc.Top+Rc.Bottom) div 2 -5 , Text );
    end ;
end;
{$ELSE}
procedure TOM_CLIENSETAB.OnDrawColumnCell (Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var Value : String ;
    Text  : String ;
    Rc    : TRect ;
begin
  Value := FListe.Columns[DataCol].Field.AsString ;
  if Value <> '' then
    begin
    Text := RechDom( 'TTETABLISSEMENT', Value, False ) ;
    if Text = '' then Exit ;
    Rc := Rect ;
    FListe.Canvas.FillRect( Rc );
    FListe.Canvas.TextRect( Rc, (Rc.Left+3), (Rc.Top+Rc.Bottom) div 2 -5 , Text );
    end ;
end;
{$ENDIF EAGLCLIENT}


procedure TOM_CLIENSETAB.GenererLiensAuto;
var lTobEtabs : TOB ;
    lTobLiens : TOB ;
    i,j       : Integer ;
begin
  try
    BeginTrans ;

    // suppression des liens actuels
    ExecuteSQL('DELETE FROM CLIENSETAB') ;

    // liste des établissements
    lTobEtabs := TOB.Create('ListeEtab', nil, -1 ) ;
    lTobEtabs.LoadDetailFromSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS') ;

    // Constructions des couples
    For i := 0 to ( lTobEtabs.Detail.Count - 1 ) do
      For j := 0 to ( lTobEtabs.Detail.Count - 1 ) do
        begin

        // Pas de génération d'un établissement vers lui-même...
        if i=j then Continue ;

        // Création de l'objet
        lTobLiens := Tob.Create('CLIENSETAB', nil, -1) ;
        lTobLiens.PutValue('CLE_ETABORIG',   lTobEtabs.Detail[i].GetValue('ET_ETABLISSEMENT') ) ;
        lTobLiens.PutValue('CLE_COMPTEORIG', GetControlText('CLE_COMPTEORIG') ) ;
        lTobLiens.PutValue('CLE_ETABDEST',   lTobEtabs.Detail[j].GetValue('ET_ETABLISSEMENT') ) ;
        lTobLiens.PutValue('CLE_COMPTEDEST', GetControlText('CLE_COMPTEDEST') ) ;

        // Enregistrement
        lTobLiens.InsertDB( nil ) ;

        FreeAndNil( lTobLiens ) ;
        end ;

    CommitTrans ;
    PgiInfo( TraduireMemoire('La génération automatique du paramétrage est terminée.') ) ;

  except
    on E: Exception do
      begin
      PgiError( 'La génération automatique des liens inter-établissements a échoué. Erreur retournée : ' + E.Message, Ecran.Caption ) ;
      rollback ;
      end ;
  end ;

end;

procedure TOM_CLIENSETAB.BGenererClick(Sender: TObject);
var lErr : Integer ;
    lMsg : String ;
begin

  if not ValideComptes( lErr, lMsg ) then
    begin
    PgiError( lMsg, Ecran.Caption ) ;
    Exit ;
    end ;

  if ( GetControlText('CLE_ETABORIG') = '' ) and ( GetControlText('CLE_ETABDEST') = '' ) then
    if PgiAskCancel('Attention, vous allez mettre en place un paramétrage unique d''autorisation de saisie multi établissements.'+ #10#13 + 'Lors ce cette validation , tous les autres paramétrages de liaisons inter établissements vont être détruits. Voulez-vous continuer ?', Ecran.Caption) = mrYes then
      begin
      GenererLiensAuto ;
      TFFicheListe(Ecran).bDefaireClick(nil) ; //CancelRecord ;
      RefreshDB ;
      end ;

end;

procedure TOM_CLIENSETAB.OnChangeField(F: TField);
begin
  inherited;

  // Affichage du bouton de génération automatique
  if ( F.FieldName='CLE_ETABORIG' ) or ( F.FieldName='CLE_ETABDEST' ) then
    SetControlVisible( 'BGENERER',  ( GetField('CLE_ETABORIG') = '' ) and ( GetField('CLE_ETABDEST') = '' ) ) ;

end;

function TOM_CLIENSETAB.ValideComptes(var vErr: Integer; var vMsg: String): Boolean;
begin

  result := False ;

  if ( GetControlText('CLE_COMPTEORIG') = '' ) then
    begin
    vErr := 1 ;
    vMsg := TraduireMemoire('La saisie du compte de liaison d''origine est obligatoire.') ;
    SetFocusControl('CLE_COMPTEORIG') ;
    Exit ;
    end ;

  if ( GetControlText('CLE_COMPTEDEST') = '' ) then
    begin
    vErr := 2 ;
    vMsg := TraduireMemoire('La saisie du compte de liaison de destination est obligatoire.') ;
    SetFocusControl('CLE_COMPTEDEST') ;
    Exit ;
    end ;

  if not ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' + GetControlText('CLE_COMPTEORIG') + '"' ) then
    begin
    vErr := 3 ;
    vMsg := TraduireMemoire('La saisie de la zone "compte de liaison d''origine" ne correspond à aucun compte général valide.') ;
    SetFocusControl('CLE_COMPTEORIG') ;
    Exit ;
    end ;

  if not ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' + GetControlText('CLE_COMPTEDEST') + '"' ) then
    begin
    vErr := 4 ;
    vMsg := TraduireMemoire('La saisie de la zone "compte de liaison de destination" ne correspond à aucun compte général valide.') ;
    SetFocusControl('CLE_COMPTEDEST') ;
    Exit ;
    end ;

  vErr   := 0 ;
  vMsg   := '' ;
  result := True ;

end;

Initialization
  registerclasses ( [ TOM_CLIENSETAB ] ) ;
end.
