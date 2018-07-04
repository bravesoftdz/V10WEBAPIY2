{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 20/06/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPLIENSETAB ()
Mots clefs ... : TOF;CPLIENSETAB
*****************************************************************}
Unit CLIENSSOC_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Forms,
     Classes,
{$IFDEF VER150}
     variants,
{$ENDIF}
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
     utilPGI,
     ParamSoc,
     UTOM ;

Procedure CPLanceFiche_LiensSociete( vStRange, vStLequel, vStArgs : string ) ;

Type
  TOM_CLIENSSOC = Class (TOM)
    procedure OnNewRecord                 ; override ;
    procedure OnUpdateRecord              ; override ;
    procedure OnLoadRecord                ; override ;
    procedure OnArgument (S : String )    ; override ;
    procedure OnChangeField ( F: TField ) ; override ;
    procedure OnDeleteRecord              ; override ;
    procedure OnAfterUpdateRecord         ; override ;
  private

   {$IFDEF EAGLCLIENT}
     FListe : THGrid ;
   {$ELSE}
     FListe : THDBGrid ;
   {$ENDIF EAGLCLIENT}

      FTobSoc   : Tob ;
      Societe   : THValComboBox ;

      procedure GenererLiensAuto ;
      procedure BGenererClick( Sender : TObject ) ;
      function  ValideComptes( var vErr : Integer ; var vMsg : String ) : Boolean ;
      procedure SocieteChanged( Sender : TObject ) ;
      function  GetDataBaseName : String ;

   {$IFDEF EAGLCLIENT}
      procedure PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
   {$ELSE}
      procedure OnDrawColumnCell (Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
   {$ENDIF EAGLCLIENT}
    procedure BRibClick(Sender : TObject);
  end ;

Implementation

uses
  UProcGen, {StrRight}
  BanqueCP_TOM, Constantes, Commun,
  Ent1;

// ===========================================================================
Procedure CPLanceFiche_LiensSociete( vStRange, vStLequel, vStArgs : string ) ;
begin
    AGLLanceFiche('CP', 'CPLIENSSOC', vStRange, vStLequel, vStArgs);
end ;
// ===========================================================================

procedure TOM_CLIENSSOC.OnNewRecord ;
begin
  Inherited ;

  SetField('CLS_SOCIETE', '' ) ;
  SetField('CLS_GENERAL', '' ) ;
  SetField('CLS_DOSSIER', '' ) ;

end ;

procedure TOM_CLIENSSOC.OnUpdateRecord ;
var lErr : Integer ;
    lMsg : String ;
    lSoc : string ;
    RacG : string; {JP 28/11/07}
begin
  Inherited ;

  NextPrevControl( Ecran ) ;

  lSoc := GetControlText('SOCIETE') ;

  // Validation de la société
  if ( lSoc = '' ) then
    begin
    LASTERROR := 1 ;
    LASTERRORMSG := TraduireMemoire('La saisie de la société de destination est obligatoire.') ;
    SetFocusControl('SOCIETE') ;
    Exit ;
    end ;

  // Validation du compte
  if not ValideComptes( lErr, lMsg ) then
    begin
    LASTERROR    := lErr ;
    LASTERRORMSG := lMsg ;
    Exit ;
    end ;

  // MAJ des champs inivisibles
  SetField('CLS_SOCIETE', lSoc ) ;
  SetField('CLS_DOSSIER', GetDataBaseName ) ;

  {JP 28/11/07 : Comme en synchronisation, on ne gère que les 517 et les 451, on avertit l'utiliateur}
  if GetParamSocSecur('SO_TRBASETRESO', '') = '' then Exit;
  RacG := Copy(VarToStr(GetField('CLS_GENERAL')), 1, 3);
  if (RacG <> '517') and (RacG <> '451') then begin
    if PgiAsk(TraduireMemoire('Seuls les comptes 451 et 517 sont traités lors de la synchronisation.') + #13 +
              TraduireMemoire('En choisissant le compte "') + VarToStr(GetField('CLS_GENERAL')) + '", ' +
              TraduireMemoire('vous vous ne pourrez pas synchroniser les') + #13 +
              TraduireMemoire('écritures de ce compte en Trésorerie.') + #13#13 +
              TraduireMemoire('Souhaitez-vous poursuivre ?'), Ecran.Caption) = mrNo then LastError := -1;

  end;
end;

procedure TOM_CLIENSSOC.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_CLIENSSOC.OnArgument (S : String ) ;
var lTob    : TOB ;
    i       : Integer ;
begin
  Inherited ;

  {$IFDEF EAGLCLIENT}
  FListe := THGrid( GetControl('FLISTE') ) ;
  FListe.PostDrawCell := PostDrawCell ;
  {$ELSE}
  FListe := THDBGrid( GetControl('FLISTE') ) ;
  FListe.OnDrawColumnCell := OnDrawColumnCell ;
  {$ENDIF EAGLCLIENT}

  // Société de départ fixe
  SetControlText('SOCIETEORIG', GetParamSocSecur('SO_LIBELLE', '') ) ;
//  SetControlText('DOSSIERORIG', V_PGI.SchemaName ) ;
    SetControlText('DOSSIERORIG', GetParamSocSecur('SO_SOCIETE', '') ) ; // FQ 17962

  // Récupération société
  FTobSoc := RecupInfosSocietes ( 'SO_SOCIETE;SO_LIBELLE' ) ;

  // MAJ de la liste des société
  Societe := THValComboBox( GetControl('SOCIETE', True ) ) ;
  Societe.Items.Add ( TraduireMemoire('<<Tous>>') )  ;
  Societe.Values.Add( '' )  ;
  for i := 0 to FTobSoc.Detail.count - 1 do
    begin
    lTob := FTobSoc.Detail[i] ;
    if lTob.GetString('SCHEMANAME') = V_PGI.SchemaName then
      Continue ;
    Societe.Values.Add( lTob.GetString('SO_SOCIETE') )  ;
    Societe.Items.Add(  lTob.GetString('SO_LIBELLE') )  ;
    end ;

  // Evt bouton générer
  TToolBarButton97( GetControl( 'BGENERER', True ) ).OnClick := BGenererClick ;
  Societe.OnClick  := SocieteChanged ;

  {JP 31/07/06 : Ajout de la notion de compte courant dans la table BANQUECP et donc
   rattachement du compte général à un "simili" RIB}
  SetControlEnabled('BRIB', GetParamSocSecur('SO_TRBASETRESO', '') <> '');
  TToolBarButton97(GetControl('BRIB')).OnClick := BRIBClick;
end ;


procedure TOM_CLIENSSOC.GenererLiensAuto;
var lTobLiens : Tob ;
    lTob      : Tob ;
    i         : Integer ;
begin

  try
    BeginTrans ;

    // suppression des liens actuels
    ExecuteSQL('DELETE FROM CLIENSSOC') ;

    // Constructions des couples
    For i := 0 to ( FTobSoc.Detail.Count - 1 ) do
      begin

      lTob := FTobSoc.Detail[i] ;
      if lTob.GetString('SCHEMANAME') = V_PGI.SchemaName then
        Continue ;

      // Création de l'objet
      lTobLiens := Tob.Create('CLIENSSOC', nil, -1) ;
      lTobLiens.PutValue('CLS_SOCIETE', lTob.GetString('SO_SOCIETE') ) ;
      lTobLiens.PutValue('CLS_GENERAL', GetControlText('CLS_GENERAL') ) ;
      lTobLiens.PutValue('CLS_DOSSIER', lTob.GetString('SCHEMANAME') ) ;

      // Enregistrement
      lTobLiens.InsertDB( nil ) ;

      FreeAndNil( lTobLiens ) ;
      end ;

    CommitTrans ;
    PgiInfo( TraduireMemoire('La génération automatique du paramétrage est terminée.') ) ;

  except
    on E: Exception do
      begin
      PgiError( 'La génération automatique des liaisons inter-sociétés a échoué. Erreur retournée : ' + E.Message, Ecran.Caption ) ;
      rollback ;
      end ;
  end ;

end;

procedure TOM_CLIENSSOC.BGenererClick(Sender: TObject);
var lErr : Integer ;
    lMsg : String ;
begin

  NextPrevControl( Ecran ) ;

  if not ValideComptes( lErr, lMsg ) then
    begin
    PgiError( lMsg, Ecran.Caption ) ;
    Exit ;
    end ;

  if Societe.ItemIndex = 0 then
    if PgiAskCancel('Attention, vous allez mettre en place un paramétrage unique de liaisons inter-sociétés.'+ #10#13 + 'Lors ce cette validation , tous les autres paramétrages de liaisons inter-sociétés vont être détruits. Voulez-vous continuer ?', Ecran.Caption) = mrYes then
      begin
      GenererLiensAuto ;
      TFFicheListe(Ecran).bDefaireClick(nil) ; //CancelRecord ;
      RefreshDB ;
      end ;

end;

procedure TOM_CLIENSSOC.OnChangeField(F: TField);
begin
  inherited;

  if F.FieldName = 'CLS_SOCIETE' then
    begin
    if GetField('CLS_SOCIETE')=''
      then Societe.ItemIndex := 0
      else SetControlText('SOCIETE', GetField('CLS_SOCIETE') ) ;
    SocieteChanged( nil ) ;
    end ;

end;

function TOM_CLIENSSOC.ValideComptes(var vErr: Integer; var vMsg: String): Boolean;
begin

  result := False ;

  if ( GetControlText('CLS_GENERAL') = '' ) then
    begin
    vErr := 1 ;
    vMsg := TraduireMemoire('La saisie du compte de liaison est obligatoire.') ;
    SetFocusControl('CLS_GENERAL') ;
    Exit ;
    end ;

  if not ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' + GetControlText('CLS_GENERAL') + '"' ) then
    begin
    vErr := 2 ;
    vMsg := TraduireMemoire('La saisie de la zone "compte de liaison" ne correspond à aucun compte général valide.') ;
    SetFocusControl('CLS_GENERAL') ;
    Exit ;
    end ;

  vErr   := 0 ;
  vMsg   := '' ;
  result := True ;

end;

procedure TOM_CLIENSSOC.SocieteChanged(Sender: TObject);
begin
//  SetControlText( 'DOSSIER', GetDataBaseName ) ;
  SetControlText('DOSSIER', Societe.Value ) ;  // FQ 17962
  SetControlVisible( 'BGENERER',  Societe.ItemIndex=0 ) ;
end;

function TOM_CLIENSSOC.GetDataBaseName: String;
var lTob : TOB ;
    lSoc : string ;
begin
  result := '' ;
  lSoc   := GetControlText('SOCIETE') ;
  if lSoc = '' then Exit ;

  lTob := FTobSoc.FindFirst( ['SO_SOCIETE'], [ lSoc ], True ) ;
  if lTob<>nil then
    result := lTob.GetString('SCHEMANAME') ;
end;

{$IFDEF EAGLCLIENT}
procedure TOM_CLIENSSOC.PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var Value : String ;
    Text  : String ;
    Rc    : TRect ;
    lTob  : Tob ;
begin
  if ACol<>0 then Exit ;
  Value := FListe.Cells[ACol, ARow] ;
  if Value <> '' then
    begin
    lTob := FTobSoc.FindFirst( ['SO_SOCIETE'], [ Value ], True ) ;
    if lTob<>nil then
      Text := lTob.GetString('SO_LIBELLE') ;
    if Text = '' then Exit ;
    Rc := FListe.CellRect(ACol, ARow) ;
    Canvas.FillRect( Rc );
    Canvas.TextRect( Rc, (Rc.Left+3), (Rc.Top+Rc.Bottom) div 2 -5 , Text );
    end ;
end;
{$ELSE}
procedure TOM_CLIENSSOC.OnDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;  State: TGridDrawState);
var Value : String ;
    Text  : String ;
    Rc    : TRect ;
    lTob  : Tob ;
begin
  if DataCol<>0 then Exit ;
  Value := FListe.Columns[DataCol].Field.AsString ;
  if Value <> '' then
    begin
    lTob := FTobSoc.FindFirst( ['SO_SOCIETE'], [ Value ], True ) ;
    if lTob<>nil then
      Text := lTob.GetString('SO_LIBELLE') ;
    if Text = '' then Exit ;
    Rc := Rect ;
    FListe.Canvas.FillRect( Rc );
    FListe.Canvas.TextRect( Rc, (Rc.Left+3), (Rc.Top+Rc.Bottom) div 2 -5 , Text );
    end ;
end;
{$ENDIF EAGLCLIENT}

{---------------------------------------------------------------------------------------}
procedure TOM_CLIENSSOC.BRibClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Ok  : Boolean;
  Lib : string;
  Gen : string;
begin
  if GetParamSocSecur('SO_TRBASETRESO', '') = '' then Exit;

  Gen := GetField('CLS_GENERAL');
  if Gen = '' then Gen := GetControlText('CLS_GENERAL');
  {On regarde si le rib existe ...}
  Ok := ExisteSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_GENERAL ="' + Gen +
                  '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"');
  {... sinon, on le crée}
  if not Ok then begin
    NextPrevControl(Ecran);
    Lib := GetInfosFromDossier('DOS_SOCIETE', GetParamSocSecur('SO_SOCIETE', ''), 'DOS_NOMBASE');
    Lib := Lib + ' vers ' + GetInfosFromDossier('DOS_SOCIETE', GetControlText('SOCIETE'), 'DOS_NOMBASE');
    ExecuteSql('INSERT INTO BANQUECP (BQ_CODE, BQ_GENERAL, BQ_BANQUE, BQ_AGENCE, BQ_LIBELLE, BQ_DEVISE, BQ_PAYS, ' +
               'BQ_SOCIETE, BQ_NODOSSIER, BQ_NATURECPTE, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, '+
               'BQ_CODECIB, BQ_DOMICILIATION) ' +
               'VALUES ("' + Copy(Gen, 1, 12) + StrRight(V_PGI.NoDossier, 5) + '", "' +
                             Gen + '", "' + GetBanqueCourant + '", "' + GetAgenceCourant +
                             '", "' + Lib + '", "' + V_PGI.DevisePivot + '", "' +
                             CodePaysDeIso(GetParamSocSecur('SO_PAYSLOCALISATION','')) + '", "' +
                             V_PGI.CodeSociete + '", "' + V_PGI.NoDossier + '", "' + tcb_Courant + '", "' +
                             CODEATTENTE5 + '", "' + CODEATTENTE5 + '", "' + CODEATTENTE5 + '", "53", "' +
                             GetCibCourant + '", "' + TraduireMemoire('Agence comptes courants') + '")');
    //HShowMessage('0;' + Ecran.Caption + ';Un compte courant a été créé et rattaché par défaut à ce compte.;I;O;O;O;', '', '');
  end;
//  Lib := GetInfosFromDossier('DOS_SOCIETE', GetParamSocSecur('SO_SOCIETE', ''), 'DOS_NODOSSIER');
  FicheBanqueCP(Gen, taModif, 0, V_PGI.NoDossier);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLIENSSOC.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if GetParamSocSecur('SO_TRBASETRESO', '') = '' then Exit;
  if not ExisteSQL('SELECT * FROM BANQUECP WHERE BQ_GENERAL = "' + GetField('CLS_GENERAL') +
                   '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"') then BRibClick(GetControl('BRIB'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLIENSSOC.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if GetParamSocSecur('SO_TRBASETRESO', '') = '' then Exit;
  ExecuteSql('DELETE FROM BANQUECP WHERE BQ_GENERAL = "' + GetField('CLS_GENERAL') + '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"');
end;


initialization
  registerclasses ( [ TOM_CLIENSSOC ] ) ;
end.
