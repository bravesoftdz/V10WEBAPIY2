unit UlibSaisieGrille;

interface

uses
  Hctrls
  ,Classes
  ,Forms
  ,Dialogs
  {$IFDEF EAGLCLIENT}
  {$ELSE}
    ,db
    ,dbTables
  {$ENDIF}
  ,ParamSoc      // pour le GetParamSoc
  ,Controls      // pour le mrYes
  ,paramdbg      // pour le ParamListe
  ,UTob
  ,HEnt1
  ,Ent1
  ,SysUtils
  ,Ed_Tools       // pour VideListe
  ,hmsgbox        // pour MessageAlerte
  ,uLibEcriture
  ;

Type

TSaisieGrille = Class( TObject )
  private

      // Table cible
      FTableNom          : String ;
      FTablePref         : String ;
      FTableCle1         : String ;

      // Données de la grille
      FTobListe          : TOB;               // Contenu de la Grille
      FTobBuffer         : TOB;               // Ligne en cours d'édition
      FTobDeleted        : TOB ;              // lignes supprimées

      // Gestionnaire de données
      FTobCompta         : TobCompta ;        // TOB de validation des données

      // Interface
      FEcran            : TForm;             // interface
      FGrid             : THGrid;            // Grille de saisie
      FGridListe        : String;            // Liste paramètrable

      // Evènements
      FUserBeforeDisplay  : TNotifyEvent ;
      FUserAfterDisplay   : TNotifyEvent ;
      FUserKeyDown        : TKeyEvent;
      FUserRowEnter       : TChangeRC ;
      FUserRowExit        : TChangeRC ;
      FUserCellEnter      : TChangeCell ;
      FUserCellExit       : TChangeCell ;

      // gestion de la recherche dans la liste
      FFindDialog       : TFindDialog;       // Objet de recherche
      FirstFind         : Boolean;

      // Indicateurs
      FEtat             : TDataSetState ;    // Etat des données
      FOldRow           : Integer ;          // Position avant insertion
      FTableWhere: String;          // TOM active

      procedure SetGrille      (const Value: THGrid);
      procedure SetTableWhere  ( vSQL : String ) ;

      function  GetCount :  Integer ;
      function  GetRow   :  Integer ;
      procedure SetRow ( value : Integer ) ;

      procedure InitEvtGrid ;
      procedure InitFormatGrid ;

  //////////////////////////////////////////////////////////////////////////////////////////
  Public

      // ======================
      // ===== PROPRIETES =====
      // ======================

      // Données
      property TOBListe:    TOB              read  FTobListe ;
      property TOBBuffer:   TOB              read  FTobBuffer ;
      property TOBSuppr:    TOB              read  FTobDeleted ;

      // Interface
      property Ecran:       TForm            read  FEcran         write  FEcran ;
      property LaGrille:    THGrid           read  FGrid          write  SetGrille ;
      property LaListe:     String           read  FGridListe     write  FGridListe ;

      // MCD
      property NomTable:    String           read  FTableNom      write  FTableNom ;
      property Prefixe:     String           read  FTablePref     write  FTablePref ;
      property cle1:        String           read  FTableCle1     write  FTableCle1 ;

      // Indicateurs
      property Count:       Integer          read  GetCount ;
      property Row:         Integer          read  GetRow ;
      property Etat:        TDataSetState    read  FEtat          write  FEtat ;
      property TableWhere:  String           read  FTableWhere    write  SetTableWhere ;

      // Evènements
      property OnBeforeDisplay: TNotifyEvent read  FUserBeforeDisplay  write FUserBeforeDisplay;
      property OnAfterDisplay:  TNotifyEvent read  FUserAfterDisplay   write FUserAfterDisplay;
      property OnKeyDown:       TKeyEvent    read  FUserKeyDown        write FUserKeyDown;
      property OnRowEnter:      TChangeRC    read  FUserRowEnter       write FUserRowEnter;
      property OnRowExit:       TChangeRC    read  FUserRowExit        write FUserRowExit;
      property OnCellEnter:     TChangeCell  read  FUserCellEnter      write FUserCellEnter;
      property OnCellExit:      TChangeCell  read  FUserCellExit       write FUserCellExit;

      // ====================
      // ===== METHODES =====
      // ====================

      // Instanciation
      constructor Create( vEcran: TForm ; vGille : THGrid ; vListe : String ; vTable : String ) ;
      Destructor  Destroy ; override ;

      // Actions sur les lignes
      procedure EditRow;
      procedure InsertRow;
      procedure PostRow;
      procedure CancelRow;
      procedure DeleteRow;
      procedure PutLigneToBuffer;
      procedure PutBufferToLigne;

      // Procedure DataSet
      Function  GetSQL      : String ;

      // Modification des Tob
      function  GetTOBCourante : TOB ;
      procedure PutValue( FChamp: String; FValeur: Variant );
      function  GetValue( FChamp: String): Variant;
      procedure PutValueAll( FChamp: String; FValeur: Variant );

      // Appel du paramétrage de liste
      procedure ParamGridListe;
      procedure UpdateGridListe;

      // Raffraichissement des lignes
      procedure ChargeLignesFromTOB ( vTobLignes : TOB ; vBoEcran : Boolean = True ) ;
      procedure ChargeLignesFromSQL ( vBoEcran : Boolean = True ) ;
      procedure RefreshLignes( vRow : Integer = 0 );

      // Recherche dans la grille
      procedure FindDialogFind(Sender: TObject);

      // navigation dans la grille
      procedure RowEnter    ( Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean );
      procedure RowExit     ( Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean );
      procedure CellEnter   ( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
      procedure CellExit    ( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
      procedure KeyDown     ( Sender: TObject; var Key: Word; Shift: TShiftState) ;
//      procedure FormKeyDown ( Sender: TObject; var Key: Word; Shift: TShiftState) ;
  end ;

implementation

{ TSaisieGrille }

procedure TSaisieGrille.CancelRow;
begin

  // En édition de ligne, on annule les modif du buffer -> maj de la grille depuis FTobListe
  if FEtat = dsEdit then
    begin
    PutBufferToLigne ;
    end
  // En insertion, suppression de la nouvelle ligne
  else if FEtat = dsInsert then
    begin
    FTobListe.Detail[ FGrid.Row - 1 ].Free ;
    FGrid.Row := FOldRow ;
    FOldRow   := 0 ;
    end ;

  FEtat := dsBrowse;

end;

procedure TSaisieGrille.DeleteRow;
begin
  if FEtat<>dsBrowse then Exit ;

  if (GetRow > 0) and (GetRow <= FTobListe.Detail.Count) then
    begin
    FTobListe.Detail[ GetRow - 1 ].ChangeParent( FTobDeleted, -1 ) ;
    RefreshLignes ;
    end ;
end;

procedure TSaisieGrille.EditRow;
begin
  if FEtat <> dsEdit then
    begin
    FEtat := dsEdit ;
    PutLigneToBuffer;
    end;
end;

procedure TSaisieGrille.InsertRow;
var lTobNew : TOB ;
    lInIdx  : Integer ;
begin
  if FEtat <> dsBrowse then Exit ;

  // Position initiale
  FOldRow := FGrid.Row ;

  if FOldRow > 0
    then lInIdx := FOldRow - 1
    else lInIdx := - 1 ;

  lTobNew := Tob.Create( FTableNom, FTobListe, lInIdx ) ;

  RefreshLignes ;

  FEtat := dsInsert ;

end;


procedure TSaisieGrille.FindDialogFind(Sender: TObject);
var
  Cancel: Boolean;
begin
  Rechercher(FGrid,FFindDialog, FirstFind);
  RowEnter(FGrid, FGrid.Row, Cancel, True);
end;

function TSaisieGrille.GetValue(FChamp: String): Variant;
var lTobLigne : TOB ;
begin
  Result := '' ;
  lTobLigne := GetTOBCourante ;
  if lTobLigne<>nil then
    Result := lTobLigne.GetValue( FChamp );
end;


procedure TSaisieGrille.ParamGridListe;
begin
{$IFDEF EAGLCLIENT}
  ParamListe (FGridListe,nil,'Personnalisation des listes');
{$ELSE}
  ParamListe (FGridListe,nil,nil,'Personnalisation des listes');
{$ENDIF}
  UpdateGridListe;
end;


procedure TSaisieGrille.PutValue(FChamp: String; FValeur: Variant);
var lTobLigne : TOB ;
begin
  lTobLigne := GetTOBCourante ;
  if lTobLigne<>nil then
    lTobLigne.PutValue( FChamp, FValeur );
end;

function TSaisieGrille.GetRow: Integer;
begin
  Result := FGrid.Row ;
end;

procedure TSaisieGrille.RefreshLignes( vRow : Integer );
begin

  if vRow > 0 then
    begin
    if vRow <= FTOBListe.Detail.count then
      begin
      FTOBListe.Detail[ vRow-1 ].PutLigneGrid( FGrid, vRow, FALSE, FALSE, FGridListe );
      end ;
    end
  else
    begin
    FGrid.RowCount:=FGrid.FixedRows+1 ;
    FTOBListe.PutGridDetail( FGrid, false, false, FGridListe, false );
    end ;

end;

procedure TSaisieGrille.SetGrille(const Value: THGrid);
begin
   FGrid := Value;
   FFindDialog := TFindDialog.Create( Ecran );
   FFindDialog.name := 'FindDialog';
   FFindDialog.OnFind := FindDialogFind;
end;

procedure TSaisieGrille.UpdateGridListe;
Var FRecordSource : String ;
    FLien         : String ;
    FSortGrid     : String ;
    F,T,L,A       : String ;
    FParams       : String ;
    tt            : string ;
    NC            : string ;
    Perso         : string ;
    OkTri         : boolean;
    OkNumCol      : boolean;

begin
  // ??
{ if FGrid.ValCombo<>nil then
    FGrid.ValCombo.Free;
}
  // ??
  if ChargeHListe( FGridListe,FRecordSource,FLien,FSortGrid,F,T,L,A,FParams,tt,NC,Perso,OkTri,OkNumCol) then
    begin
    // Test
    OkTri := False ;
    end ;

  OkTri := True ;

end;

constructor TSaisieGrille.Create( vEcran: TForm ; vGille : THGrid ; vListe : String ; vTable : String );
begin

  FTableNom      := vTable;
  FTableCle1     := TableToCle1( vTable ) ;
  FTablePref     := TableToPrefixe( vTable ) ;

  FEcran       := vEcran;
  FGrid        := vGille;
  FGridListe   := vListe;

  FTobListe   := Tob.Create( 'FTobListe',     nil,   -1) ;
  FTobBuffer  := Tob.Create( FTableNom,       nil,   -1) ;
  FTobDeleted := Tob.Create( 'FTobDeleted',   nil,   -1) ;

  // Réaffectation des evts de la grille
  InitEvtGrid ;
  InitFormatGrid ;  

  // Init des variables
  FEtat   := dsInactive ;

end;

destructor TSaisieGrille.Destroy;
begin

  // TOBs
  FreeAndNil( FTobListe ) ;
  FreeAndNil( FTobDeleted ) ;
  FreeAndNil( FTobBuffer ) ;

  // Composants
  FreeAndNil( FFindDialog ) ;

  inherited Destroy;

end;

procedure TSaisieGrille.PostRow;
begin
  if FEtat=dsBrowse then exit;

  // Recopie des données du buffer dans la liste
  FTobBuffer.Dupliquer( FTobListe.Detail[ GetRow - 1 ], True, True, True ) ;
  FEtat := dsBrowse ;
end;

procedure TSaisieGrille.PutLigneToBuffer;
begin
  if FEtat <> dsEdit then Exit ;
  FTobBuffer.Dupliquer( FTobListe.Detail[ GetRow - 1 ], True, True, True ) ;
end;

procedure TSaisieGrille.PutBufferToLigne;
begin
  if FEtat <> dsEdit then Exit ;
  FTobListe.Detail[ GetRow - 1 ].Dupliquer( FTobBuffer, True, True, True ) ;
end;

function TSaisieGrille.GetCount: Integer;
begin
  Result := FTobListe.Detail.count ;
end;

procedure TSaisieGrille.SetTableWhere(vSQL: String);
begin
  FTableWhere := vSQL ;
end;


function TSaisieGrille.GetSQL: String;
begin
  if FTableNom = '' then Exit ;
  Result := 'SELECT * FROM ' + FTableNom + FTableWhere ;
end;

function TSaisieGrille.GetTOBCourante: TOB;
begin
  Result := nil ;
  if ( GetRow > 0 ) and ( GetRow <= FTobListe.Detail.Count) then
    Result := FTobListe.Detail[ GetRow - 1 ] ;
end;

procedure TSaisieGrille.ChargeLignesFromSQL(vBoEcran: Boolean);
begin

  if FEtat<>dsBrowse then Exit ;

  // Libération des Tob actuelles
  FTobListe.ClearDetail ;
  FTobDeleted.ClearDetail ;

  FTobListe.LoadDetailDBFromSQL( FTableNom, GetSQL, False ) ;

  // Evènement optionnel pré affichage
  if Assigned( FUserBeforeDisplay ) then
    FUserBeforeDisplay( FGrid ) ;

  // Affichage
  RefreshLignes ;

  // Evènement optionnel post affichage
  if Assigned( FUserAfterDisplay ) then
    FUserAfterDisplay( FGrid ) ;

end;

procedure TSaisieGrille.ChargeLignesFromTOB( vTobLignes: TOB; vBoEcran: Boolean);
begin
  if FEtat<>dsBrowse then Exit ;

  // Libération des Tob actuelles
  FTobListe.ClearDetail ;
  FreeAndNil( FTobListe ) ;
  FTobDeleted.ClearDetail ;

  FTobListe := vTobLignes ;

  // Evènement optionnel pré affichage
  if Assigned( FUserBeforeDisplay ) then
    FUserBeforeDisplay( FGrid ) ;

  // Affichage
  RefreshLignes ;

  // Evènement optionnel post affichage
  if Assigned( FUserAfterDisplay ) then
    FUserAfterDisplay( FGrid ) ;

end;

procedure TSaisieGrille.PutValueAll(FChamp: String; FValeur: Variant);
var i : Integer ;
begin
  for i := 0 to GetCount - 1 do
    TobListe.PutValue( FChamp, FValeur );
end;

procedure TSaisieGrille.CellEnter( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean );
begin
  if Assigned( FUserCellEnter ) then
    FUserCellEnter( Sender, Acol, Arow, Cancel );
end;

procedure TSaisieGrille.CellExit( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean );
begin
  if Assigned( FUserCellExit ) then
    FUserCellExit( Sender, Acol, Arow, Cancel );
end;

procedure TSaisieGrille.KeyDown( Sender: TObject ; var Key: Word ; Shift: TShiftState ) ;
begin
  if Assigned( FUserKeyDown ) then
    FUserKeyDown( Sender, Key, Shift );
end;

procedure TSaisieGrille.RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if Assigned( FUserRowEnter ) then
    FUserRowEnter( Sender, Ou, Cancel, Chg );
end;

procedure TSaisieGrille.RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if Assigned( FUserRowExit ) then
    FUserRowExit( Sender, Ou, Cancel, Chg );
end;

procedure TSaisieGrille.SetRow(value: Integer);
begin

end;

procedure TSaisieGrille.InitEvtGrid;
begin
  // Réaffectation des évènements
  FGrid.OnRowEnter   := RowEnter ;
  FGrid.OnRowExit    := RowExit ;
  FGrid.OnCellEnter  := CellEnter ;
  FGrid.OnCellExit   := CellExit ;
  FGrid.OnKeyDown    := KeyDown ;
end;

procedure TSaisieGrille.InitFormatGrid;
begin
  FGrid.VidePile(True) ;
  FGrid.RowCount := 2 ;
  FGrid.ListeParam:= FGridListe ;

  UpdateGridListe;
end;

end.
