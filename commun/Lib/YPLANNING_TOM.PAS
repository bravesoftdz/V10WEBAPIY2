{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/01/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YPLANNING
Mots clefs ... : TOM;YPLANNING
*****************************************************************}
Unit YPLANNING_TOM ;

Interface

Uses
  Variants,
  Classes,
  sysutils,
  {$IFNDEF EAGLCLIENT}
    db,
    {$IFNDEF DBXPRESS}
      dbtables,
    {$ELSE DBXPRESS}
      uDbxDataSet,
    {$ENDIF DBXPRESS}
  {$ENDIF}
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  uTOMComm
  ;

Type
  TOM_YPLANNING = Class (tTOMComm)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  public
    function GetLastErrorMsg(Const LastError: integer): string;
  private
  protected
    procedure ControlField(Const FieldName: string); override;
    procedure CalculField(Const FieldName: string); override;
    function  RecordIsValid: boolean; override;
  end ;

Const
                           
	TexteMessage: array[1..10] of string = (
	  {1}   'La TobYPL doit être virtuelle.'
    {2}  ,'Le champ YPF_PREFIXE doit exister.'
    {3}  ,'Le champ YPF_GUIDORI doit exister.'
    {4}  ,'Le champ YPF_PREFIXE est incorrect (APL;PCN;).'
    {5}  ,'Le champ YPF_GUIDORI est incorrect.'
    {6}  ,'Le champ YPF_GUIDYRS est incorrect.'
    {7}  ,'Planning intégré'
    {8}  ,'Erreur lors de la lecture des données de type %s.#13#10Traitement annulé.'
    {9}  ,'Le champ APL_RESSOURCE doit exister.#13#10Traitement annulé.'
    {10} ,'Ressource non trouvée dans YRESSOURCE.#13#10Traitement annulé.'
		);

	ErrPrefixe = 4;
  ErrGuidOri = 5;
  ErrGuidYrs = 6;

Implementation

procedure TOM_YPLANNING.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_YPLANNING.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YPLANNING.OnUpdateRecord ;
begin
  Inherited ;

	DisableControl;
  try
  finally
    EnableControl;
  end;

  if RecordIsValid then
  begin
  end;
end;

procedure TOM_YPLANNING.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YPLANNING.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_YPLANNING.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_YPLANNING.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_YPLANNING.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_YPLANNING.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_YPLANNING.CalculField(Const FieldName: string);
begin
  inherited;
end;

procedure TOM_YPLANNING.ControlField(Const FieldName: string);
begin
  DisableControl;
  try
    if FieldName = '...' then
    begin
    end
  finally
    EnableControl;
  end;

  inherited;

  if LastError <> 0 then
    LastErrorMsg := TexteMessage[LastError];
end;

function TOM_YPLANNING.RecordIsValid: boolean;
begin
  Result := false;

  { Contrôle bas niveau }
  if not inherited RecordIsvalid then exit;

  if Assigned(ecran) then
  begin { Contrôle écran }
  end
  else
  begin { Contrôle fTob }
  end;

  { Contrôle écran + fTob }
  if LastError = 0 then
  begin
    if (GetString('YPL_PREFIXE') = '') or (Pos(GetString('YPL_PREFIXE'), 'APL;PCN;RAI;JEV') <= 0) then // $$$ JP 26/04/07: ajout JEV
      LastError := ErrPrefixe
    else if GetString('YPL_GUIDORI') = '' then
      LastError := ErrGuidOri
    else if GetString('YPL_GUIDYRS') = '' then
      LastError := ErrGuidYrs
  end;

  Result := LastError = 0;
  if LastError > 0 then
  begin
    LastErrorMsg := TexteMessage[LastError];
    if Assigned(ecran) then
    begin

    end
    else
    begin
      fTob.AddChampSupValeur('Error', LastErrorMsg, false);
      if V_Pgi.Sav then PgiInfo(LastErrorMsg, 'YPlanning_Tom.RecordIsValid');
    end;
  end;
end;

function TOM_YPLANNING.GetLastErrorMsg(const LastError: integer): string;
begin
  Result := TexteMessage[LastError]; 
end;

Initialization
  registerclasses ( [ TOM_YPLANNING ] ) ;
end.
