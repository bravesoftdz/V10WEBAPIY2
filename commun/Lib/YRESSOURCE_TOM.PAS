{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/01/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YRESSOURCE
Mots clefs ... : TOM;YRESSOURCE
*****************************************************************}
Unit YRESSOURCE_TOM ;

Interface

Uses
  Variants,
  Classes,
  //sysutils,
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
  TOM_YRESSOURCE = Class (tTOMComm)
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

  //20080828_C.B_GA15215_debut
	TexteMessage: array[1..5] of string = (
	  {1}  'Il y a déjà dans la base une ressource de type Salarié pour l''utilisateur %s. La ressource %s n''a pas pu être créée dans la table YRESSOURCE.',
    {2}  'Erreur lors de la création de la ressource %s dans la table YRESSOURCE.#13#10Traitement annulé.',
    {3}  'Planning intégré',
    {4}  'Erreur lors de la création du salarié %s dans la table YRESSOURCE.#13#10Traitement annulé.',
    {5}  'Erreur lors de la création de l''utilisateur %s dans la table YRESSOURCE.#13#10Traitement annulé.'
		); 
  //20080828_C.B_GA15215_fin 

	ErrXXX = 1;

Implementation

procedure TOM_YRESSOURCE.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_YRESSOURCE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YRESSOURCE.OnUpdateRecord ;
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

procedure TOM_YRESSOURCE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YRESSOURCE.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_YRESSOURCE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_YRESSOURCE.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_YRESSOURCE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_YRESSOURCE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_YRESSOURCE.CalculField(Const FieldName: string);
begin
  inherited;
end;

procedure TOM_YRESSOURCE.ControlField(Const FieldName: string);
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

function TOM_YRESSOURCE.RecordIsValid: boolean;
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
      if V_Pgi.Sav then PgiInfo(LastErrorMsg, 'YRESSOURCE_Tom.RecordIsValid');
    end;
  end;
end;

function TOM_YRESSOURCE.GetLastErrorMsg(const LastError: integer): string;
begin
  Result := TexteMessage[LastError]; 
end;

Initialization
  registerclasses ( [ TOM_YRESSOURCE ] ) ;
end.
