{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YTARIFSFONCTIONNALITE (YTARIFSFONCTIONNALITE)
Mots clefs ... : TOM;YTARIFSFONCTIONNALITE
*****************************************************************}
Unit TarifsFonctionnalite_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     UTOM,
     UTob,
     wtom
     ;

Type
  Tom_yTarifsFonctionnalite = Class (twTOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  protected
    procedure ControlField(FieldName: string); override;
    procedure CalculField(FieldName: string); override;
    function  RecordIsValid: boolean; override;
  end ;

Implementation

Const
	TexteMessage: array[1..1] of string
          = (
	      		{ 1}  'Au cas où ...'
            );

	Erreur =  1;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.OnNewRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.OnDeleteRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.OnUpdateRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.OnLoadRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.OnArgument ( S: String ) ;
begin
  Inherited ;
  tWinControl(GetControl('YFO_OKAFFAIRE')).visible     := (ctxAFFAIRE in V_PGI.PGIContexte);
  tWinControl(GetControl('YFO_OKPOLESITEGRP')).visible := (ctxGPAO    in V_PGI.PGIContexte);
  tWinControl(GetControl('YFO_OKJHL')).visible         := (ctxCHR     in V_PGI.PGIContexte);
  tWinControl(GetControl('YFO_OKYTALIGNE')).visible    := (ctxCHR     in V_PGI.PGIContexte) or (ctxAFFAIRE in V_PGI.PGIContexte);
  tWinControl(GetControl('YFO_OKCOMMERCIAL')).visible  := False;
  tWinControl(GetControl('YFO_OKCPT')).visible         := (ctxGPAO    in V_PGI.PGIContexte);
  tWinControl(GetControl('YFO_OKRESSOURCE')).visible   := (ctxCHR     in V_PGI.PGIContexte) or (ctxAFFAIRE in V_PGI.PGIContexte);
  tWinControl(GetControl('YFO_OKYTQLIGNE')).visible    := (ctxCHR     in V_PGI.PGIContexte) or (ctxAFFAIRE in V_PGI.PGIContexte);
  tWinControl(GetControl('YFO_OKMARQUE')).visible      := False;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.OnClose ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.OnCancelRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.CalculField(FieldName: string);
begin
  inherited;
  if (FieldName='YFO_FONCTIONNALITE') then ;
end;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure Tom_yTarifsFonctionnalite.ControlField(FieldName: string);
begin
	DisableControl;
  try
    ;
  finally
    EnableControl;
  end;

  inherited;

	if LastError <> 0 then
 		LastErrorMsg := TexteMessage[LastError];
end;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
function Tom_yTarifsFonctionnalite.RecordIsValid: boolean;
begin
  if      (False) then LastError := Erreur
  ;

  Result := (LastError = 0);

  if LastError > 0 then
  begin
  	LastErrorMsg := TexteMessage[LastError];
    Case LastError of
      Erreur : SetFocusControl('YFO_FONCTIONNALITE');
    end;
  end;
end;

Initialization
  registerclasses ( [ TOM_YTARIFSFONCTIONNALITE ] ) ;
end.



