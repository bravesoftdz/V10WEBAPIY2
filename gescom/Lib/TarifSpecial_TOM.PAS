{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/11/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TARIFSPECIAL (TARIFSPECIAL)
Mots clefs ... : TOM;TARIFSPECIAL
*****************************************************************}
Unit TarifSpecial_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
     Fiche,
     FichList,
{$ELSE}
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
     wTom;

Type
  TOM_TARIFSPECIAL = Class (twTOM)
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
  	    { Control Field }
   	 procedure ControlField(FieldName: string);
   	 procedure CalculField(FieldName: string);
   	 function  RecordIsValid: boolean; override;
    end ;

Implementation

uses
   wCommuns;

Const
   sFournisseur  = 'FOU'; //Nature Auxiliaire
   sClient       = 'CLI'; //Nature Auxiliaire

	{ Tableau des libellés de messages pour les ressources }
	TexteMessage: array[1..1] of string
          = (
			{1}  'Incohérences dans la définition de la plage de validté : Date de Début > à la Date de Fin'
            );
	ErrDate =  1;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.OnNewRecord ;
begin
  Inherited ;
   if (IKC='C') then
      SetBoolean('YTP_TARIFACTIF',True);

end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.OnDeleteRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.OnUpdateRecord ;
begin
  DisableControl;
  try
  finally
    EnableControl;
  end;

  if RecordIsValid then
  begin
  end;

  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.OnLoadRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.OnClose ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.OnCancelRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
   Data Change Champ par Champ
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.CalculField(FieldName: string);
begin
   if (FieldName='xxx_xxxxxxxx') then ;
end;

{--------------------------------------------------------------------------------
   Validation Champ par Champ
---------------------------------------------------------------------------------}
procedure TOM_TARIFSPECIAL.ControlField(FieldName: string);
begin
	DisableControl;
  try
    if (Pos(FieldName,'xxx_xxxxxxxx')>0) then ;
  finally
    EnableControl;
  end;

//  inherited;

	if LastError <> 0 then
    LastErrorMsg := TexteMessage[LastError];
end;

{--------------------------------------------------------------------------------
   Validation Global
---------------------------------------------------------------------------------}
function TOM_TARIFSPECIAL.RecordIsValid: boolean;
begin
   if (GetDateTime ('YTP_DATEDEBUT')>GetDateTime ('YTP_DATEFIN')) then LastError := ErrDate
   ;
   Result := LastError = 0;

   if LastError > 0 then
   begin
   	LastErrorMsg := TexteMessage[LastError];
      Case LastError of
         ErrDate : SetFocusControl('YTP_DATEDEBUT');
      end;
   end;
end;

Initialization
  registerclasses ( [ TOM_TARIFSPECIAL ] ) ;
end.
