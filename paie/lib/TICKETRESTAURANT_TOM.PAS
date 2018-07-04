{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TICKETRESTAURANT (TICKETRESTAURANT)
Mots clefs ... : TOM;TICKETRESTAURANT
*****************************************************************}
Unit TICKETRESTAURANT_TOM ;

Interface

Uses
//unused     StdCtrls,
//unused     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
//unused     Fiche,
//unused     FichList,
{$ELSE}
//unused     eFiche,
//unused     eFichList,
{$ENDIF}
     dialogs,
//unused     forms,
//unused     sysutils,
//unused     ComCtrls,
//unused     HTB97,
//unused     HCtrls,
     HEnt1,
//unused     HMsgBox,
     PgOutils,
     UTOM,
     uPaieRemunerations,
     UTob ;

Type
  TOM_TICKETRESTAURANT = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation
var
  TOB_Rem       : TOB;
procedure TOM_TICKETRESTAURANT.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_TICKETRESTAURANT.OnDeleteRecord ;
var
   Champ :array[1..1] of Hstring;
   Valeur :array[1..1] of variant ;

   ExisteCod : Boolean;

begin
  Inherited ;
//  ExisteCod := False ;

  Champ[1] := 'PRT_TYPTICKET';
  Valeur[1] := GetField('PTT_TYPTICKET');
  ExisteCod := RechEnrAssocier('CDETICKETS',Champ,Valeur);
  if ExisteCod = TRUE then
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! Ce type de titre été utilisé pour des commandes.'+
                  '#13#10Vous ne pouvez le supprimer!';
  end;

  Champ[1] := 'ETB_TYPTICKET';
  Valeur[1] := GetField('PTT_TYPTICKET');
  ExisteCod := RechEnrAssocier('ETABCOMPL',Champ,Valeur);
  if ExisteCod = TRUE then
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! Ce type de titre est utilisé par, au moins, '+
                    'un établissement.'+
                    '#13#10Vous ne pouvez le supprimer!';
  end;
  Champ[1] := 'PSE_TYPTICKET';
  Valeur[1] := GetField('PTT_TYPTICKET');
  ExisteCod := RechEnrAssocier('DEPORTSAL',Champ,Valeur);
  if ExisteCod = TRUE then
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! Ce type de titre est utilisé pour, au moins, '+
                    'un salarié.'+
                    '#13#10Vous ne pouvez le supprimer!';
  end;
end ;

procedure TOM_TICKETRESTAURANT.OnUpdateRecord ;
var
  RubRem                        : string;
begin
  Inherited ;
  RubRem := GetField('PTT_RETSALARIAL');
  if (TOB_Rem.FindFirst (['PRM_RUBRIQUE'], [RubRem], TRUE) = NIL) and
     (RubRem <> '') then
  begin
    ShowMessage ('La rubrique n''existe pas');
    SetFocusControl('PTT_RETSALARIAL');
  end;

  RubRem := GetField('PTT_RETPATRONALE');
  if (TOB_Rem.FindFirst (['PRM_RUBRIQUE'], [RubRem], TRUE) = NIL) and
     (RubRem <> '') then
  begin
    ShowMessage ('La rubrique n''existe pas');
    SetFocusControl('PTT_RETPATRONALE');
  end;
  
  if (GetField('PTT_NBTICKET') = 0 ) then
    ShowMessage ('La quantité est à zéro');
  if (GetField('PTT_VFACIALE') = 0) then
    ShowMessage ('La valeur faciale est à zéro');
  if (GetField('PTT_PPATRONALE') = 0) then
    ShowMessage ('La part employeur est à zéro');
end ;

procedure TOM_TICKETRESTAURANT.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_TICKETRESTAURANT.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_TICKETRESTAURANT.OnChangeField ( F: TField ) ;
var
  RubRem                        : string;
begin
  Inherited ;
  if (F.FieldName = 'PTT_RETSALARIAL') then
    RubRem := GetField('PTT_RETSALARIAL');
  if (F.FieldName = 'PTT_RETPATRONALE') then
    RubRem := GetField('PTT_RETPATRONALE');
  if (TOB_Rem.FindFirst (['PRM_RUBRIQUE'], [RubRem], TRUE) = NIL) and
     (RubRem <> '') then
  begin
    ShowMessage ('La rubrique n''existe pas');
    if (F.FieldName = 'PTT_RETSALARIAL') then
      SetFocusControl('PTT_RETSALARIAL');
    if (F.FieldName = 'PTT_RETPATRONALE') then
      SetFocusControl('PTT_RETPATRONALE');
  end;
end ;

procedure TOM_TICKETRESTAURANT.OnArgument ( S: String ) ;
begin
  Inherited ;
  TOB_Rem := InitTOB_Rem();
end ;

procedure TOM_TICKETRESTAURANT.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_TICKETRESTAURANT.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_TICKETRESTAURANT ] ) ; 
end.
