{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/08/2017
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BSVPLANCLA (BSVPLANCLA)
Mots clefs ... : TOM;BSVPLANCLA
*****************************************************************}
Unit BSVPLANCLA_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
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
     UTob ;

Type
  TOM_BSVPLANCLA = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure OnAfterCommit              ; override ;
    procedure OnAfterRollBack            ; override ;
    procedure OnAfterBeginTrans          ; override ;
    private
      function LastNIV : Integer;
    end ;

Implementation

procedure TOM_BSVPLANCLA.OnNewRecord ;
begin
  Inherited ;
  SetField('BP0_NIV',LastNIV+1);
  SetField('BP0_LIBELLE','');
  SetField('BP0_TYPECLA','001');
  SetField('BP0_FIXE','');
  SetField('BP0_CHPSCLA','');
  SetField('BP0_CHPSLIB','-');
  SetControlVisible('BP0_FIXE',false);
  SetControlVisible('TBP0_FIXE',false);
  SetControlVisible('BP0_CHPSCLA',true);
  SetControlVisible('TBP0_CHPSCLA',true);
  SetControlVisible('BP0_CHPSLIB',true);
end ;

procedure TOM_BSVPLANCLA.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnLoadRecord ;
begin
  Inherited ;
  if GetField('BP0_TYPECLA') = '000' then
  begin
    SetControlVisible('BP0_FIXE',true);;
    SetControlVisible('BP0_CHPSCLA',false);;
    SetControlVisible('BP0_CHPSLIB',false);;
    SetControlVisible('TBP0_FIXE',true);;
    SetControlVisible('TBP0_CHPSCLA',false);;
  end else
  begin
    SetControlVisible('BP0_FIXE',false);;
    SetControlVisible('BP0_CHPSCLA',true);;
    SetControlVisible('BP0_CHPSLIB',true);;
    SetControlVisible('TBP0_FIXE',false);;
    SetControlVisible('TBP0_CHPSCLA',true);;
  end;

end ;

procedure TOM_BSVPLANCLA.OnChangeField ( F: TField ) ;
begin
  Inherited ;
  if F.FieldName = 'BP0_TYPECLA' then
  begin
    if GetField('BP0_TYPECLA') = '000' then
    begin
      SetControlVisible('BP0_FIXE',true);;
      SetControlVisible('BP0_CHPSCLA',false);;
      SetControlVisible('BP0_CHPSLIB',false);;
      SetControlVisible('TBP0_FIXE',true);;
      SetControlVisible('TBP0_CHPSCLA',false);;
    end else
    begin
      SetControlVisible('BP0_FIXE',false);;
      SetControlVisible('BP0_CHPSCLA',true);;
      SetControlVisible('BP0_CHPSLIB',true);;
      SetControlVisible('TBP0_FIXE',false);;
      SetControlVisible('TBP0_CHPSCLA',true);;
    end;
  end;
end ;

procedure TOM_BSVPLANCLA.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_BSVPLANCLA.OnAfterRollBack;
begin
  Inherited ;
end ;

function TOM_BSVPLANCLA.LastNIV: Integer;
var QQ : TQuery;
begin
  Result := 0;
  QQ := OpenSQL('SELECT MAX(BP0_NIV) FROM BSVPLANCLA',True,1,'',true);
  if Not QQ.eof then
  begin
    Result := QQ.fields[0].AsInteger;
  end;
  ferme (QQ);
end;

Initialization
  registerclasses ( [ TOM_BSVPLANCLA ] ) ; 
end.
